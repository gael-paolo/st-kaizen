"""Portal Kaizen — un solo archivo.

Google Sheet (mismo archivo del plan), pestañas:
  plan_kaizen   14 columnas (id, fecha_hallazgo, sede, origen, hallazgo, ocurrencia, consecuencia, valor,
                prioridad, accion, tipo_accion, responsable, sucursal, deadline)
  seguimiento   las 6 columnas de siempre
  catalogos     categoria | valor              -> listas cerradas (el orden de las filas es el de la lista)
  accesos       usuario | clave | email        -> claves ('admin', 'usuario' compartida o una por persona) y correo
Alertas por correo (Gmail): ver correo.py y [gmail] en secrets.
Sin credenciales (sin secrets) funciona en modo local leyendo plan_kaizen.xlsx, sin guardar nada.
"""
import hmac
import os
import re
import threading
import time
from datetime import date, datetime, timedelta
import unicodedata
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import correo

st.set_page_config(page_title="Portal Kaizen", page_icon="🏭", layout="wide")

# ============================================================ constantes
H_PLAN = ["id", "fecha_hallazgo", "sede", "origen", "hallazgo", "ocurrencia", "consecuencia", "valor",
          "prioridad", "accion", "tipo_accion", "responsable", "sucursal", "deadline"]
H_SEG = ["id", "seguimiento", "fecha_seguimiento", "porc_avance", "estado", "evidencia"]
HOJAS = ["plan_kaizen", "seguimiento", "catalogos", "accesos"]
REQ = {"plan_kaizen": H_PLAN, "seguimiento": H_SEG, "catalogos": ["categoria", "valor"],
       "accesos": ["usuario", "clave"]}
# categorías cerradas -> etiqueta (en el orden del formulario)
CATS = {"sede": "Sede", "sucursal": "Sucursal", "origen": "Origen", "prioridad": "Prioridad",
        "tipo_accion": "Tipo de acción", "responsable": "Responsable"}
OTRO_ORIGEN = "➕ Otro (escribir uno nuevo)"
TZ = "America/La_Paz"
TTL = 60                    # segundos que se comparte una lectura de Sheets entre todos
DIAS_RIESGO = 7
MAX_TXT = 1500
ICONO = {"Atrasado": "🔴", "En riesgo": "🟡", "A tiempo": "🟢", "Sin fecha": "⚪", "Cerrado": "✅"}
ESTADOS = ["No iniciado", "En proceso", "Cerrado"]
PLAZOS = ["Atrasado", "En riesgo", "A tiempo", "Sin fecha", "Cerrado"]
RUTA_LOCAL = os.environ.get("KAIZEN_FILE", os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                                         "plan_kaizen.xlsx"))


@st.cache_resource
def compartido():
    """Estado común a todas las sesiones del proceso. Streamlit re-ejecuta este archivo en cada
    interacción, así que un `Lock()` o un dict global se recrearía cada vez: deben vivir aquí."""
    return {"lock": threading.Lock(), "fallos": {}}       # fallos: usuario -> instantes de intentos fallidos


def hoy():
    return pd.Timestamp(datetime.now(ZoneInfo(TZ)).date())


# ============================================================ lectura/escritura de datos
class Faltan(Exception):
    """Faltan pestañas en el Sheet. `datos` trae las que sí existen."""
    def __init__(self, faltan, datos):
        super().__init__(", ".join(faltan))
        self.faltan, self.datos = faltan, datos


def a_df(values, hoja):
    if not values:
        raise ValueError(f"La pestaña '{hoja}' está vacía (falta la fila de encabezados).")
    cab = [str(h).strip().lower() for h in values[0]]
    if hoja == "accesos":       # acepta los encabezados que se usan de forma natural
        alias = {"rol": "usuario", "contraseña": "clave", "contrasena": "clave", "password": "clave"}
        cab = [alias.get(c, c) if alias.get(c, c) not in cab else c for c in cab]
    falta = [c for c in REQ[hoja] if c not in cab]
    if falta:
        raise ValueError(f"A la pestaña '{hoja}' le faltan las columnas: {', '.join(falta)}.")
    filas = [(list(r) + [""] * len(cab))[:len(cab)] for r in values[1:]]
    df = pd.DataFrame(filas, columns=cab, dtype=object)
    df = df.loc[:, ~df.columns.duplicated()]
    vacia = df.apply(lambda c: c.map(lambda x: str(x).strip() == "")).all(axis=1)
    return df[~vacia].reset_index(drop=True)


def limpiar(v):
    """Valor seguro para escribir en Sheets (sin inyección de fórmulas)."""
    if v is None:
        return ""
    if isinstance(v, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(v).strftime("%Y-%m-%d")
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    return "'" + s if s[:1] in "=+-@" else s


class Sheets:
    es_local = False   # (no usar isinstance: Streamlit redefine las clases en cada ejecución)

    def __init__(self, sh):
        self.sh = sh

    def _leer(self, hojas):
        res = self.sh.values_batch_get(hojas, params={"valueRenderOption": "UNFORMATTED_VALUE"})
        return {h: a_df(vr.get("values", []), h) for h, vr in zip(hojas, res["valueRanges"])}

    def leer(self):
        try:
            return self._leer(HOJAS)
        except ValueError:
            raise
        except Exception:
            titulos = {w.title for w in self.sh.worksheets()}
            faltan = [h for h in HOJAS if h not in titulos]
            if not faltan:
                raise
            raise Faltan(faltan, self._leer([h for h in HOJAS if h in titulos]))

    def agregar(self, hoja, fila):
        """Agrega UNA fila al final (nunca reescribe la hoja). Devuelve el id si es del plan."""
        with compartido()["lock"]:
            ws = self.sh.worksheet(hoja)
            cab = [str(h).strip().lower() for h in ws.row_values(1)]
            if hoja == "plan_kaizen":   # id nuevo calculado con la hoja al día y bajo candado
                ids = pd.to_numeric(pd.Series(ws.col_values(cab.index("id") + 1)[1:]), errors="coerce")
                fila["id"] = int(ids.max()) + 1 if ids.notna().any() else 1
            ws.append_row([limpiar(fila.get(c, "")) for c in cab], value_input_option="USER_ENTERED",
                          table_range="A1")
            return fila.get("id")


class Local:
    """Modo local: lee plan_kaizen.xlsx; lo que se registra vive solo en memoria."""
    es_local = True

    def __init__(self, ruta):
        if not os.path.exists(ruta):
            raise ValueError(f"No encuentro '{ruta}'. En local: descarga el Sheet como Excel (Archivo > Descargar) y "
                             "guárdalo como plan_kaizen.xlsx. En la nube: faltan los secrets de la cuenta de servicio "
                             "y del Sheet (ver README).")
        crudo = pd.read_excel(ruta, sheet_name=None, header=None, dtype=object)
        crudo = {k.strip().lower(): v for k, v in crudo.items()}

        def matriz(df):
            f = lambda x: ("" if (not isinstance(x, str) and pd.isna(x)) else  # noqa: E731
                           x.strftime("%Y-%m-%d") if isinstance(x, (datetime, date)) else x)
            return [[f(x) for x in fila] for fila in df.values.tolist()]

        self.d = {h: a_df(matriz(crudo[h]), h) for h in HOJAS if h in crudo}
        if "plan_kaizen" not in self.d:
            raise ValueError("El Excel debe tener una pestaña llamada 'plan_kaizen'.")
        self.d.setdefault("seguimiento", pd.DataFrame(columns=H_SEG))
        if "catalogos" not in self.d:
            self.d["catalogos"] = derivar_catalogos(self.d["plan_kaizen"])

    def leer(self):
        return {k: v.copy() for k, v in self.d.items()}

    def agregar(self, hoja, fila):
        with compartido()["lock"]:
            df = self.d[hoja]
            if hoja == "plan_kaizen":
                ids = pd.to_numeric(df["id"], errors="coerce")
                fila["id"] = int(ids.max()) + 1 if ids.notna().any() else 1
            nueva = {c: limpiar(fila.get(c, "")) for c in df.columns}
            self.d[hoja] = pd.concat([df, pd.DataFrame([nueva])], ignore_index=True)
            return fila.get("id")


def leer_credenciales(sec):
    """Cuenta de servicio e ID del Sheet desde los secrets. Devuelve (info, spreadsheet) o (None, None).

    Acepta la cuenta en `[connections.gsheets]` o en `[gcp_service_account]`, y el Sheet como
    `spreadsheet` (URL o ID) o `spreadsheet_id`, ya sea suelto o dentro de la sección de la cuenta.
    """
    try:
        info = dict(sec["connections"]["gsheets"])
    except Exception:
        try:
            info = dict(sec["gcp_service_account"])
        except Exception:
            return None, None
    hoja = None
    for k in ("spreadsheet", "spreadsheet_id"):
        try:
            hoja = hoja or sec[k]
        except Exception:
            pass
        hoja = hoja or info.get(k)
    for k in ("spreadsheet", "spreadsheet_id"):
        info.pop(k, None)
    if not info.get("private_key") or not info.get("client_email") or not hoja:
        return None, None
    info["private_key"] = info["private_key"].replace("\\n", "\n")
    return info, str(hoja).strip()


@st.cache_resource
def tienda():
    info, hoja = leer_credenciales(st.secrets) if _hay_secrets() else (None, None)
    if not info:                # sin cuenta de servicio -> modo local (plan_kaizen.xlsx)
        return Local(RUTA_LOCAL)
    import gspread
    gc = gspread.service_account_from_dict(info)
    return Sheets(gc.open_by_url(hoja) if hoja.startswith("http") else gc.open_by_key(hoja))


def _hay_secrets():
    try:
        len(st.secrets)
        return True
    except Exception:           # no existe secrets.toml
        return False


def derivar_catalogos(plan):
    """Catálogo con TODOS los valores que ya existen en el plan, tal cual (se depura en el Sheet)."""
    uni = lambda s: [v for v in dict.fromkeys(" ".join(str(x).split()) for x in s) if v]  # noqa: E731
    filas = []
    for cat in ("sede", "sucursal", "origen"):
        filas += [(cat, v, "") for v in sorted(uni(plan[cat]), key=str.lower)]
    base = {"prioridad": ["Baja", "Media", "Alta", "Crítica"],
            "tipo_accion": ["Correctiva", "Mejora", "Preventiva"]}
    for cat, valores in base.items():
        extra = [v for v in uni(plan[cat]) if v not in valores]
        filas += [(cat, v, "") for v in valores + extra]
    filas += [("responsable", v, "") for v in sorted({v.lower() for v in uni(plan["responsable"])})]
    return pd.DataFrame(filas, columns=["categoria", "valor", "padre"])


# ============================================================ limpieza y reglas de negocio
def parse_fecha(s):
    """Serial de Sheets, ISO o dd/mm/aaaa -> fecha (NaT si no se entiende)."""
    s = pd.Series(s, dtype=object)
    num = pd.to_numeric(s, errors="coerce")
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    es_num = num.notna() & (num > 20000) & (num < 80000)
    if es_num.any():
        out[es_num] = pd.to_datetime(num[es_num], unit="D", origin="1899-12-30")
    resto = ~es_num & s.map(lambda x: str(x).strip() != "").astype(bool)
    if resto.any():
        t = s[resto].astype(str).str.strip()
        out[resto] = pd.to_datetime(t, format="%Y-%m-%d", errors="coerce").fillna(
            pd.to_datetime(t, dayfirst=True, errors="coerce", format="mixed"))
    return out.dt.normalize()


def parse_avance(s):
    """'50%', '50,5', 50 o 0.5 (celda con formato %) -> 0..100."""
    s = pd.Series(s, dtype=object)
    con_pct = s.astype(str).str.contains("%", regex=False)
    num = pd.to_numeric(s.astype(str).str.replace("%", "", regex=False).str.replace(",", ".", regex=False)
                        .str.strip(), errors="coerce")
    v = num.dropna()
    if len(v) and not con_pct.any() and (v <= 1).all() and v.between(0.0001, 0.9999).any():
        num = num * 100
    return num.fillna(0).clip(0, 100).astype(float)


def txt(s):
    return s.fillna("").astype(str).map(lambda x: " ".join(x.split()))


def preparar(d):
    """Devuelve (plan, seg, catalogos, claves, contactos, advertencias) ya limpios."""
    adv = []
    plan = d["plan_kaizen"].copy()
    for c in plan.columns:
        plan[c] = txt(plan[c])
    plan["id"] = pd.to_numeric(plan["id"], errors="coerce")
    if plan["id"].isna().any():
        adv.append(f"{int(plan['id'].isna().sum())} fila(s) del plan sin ID válido fueron ignoradas.")
    plan = plan[plan["id"].notna()].copy()
    plan["id"] = plan["id"].astype("int64")
    dup = plan["id"].duplicated()
    if dup.any():
        adv.append("IDs repetidos en el plan (se usa la primera fila): "
                   + ", ".join(map(str, sorted(plan.loc[dup, "id"].unique())[:10])))
        plan = plan[~dup]
    plan["fecha_hallazgo"] = parse_fecha(plan["fecha_hallazgo"])
    plan["deadline"] = parse_fecha(plan["deadline"])
    for c in ("ocurrencia", "consecuencia"):
        plan[c] = pd.to_numeric(plan[c], errors="coerce").fillna(0).astype("int64")
    plan["valor"] = plan["ocurrencia"] * plan["consecuencia"]
    plan = plan.reset_index(drop=True)

    seg = d["seguimiento"].copy()
    for c in seg.columns:
        seg[c] = txt(seg[c])
    seg["id"] = pd.to_numeric(seg["id"], errors="coerce")
    seg = seg[seg["id"].notna()].copy()
    seg["id"] = seg["id"].astype("int64")
    huerfano = ~seg["id"].isin(plan["id"])
    if huerfano.any():
        adv.append(f"{int(huerfano.sum())} seguimiento(s) apuntan a una acción que no existe.")
    seg = seg[~huerfano].copy()
    seg["fecha_seguimiento"] = parse_fecha(seg["fecha_seguimiento"])
    seg["porc_avance"] = parse_avance(seg["porc_avance"])
    seg["_n"] = range(len(seg))     # orden en la hoja, para desempatar fechas iguales

    cats = d["catalogos"].copy()
    if "padre" not in cats.columns:
        cats["padre"] = ""
    for c in ("categoria", "valor", "padre"):
        cats[c] = txt(cats[c])
    cats["categoria"] = cats["categoria"].str.lower()
    cats = cats[(cats["categoria"] != "") & (cats["valor"] != "")].reset_index(drop=True)

    claves, contactos = {}, {}     # claves: usuario -> clave ('admin', 'usuario' = compartida, o una por persona)
    if "accesos" in d:
        ac = d["accesos"]
        for i, u in enumerate(ac["usuario"]):
            u = str(u).strip().lower()
            if not u:
                continue
            c = str(ac["clave"].iloc[i]).strip()
            if c and c != "nan":
                claves.setdefault(u, c)
            mail = "".join(str(ac["email"].iloc[i]).split()) if "email" in ac.columns else ""   # sin espacios raros
            nom = str(ac["nombre"].iloc[i]).strip() if "nombre" in ac.columns else ""
            contactos.setdefault(u, {"email": mail if "@" in mail else "", "nombre": "" if nom == "nan" else nom})
    return plan, seg, cats, claves, contactos, adv


def enriquecer(plan, seg):
    """Estado, avance y semáforo por acción (vectorizado; se calcula una vez por recarga)."""
    ult = seg.sort_values(["fecha_seguimiento", "_n"], kind="stable").groupby("id").tail(1)
    ult = ult[["id", "porc_avance", "fecha_seguimiento"]].rename(columns={"fecha_seguimiento": "fecha_ult"})
    df = plan.merge(ult, on="id", how="left")
    df["n_seg"] = df["id"].map(seg.groupby("id").size()).fillna(0).astype(int)
    df["porc_avance"] = df["porc_avance"].fillna(0.0)
    df["estado"] = np.select([df["n_seg"] == 0, df["porc_avance"] >= 100], ["No iniciado", "Cerrado"],
                             default="En proceso")
    dias = (df["deadline"] - hoy()).dt.days
    df["dias"] = dias
    df["plazo"] = np.select([df["estado"] == "Cerrado", dias.isna(), dias < 0, dias <= DIAS_RIESGO],
                            ["Cerrado", "Sin fecha", "Atrasado", "En riesgo"], default="A tiempo")
    df["plazo_txt"] = df["plazo"].map(lambda x: f"{ICONO[x]} {x}")
    return df


@st.cache_data(ttl=TTL, show_spinner="Cargando datos…")
def cargar():
    """Una sola lectura de Sheets sirve a todos los usuarios durante TTL segundos."""
    plan, seg, cats, claves, contactos, adv = preparar(tienda().leer())
    return dict(plan=plan, seg=seg, cats=cats, claves=claves, contactos=contactos, adv=adv,
                acc=enriquecer(plan, seg),
                hora=datetime.now(ZoneInfo(TZ)).strftime("%H:%M:%S"))


def opciones(cats, cat, padre=None):
    c = cats[cats["categoria"] == cat]
    if padre is not None:
        c = c[(c["padre"].str.lower() == padre.lower()) | (c["padre"] == "")]
    return list(dict.fromkeys(c["valor"]))      # el orden es el de las filas del Sheet


# ============================================================ escritura con validación
def clave(v):
    """Para comparar textos: sin tildes, mayúsculas, espacios ni signos ('5S ' == '5s')."""
    t = "".join(c for c in unicodedata.normalize("NFD", str(v)) if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]+", "", t.lower())


def canon_origen(valor, cats):
    """Un origen escrito a mano se alinea con el del catálogo si solo difiere en tildes/mayúsculas/signos."""
    v = " ".join(str(valor).split())
    for o in opciones(cats, "origen"):
        if clave(o) == clave(v):
            return o
    return v


def origenes_nuevos(D):
    """Orígenes del plan que no están en el catálogo -> Series (valor -> n.º de acciones)."""
    ok = set(opciones(D["cats"], "origen"))
    o = D["acc"]["origen"]
    return o[(o != "") & ~o.isin(ok)].value_counts()


def validar_alta(d, D):
    err = [f"{CATS[k]} es obligatorio." for k in ("sede", "origen", "prioridad", "tipo_accion", "responsable")
           if not d[k]]
    err += [f"{x} es obligatorio." for x, k in (("Hallazgo", "hallazgo"), ("Acción propuesta", "accion"))
            if not d[k].strip()]
    for k in ("sede", "sucursal", "prioridad", "tipo_accion", "responsable"):   # listas cerradas
        ops = opciones(D["cats"], k)
        if d[k] and d[k] not in ops:
            err.append(f"'{d[k]}' ya no es una opción válida de {CATS[k]}.")
        elif ops and not d[k] and k == "sucursal":
            err.append("Sucursal es obligatorio.")
    if len(d["origen"]) > 80:                                                  # origen: se puede escribir uno nuevo
        err.append("El origen admite hasta 80 caracteres.")
    if pd.Timestamp(d["fecha_hallazgo"]) > hoy():
        err.append("La fecha del hallazgo no puede ser futura.")
    if pd.Timestamp(d["deadline"]) < pd.Timestamp(d["fecha_hallazgo"]):
        err.append("El deadline no puede ser anterior a la fecha del hallazgo.")
    if len(d["hallazgo"]) > MAX_TXT or len(d["accion"]) > MAX_TXT:
        err.append(f"Hallazgo y acción admiten hasta {MAX_TXT} caracteres.")
    if not err:
        p = D["plan"]
        igual = p[(p["hallazgo"] == " ".join(d["hallazgo"].split())) & (p["responsable"] == d["responsable"])
                  & (p["accion"] == " ".join(d["accion"].split()))]
        if len(igual):
            err.append(f"Ya existe la acción #{igual.iloc[0]['id']} con el mismo hallazgo, acción y responsable.")
    return err


def guardar_alta(d, D):
    d = {**d, "origen": canon_origen(d["origen"], D["cats"])}
    err = validar_alta(d, D)
    if err:
        return None, err
    fila = {**d, "valor": d["ocurrencia"] * d["consecuencia"]}
    try:
        nuevo = tienda().agregar("plan_kaizen", fila)
    except Exception as e:
        return None, [f"No se pudo guardar en Google Sheets: {e}"]
    cargar.clear()
    return nuevo, []


def puede_seguir(rol, usuario, a):
    return rol == "admin" or a["responsable"] == usuario


def guardar_seg(rol, usuario, D, id_accion, texto, avance, evidencia):
    """La fecha del seguimiento NO la elige nadie: siempre es la de hoy (hora de Bolivia)."""
    fecha = hoy()
    f = D["acc"][D["acc"]["id"] == id_accion]
    if f.empty:
        return ["La acción no existe."]
    a = f.iloc[0]
    err = []
    if not puede_seguir(rol, usuario, a):
        return ["No tienes permiso para registrar seguimiento en esta acción."]
    if a["estado"] == "Cerrado" and rol != "admin":
        err.append("La acción está cerrada; solo el administrador puede reabrirla.")
    if not texto.strip():
        err.append("Describe el avance realizado.")
    if len(texto) > MAX_TXT or len(evidencia) > MAX_TXT:
        err.append(f"El texto admite hasta {MAX_TXT} caracteres.")
    if avance < a["porc_avance"] and rol != "admin":
        err.append(f"El avance no puede bajar del {a['porc_avance']:.0f}% ya registrado.")
    s = D["seg"]
    if len(s[(s["id"] == id_accion) & (s["seguimiento"] == " ".join(texto.split()))
             & (s["porc_avance"] == avance) & (s["fecha_seguimiento"] == fecha)]):
        err.append("Este seguimiento ya fue registrado.")
    if err:
        return err
    try:
        tienda().agregar("seguimiento", {"id": int(id_accion), "seguimiento": texto.strip(),
                                         "fecha_seguimiento": fecha, "porc_avance": float(avance),
                                         "estado": "Finalizado" if avance >= 100 else "Proceso",
                                         "evidencia": evidencia.strip()})
    except Exception as e:
        return [f"No se pudo guardar en Google Sheets: {e}"]
    cargar.clear()
    return []


# ============================================================ acceso (una clave por rol)
def ingresar(D):
    """Pantalla de acceso. Devuelve (rol, usuario) o detiene la app."""
    if tienda().es_local and not D["claves"]:   # local sin pestaña 'accesos': eliges con qué rol mirar
        with st.sidebar:
            st.warning("MODO LOCAL: datos del archivo, nada se guarda.")
            quien = st.selectbox("Ver como", ["Administrador"] + opciones(D["cats"], "responsable"))
        return ("admin", "admin") if quien == "Administrador" else ("usuario", quien)
    if "sesion" in st.session_state:
        return st.session_state["sesion"]
    _, centro, _ = st.columns([1, 1.2, 1])
    with centro:
        st.markdown("<h1 style='text-align:center'>🏭 Portal Kaizen</h1>", unsafe_allow_html=True)
        with st.form("acceso"):
            quien = st.text_input("Usuario", autocomplete="username", placeholder="Escribe tu usuario")
            clave = st.text_input("Contraseña", type="password", autocomplete="current-password")
            ok = st.form_submit_button("Ingresar", type="primary", width="stretch")
        if ok:
            u = quien.strip().lower()
            resp = {r.lower(): r for r in opciones(D["cats"], "responsable")}     # la lista NO se muestra
            rol = "admin" if u == "admin" else ("usuario" if u in resp else None)
            propia = (D["claves"].get(u) if rol == "usuario" and u != "usuario" else None)
            FALLOS = compartido()["fallos"]
            fallos = [t for t in FALLOS.get(u, []) if time.time() - t < 600]
            # admin: solo la fila 'admin'. Responsable: su clave propia o, si no tiene, la compartida 'usuario'.
            real = (D["claves"].get("admin", "") if rol == "admin"
                    else (propia or D["claves"].get("usuario", "")) if rol == "usuario" else "")
            if len(fallos) >= 5:
                st.error("Demasiados intentos. Espera 10 minutos.")
            elif real and clave and hmac.compare_digest(clave.encode(), real.encode()):
                FALLOS.pop(u, None)
                st.session_state["sesion"] = (rol, "admin" if rol == "admin" else resp[u])
                st.rerun()
            else:                       # mismo mensaje exista o no el usuario (no revela quién está registrado)
                FALLOS[u] = fallos + [time.time()]
                st.error("Usuario o contraseña incorrectos.")
    st.stop()


# ============================================================ componentes de pantalla
COL_TABLA = {
    "id": st.column_config.NumberColumn("ID", width="small", format="%d"),
    "plazo_txt": st.column_config.TextColumn("Plazo", width="medium"),
    "hallazgo": st.column_config.TextColumn("Hallazgo", width="large"),
    "responsable": st.column_config.TextColumn("Responsable"),
    "prioridad": st.column_config.TextColumn("Prioridad", width="small"),
    "estado": st.column_config.TextColumn("Estado", width="small"),
    "porc_avance": st.column_config.ProgressColumn("Avance", min_value=0, max_value=100, format="%d%%",
                                                   width="small"),
    "deadline": st.column_config.DateColumn("Deadline", format="DD/MM/YYYY", width="small"),
    "dias": st.column_config.NumberColumn("Días", format="%d", width="small",
                                          help="Días que faltan (negativo = atraso)"),
    "sede": st.column_config.TextColumn("Sede", width="small"),
    "origen": st.column_config.TextColumn("Origen"),
}


def avisar(msg):
    st.session_state["_aviso"] = msg


def mostrar_aviso():
    if "_aviso" in st.session_state:
        st.toast(st.session_state.pop("_aviso"), icon="✅", duration="long")


def kpis(df, promedio=False):
    n, cerr = len(df), int((df["estado"] == "Cerrado").sum())
    items = [("Total", n, None), ("En proceso", int((df["estado"] == "En proceso").sum()), None),
             ("No iniciadas", int((df["estado"] == "No iniciado").sum()), None),
             ("Cerradas", cerr, f"{cerr / n * 100:.0f}%" if n else None),
             ("Atrasadas", int((df["plazo"] == "Atrasado").sum()), None),
             ("En riesgo", int((df["plazo"] == "En riesgo").sum()), None)]
    if promedio:
        items.append(("Avance prom.", f"{df['porc_avance'].mean():.0f}%" if n else "—", None))
    for col, (t, v, d) in zip(st.columns(len(items)), items):
        col.metric(t, v, delta=d, delta_color="off", delta_arrow="off", border=True)


def tabla(df, key, cols, alto=340):
    """Tabla con selección de fila; devuelve el id elegido o None."""
    vista = df[cols].reset_index(drop=True)
    ev = st.dataframe(vista, hide_index=True, width="stretch", height=alto, key=key, on_select="rerun",
                      selection_mode="single-row", column_config={c: COL_TABLA[c] for c in cols})
    filas = ev.selection.rows if ev and ev.selection else []
    return int(vista.iloc[filas[0]]["id"]) if filas else None


def panel(D, rol, usuario, id_accion):
    f = D["acc"][D["acc"]["id"] == id_accion]
    if f.empty:
        return
    a = f.iloc[0]
    with st.container(border=True):
        c1, c2 = st.columns([4, 1])
        c1.markdown(f"#### #{a['id']} · {a['hallazgo'][:110]}")
        c1.markdown(" ".join(f":gray-background[{x}]" for x in [a["sede"], a["origen"],
                                                               f"Prioridad {a['prioridad']}"] if x))
        c2.markdown(f"**{a['plazo_txt']}**")
        c2.caption(a["estado"])
        st.progress(int(a["porc_avance"]) / 100, text=f"Avance {a['porc_avance']:.0f}%")
        nombres = ["📌 Detalle", "📝 Registrar seguimiento", "📜 Historial"]
        abrir = st.session_state.get("seg_id") == id_accion
        t1, t2, t3 = st.tabs(nombres, default=nombres[1] if abrir else None)
        with t1:
            x, y = st.columns(2)
            x.markdown("**Hallazgo**")
            x.write(a["hallazgo"])
            x.markdown("**Acción propuesta**")
            x.write(a["accion"] or "—")
            fmt = lambda t: t.strftime("%d/%m/%Y") if pd.notna(t) else "—"  # noqa: E731
            y.markdown(f"**Responsable:** {a['responsable']}")
            y.markdown(f"**Tipo:** {a['tipo_accion']} · **Sucursal:** {a['sucursal'] or '—'}")
            y.markdown(f"**Hallazgo:** {fmt(a['fecha_hallazgo'])} · **Deadline:** {fmt(a['deadline'])}")
            y.markdown(f"**Riesgo:** {a['valor']} (ocurrencia {a['ocurrencia']} × consecuencia {a['consecuencia']})")
        with t2:
            form_seg(D, rol, usuario, a)
        with t3:
            h = D["seg"][D["seg"]["id"] == id_accion].sort_values(["fecha_seguimiento", "_n"], ascending=False)
            if h.empty:
                st.info("Aún no hay seguimientos registrados.")
            else:
                st.dataframe(h[["fecha_seguimiento", "porc_avance", "seguimiento", "evidencia"]],
                             hide_index=True, width="stretch", column_config={
                                 "fecha_seguimiento": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
                                 "porc_avance": st.column_config.ProgressColumn("Avance", min_value=0,
                                                                                max_value=100, format="%d%%"),
                                 "seguimiento": st.column_config.TextColumn("Seguimiento", width="large"),
                                 "evidencia": st.column_config.TextColumn("Evidencia")})


def form_seg(D, rol, usuario, a):
    if not puede_seguir(rol, usuario, a):
        st.info("Solo el responsable o el administrador puede registrar seguimientos.")
        return
    if a["estado"] == "Cerrado":
        if rol != "admin":
            st.success("🎉 Esta acción está cerrada.")
            return
        st.warning("Acción cerrada: un seguimiento con menos de 100% la reabrirá.")
    actual = int(a["porc_avance"])
    with st.form(f"seg_{a['id']}", clear_on_submit=True):
        texto = st.text_area("¿Qué avanzaste? *", height=100, max_chars=MAX_TXT,
                             placeholder="Describe el avance realizado…")
        avance = st.slider("% de avance", 0 if rol == "admin" else actual, 100, actual,
                           step=5 if actual % 5 == 0 else 1, help="Al llegar a 100% la acción se cierra.")
        evidencia = st.text_input("Evidencia (enlace o referencia, opcional)", max_chars=500)
        st.caption(f"📅 Se registra con la fecha de hoy ({hoy():%d/%m/%Y}).")
        enviar = st.form_submit_button("✅ Registrar avance", type="primary")
    if enviar:
        err = guardar_seg(rol, usuario, D, int(a["id"]), texto, avance, evidencia)
        if err:
            for e in err:
                st.error(e)
        else:
            st.session_state["seg_id"] = int(a["id"])
            avisar(f"Seguimiento registrado en #{a['id']} · {avance}%" + (" · acción CERRADA 🎉" if avance >= 100 else ""))
            st.rerun()


@st.fragment
def form_alta(D, rol, usuario):
    cats = D["cats"]
    vacias = [CATS[c] for c in ("sede", "prioridad", "tipo_accion", "responsable") if not opciones(cats, c)]
    if vacias:
        st.error("Faltan opciones en la pestaña 'catalogos' para: " + ", ".join(vacias))
        return
    pick = lambda etiqueta, cat, key, ops=None: st.selectbox(  # noqa: E731
        etiqueta, ops if ops is not None else opciones(cats, cat), index=None, placeholder="Elige…", key=key)
    # (los widgets de abajo usan el mismo `pick`; el origen añade la opción de escribir uno nuevo)
    st.caption("Los campos con * son obligatorios. Las listas se administran en la pestaña 'catalogos' del Sheet.")
    st.markdown("**1 · ¿Dónde y por qué?**")
    c1, c2, c3 = st.columns(3)
    with c1:
        sede = pick("Sede *", "sede", "a_sede")
    with c2:
        suc_ops = opciones(cats, "sucursal")
        sucursal = pick("Sucursal" + (" *" if suc_ops else ""), "sucursal", "a_suc") if suc_ops else ""
    with c3:
        sel = pick("Origen *", "origen", "a_ori", ops=opciones(cats, "origen") + [OTRO_ORIGEN])
        origen = (st.text_input("Escribe el origen *", max_chars=80, key="a_ori_txt",
                                help="Si no está en la lista. El administrador revisará los orígenes nuevos.")
                  if sel == OTRO_ORIGEN else (sel or ""))
    st.markdown("**2 · El hallazgo**")
    fh = st.date_input("Fecha del hallazgo *", hoy().date(), max_value=hoy().date(), format="DD/MM/YYYY", key="a_fh")
    hallazgo = st.text_area("Hallazgo *", height=90, max_chars=MAX_TXT, placeholder="¿Qué encontraste?", key="a_hal")
    c1, c2, c3 = st.columns(3)
    oc = c1.select_slider("Ocurrencia (1 rara – 5 frecuente)", [1, 2, 3, 4, 5], 3, key="a_oc")
    co = c2.select_slider("Consecuencia (1 leve – 5 grave)", [1, 2, 3, 4, 5], 3, key="a_co")
    c3.metric("Valor de riesgo", oc * co, help="Ocurrencia × consecuencia (máx. 25)")
    st.markdown("**3 · La acción**")
    accion = st.text_area("Acción propuesta *", height=90, max_chars=MAX_TXT,
                          placeholder="¿Qué se hará para resolverlo?", key="a_acc")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        tipo = pick("Tipo de acción *", "tipo_accion", "a_tipo")
    with c2:
        prio = pick("Prioridad *", "prioridad", "a_pri")
    resp_ops = opciones(cats, "responsable")
    with c3:
        resp = st.selectbox("Responsable *", resp_ops, key="a_resp", placeholder="Elige…",
                            index=resp_ops.index(usuario) if rol != "admin" and usuario in resp_ops else None)
    dl = c4.date_input("Deadline *", fh + timedelta(days=14), min_value=fh, format="DD/MM/YYYY", key="a_dl")
    if st.button("✅ Registrar acción", type="primary", key="a_ok"):
        d = dict(fecha_hallazgo=fh, sede=sede or "", sucursal=sucursal or "", origen=origen or "",
                 hallazgo=hallazgo, ocurrencia=oc, consecuencia=co, prioridad=prio or "", accion=accion,
                 tipo_accion=tipo or "", responsable=resp or "", deadline=dl)
        nuevo, err = guardar_alta(d, D)
        if err:
            for e in err:
                st.error(e)
        else:
            for k in [k for k in st.session_state if k.startswith("a_")]:
                del st.session_state[k]
            nuevo_o = " · origen nuevo: el administrador lo revisará" if sel == OTRO_ORIGEN else ""
            avisar(f"Acción #{nuevo} registrada para {resp}.{nuevo_o}")
            st.session_state["foco"] = nuevo
            st.session_state["ir"] = True
            st.rerun()


# ============================================================ vista del responsable
LISTA = {"usuario": "📋 Mis acciones", "admin": "📋 Plan y seguimiento"}
RAPIDOS = {"Pendientes": lambda d: d[d["estado"] != "Cerrado"], "Atrasadas": lambda d: d[d["plazo"] == "Atrasado"],
           "En riesgo": lambda d: d[d["plazo"] == "En riesgo"], "Cerradas": lambda d: d[d["estado"] == "Cerrado"],
           "Todas": lambda d: d}


def orden_urgencia(df):
    rank = df["plazo"].map({p: i for i, p in enumerate(PLAZOS)})
    return df.assign(_r=rank).sort_values(["_r", "deadline", "id"], na_position="last").drop(columns="_r")


def vista_usuario(D, usuario):
    nav = ["🏠 Mi resumen", LISTA["usuario"], "➕ Nueva acción"]
    pagina = st.sidebar.radio("Menú", nav, key="nav_usuario", label_visibility="collapsed")
    mias = D["acc"][D["acc"]["responsable"] == usuario]
    if pagina == nav[0]:
        st.header(f"Hola, {usuario} 👋")
        if mias.empty:
            st.info("Aún no tienes acciones asignadas. Puedes registrar una en **➕ Nueva acción**.")
            return
        kpis(mias)
        st.subheader("⏰ Requieren tu atención")
        urg = orden_urgencia(mias[mias["plazo"].isin(["Atrasado", "En riesgo"])])
        if urg.empty:
            st.success(f"Estás al día: nada atrasado ni por vencer en {DIAS_RIESGO} días.")
        for _, r in urg.head(8).iterrows():
            with st.container(border=True):
                c1, c2 = st.columns([5, 1.3], vertical_alignment="center")
                c1.markdown(f"**#{r['id']} · {r['hallazgo'][:90]}**  \n{r['plazo_txt']} · vence "
                            f"{r['deadline']:%d/%m/%Y} ({int(r['dias'])} d) · avance {r['porc_avance']:.0f}%")
                if c2.button("Dar seguimiento", key=f"go_{r['id']}", type="primary", width="stretch"):
                    st.session_state.update(foco=int(r["id"]), seg_id=int(r["id"]), ir=True)
                    st.rerun()
        if len(urg) > 8:
            st.caption(f"…y {len(urg) - 8} más en **Mis acciones**.")
    elif pagina == nav[1]:
        st.header("📋 Mis acciones")
        if mias.empty:
            st.info("Aún no tienes acciones asignadas.")
            return
        c1, c2 = st.columns([3, 2])
        rapido = c1.radio("Mostrar", list(RAPIDOS), horizontal=True, key="rapido")
        buscar = c2.text_input("Buscar", placeholder="ID o palabra", label_visibility="collapsed", key="buscar_u")
        df = RAPIDOS[rapido](mias)
        if buscar:
            b = buscar.strip()
            df = df[df["id"].astype(str).eq(b) | df["hallazgo"].str.contains(b, case=False, regex=False)
                    | df["accion"].str.contains(b, case=False, regex=False)]
        df = orden_urgencia(df)
        st.caption(f"{len(df)} de {len(mias)} acciones · elige una fila para ver el detalle y dar seguimiento")
        if df.empty:
            st.info("No hay acciones con ese filtro.")
            return
        elegido = tabla(df, "tabla_u", ["id", "plazo_txt", "hallazgo", "prioridad", "estado", "porc_avance",
                                        "deadline", "dias"], alto=300)
        foco = st.session_state.get("foco")
        sel = elegido or (foco if foco in set(mias["id"]) else None)
        if sel:
            panel(D, "usuario", usuario, sel)
        else:
            st.info("👆 Elige una acción de la tabla para ver su detalle.")
    else:
        st.header("➕ Nueva acción")
        form_alta(D, "usuario", usuario)


# ============================================================ vista Kaizen (administrador)
FILTROS = [("sede", "Sede"), ("origen", "Origen"), ("responsable", "Responsable"), ("prioridad", "Prioridad"),
           ("tipo_accion", "Tipo de acción"), ("estado", "Estado"), ("plazo", "Plazo")]
DIMENSIONES = {"Origen": "origen", "Sede": "sede", "Responsable": "responsable",
               "Prioridad": "prioridad", "Tipo de acción": "tipo_accion"}


def filtrar(acc):
    def limpiar_filtros():
        for k in [k for k in st.session_state if k.startswith("f_")]:
            del st.session_state[k]
    with st.sidebar:
        st.divider()
        st.markdown("**Filtros**")
        texto = st.text_input("Buscar", key="f_txt", placeholder="ID, hallazgo o acción")
        for col, etiqueta in FILTROS:
            st.multiselect(etiqueta, sorted(x for x in acc[col].unique() if x), key=f"f_{col}", placeholder="Todos")
        rango = st.date_input("Fecha del hallazgo", value=(), key="f_fechas", format="DD/MM/YYYY")
        st.button("Limpiar filtros", on_click=limpiar_filtros, width="stretch")
    df = acc
    for col, _ in FILTROS:
        if st.session_state.get(f"f_{col}"):
            df = df[df[col].isin(st.session_state[f"f_{col}"])]
    if texto:
        t = texto.strip()
        df = df[df["id"].astype(str).eq(t) | df["hallazgo"].str.contains(t, case=False, regex=False)
                | df["accion"].str.contains(t, case=False, regex=False)]
    if isinstance(rango, (tuple, list)) and len(rango) == 2:
        df = df[(df["fecha_hallazgo"] >= pd.Timestamp(rango[0])) & (df["fecha_hallazgo"] <= pd.Timestamp(rango[1]))]
    return df


def vista_admin(D):
    nav = ["📊 Panorama", LISTA["admin"], "➕ Nueva acción", "✉️ Alertas", "🛠️ Control"]
    pagina = st.sidebar.radio("Menú", nav, key="nav_admin", label_visibility="collapsed")
    aviso_origenes(D, pagina != nav[2])
    if pagina in (nav[0], nav[1]):
        df = filtrar(D["acc"])
    if pagina == nav[0]:
        panorama(df)
    elif pagina == nav[1]:
        st.header("📋 Plan y seguimiento")
        kpis(df, promedio=True)
        if df.empty:
            st.info("No hay acciones con los filtros actuales.")
            return
        df = orden_urgencia(df)
        c1, c2 = st.columns([4, 1.2], vertical_alignment="center")
        c1.caption(f"{len(df)} acciones · elige una fila para ver el detalle y registrar seguimiento")
        c2.download_button("⬇️ Descargar CSV", df.drop(columns=["plazo_txt"]).to_csv(index=False).encode("utf-8-sig"),
                           file_name="plan_kaizen.csv", width="stretch")
        elegido = tabla(df, "tabla_a", ["id", "plazo_txt", "sede", "origen", "hallazgo", "responsable",
                                        "prioridad", "estado", "porc_avance", "deadline", "dias"], alto=400)
        foco = st.session_state.get("foco")
        sel = elegido or (foco if foco in set(df["id"]) else None)
        if sel:
            panel(D, "admin", "admin", sel)
        else:
            st.info("👆 Elige una acción de la tabla para ver su detalle o registrar seguimiento.")
    elif pagina == nav[2]:
        st.header("➕ Nueva acción")
        form_alta(D, "admin", "admin")
    elif pagina == nav[3]:
        alertas(D)
    else:
        control(D)


def panorama(df):
    import plotly.express as px
    st.header("📊 Panorama de planes de acción")
    if df.empty:
        st.info("No hay acciones con los filtros actuales.")
        return
    kpis(df, promedio=True)
    cfg = {"displayModeBar": False}
    colores = {"No iniciado": "#9aa5b1", "En proceso": "#3b82f6", "Cerrado": "#22c55e"}
    c1, c2 = st.columns(2)
    n = df["estado"].value_counts().reindex(ESTADOS).dropna().reset_index()
    n.columns = ["estado", "n"]
    fig = px.pie(n, names="estado", values="n", hole=.55, color="estado", color_discrete_map=colores,
                 title="Estado de las acciones")
    fig.update_layout(height=300, margin=dict(t=50, b=10, l=10, r=10), legend_orientation="h")
    c1.plotly_chart(fig, width="stretch", config=cfg)
    p = df["plazo"].value_counts().reindex(PLAZOS).dropna().reset_index()
    p.columns = ["plazo", "n"]
    fig = px.bar(p, x="plazo", y="n", color="plazo", text="n", title="Cumplimiento de plazos",
                 color_discrete_map={"Atrasado": "#ef4444", "En riesgo": "#f59e0b", "A tiempo": "#22c55e",
                                     "Sin fecha": "#9aa5b1", "Cerrado": "#14b8a6"})
    fig.update_layout(height=300, margin=dict(t=50, b=10, l=10, r=10), showlegend=False, xaxis_title=None,
                      yaxis_title=None)
    c2.plotly_chart(fig, width="stretch", config=cfg)

    st.subheader("Avance por…")
    etiqueta = st.radio("Agrupar por", list(DIMENSIONES), horizontal=True, label_visibility="collapsed", key="dim")
    dim = DIMENSIONES[etiqueta]
    g = df.assign(**{dim: df[dim].replace("", "(sin dato)")})
    res = g.groupby(dim).agg(acciones=("id", "size"), cerradas=("estado", lambda s: (s == "Cerrado").sum()),
                             atrasadas=("plazo", lambda s: (s == "Atrasado").sum()),
                             avance=("porc_avance", "mean")).reset_index().sort_values("acciones", ascending=False)
    top = list(res.head(15)[dim])
    cnt = g[g[dim].isin(top)].groupby([dim, "estado"]).size().reset_index(name="n")
    fig = px.bar(cnt, y=dim, x="n", color="estado", orientation="h", color_discrete_map=colores,
                 category_orders={dim: top[::-1], "estado": ESTADOS})
    fig.update_layout(height=max(260, 34 * len(top) + 80), margin=dict(t=10, b=10, l=10, r=10), legend_title=None,
                      legend_orientation="h", xaxis_title=None, yaxis_title=None)
    st.plotly_chart(fig, width="stretch", config=cfg)
    st.dataframe(res.rename(columns={dim: etiqueta}), hide_index=True, width="stretch", column_config={
        "avance": st.column_config.ProgressColumn("Avance prom.", min_value=0, max_value=100, format="%d%%"),
        "acciones": st.column_config.NumberColumn("Acciones", format="%d"),
        "cerradas": st.column_config.NumberColumn("Cerradas", format="%d"),
        "atrasadas": st.column_config.NumberColumn("Atrasadas", format="%d")})


def aviso_origenes(D, mostrar=True):
    """Alerta al administrador: orígenes escritos a mano que no están en el catálogo."""
    n = origenes_nuevos(D)
    if mostrar and len(n):
        lista = ", ".join(f"'{o}' ({c})" for o, c in n.head(8).items()) + (" …" if len(n) > 8 else "")
        st.warning(f"⚠️ **{len(n)} origen(es) no estandarizado(s)** en {int(n.sum())} acción(es): {lista}. "
                   "Agrégalos a `catalogos` (categoria `origen`) para estandarizarlos o corrígelos en el plan.")


def config_gmail():
    """Credenciales de Gmail desde secrets ([gmail] user, app_password, nombre, cc, url_portal).

    Devuelve (config, motivo): config es None si no sirve y `motivo` explica por qué (sin mostrar valores).
    """
    try:
        g = dict(st.secrets["gmail"])
    except Exception:
        try:
            secciones = ", ".join(sorted(st.secrets.keys())) or "ninguna"
        except Exception:
            secciones = "ninguna"
        return None, f"No encuentro la sección [gmail] en los secrets (secciones que sí hay: {secciones})."
    faltan = [k for k in ("user", "app_password") if not str(g.get(k, "")).strip()]
    if faltan:
        return None, f"En [gmail] faltan: {', '.join(faltan)} (campos encontrados: {', '.join(g) or 'ninguno'})."
    return g, ""


def alertas(D):
    st.header("✉️ Alertas por correo")
    st.caption("Cada responsable recibe SU reporte: acciones pendientes, en proceso y resumen de cerradas. "
               "Tú eliges a quién enviar: no se envía a nadie por defecto.")
    cfg, motivo = config_gmail()
    acc, cats = D["acc"], D["cats"]
    filas = []
    for u in opciones(cats, "responsable"):
        c = D["contactos"].get(u.lower(), {})
        r = correo.resumen(acc[acc["responsable"] == u])
        filas.append(dict(usuario=u, email=c.get("email", ""), nombre=c.get("nombre") or u, **r))
    tab = pd.DataFrame(filas, columns=["usuario", "email", "nombre", "pendientes", "proceso", "cerradas",
                                       "atrasadas", "riesgo"])
    con = tab[tab["email"] != ""].set_index("usuario")
    if len(tab) - len(con):
        st.caption(f"ℹ️ {len(tab) - len(con)} responsable(s) sin correo en 'accesos' (columna `email`) no se pueden elegir.")
    if con.empty:
        st.info("Agrega la columna `email` en la pestaña `accesos` para poder enviar alertas.")
        return

    def todos_con_abiertas():
        st.session_state["al_dest"] = [u for u in con.index if con.loc[u, "pendientes"] + con.loc[u, "proceso"] > 0]

    b1, b2, _ = st.columns([1.3, 1, 3])
    b1.button("Elegir a quienes tienen acciones abiertas", on_click=todos_con_abiertas, width="stretch")
    b2.button("Quitar selección", on_click=lambda: st.session_state.update(al_dest=[]), width="stretch")
    elegidos = st.multiselect(
        "Destinatarios", list(con.index), key="al_dest", placeholder="Elige a quién enviar (ninguno por defecto)",
        format_func=lambda u: f"{con.loc[u, 'nombre']} · {con.loc[u, 'email']} · "
                              f"{con.loc[u, 'pendientes']} pend. / {con.loc[u, 'proceso']} en proceso / "
                              f"{con.loc[u, 'atrasadas']} atrasadas")
    if not elegidos:
        st.info("Elige al menos un destinatario para ver la vista previa y enviar.")
    else:
        prev = st.selectbox("Vista previa del correo de", elegidos,
                            format_func=lambda u: f"{con.loc[u, 'nombre']} ({con.loc[u, 'email']})")
        url = (cfg or {}).get("url_portal", "")
        asunto, cuerpo, _ = correo.construir(con.loc[prev, "nombre"], acc[acc["responsable"] == prev], hoy(), url)
        st.caption(f"**Asunto:** {asunto}")
        if hasattr(st, "iframe"):       # el HTML ya viene escapado (html.escape) en correo.construir
            st.iframe(cuerpo, height=760)
        else:                           # versiones antiguas de Streamlit
            components.html(cuerpo, height=760, scrolling=True)
    st.divider()
    if not cfg:
        st.warning(f"Para enviar falta configurar Gmail en los secrets (sección `[gmail]` con `user` y "
                   f"`app_password`). {motivo} Mientras tanto puedes ver las vistas previas.")
    confirmo = st.checkbox(f"Confirmo que quiero enviar {len(elegidos)} correo(s) ahora", key="al_ok",
                           disabled=not (cfg and elegidos))
    if st.button("📧 Enviar a los elegidos", type="primary", disabled=not (cfg and elegidos and confirmo)):
        url = cfg.get("url_portal", "")
        msgs = []
        for u in elegidos:
            asunto, cuerpo, texto = correo.construir(con.loc[u, "nombre"], acc[acc["responsable"] == u], hoy(), url)
            msgs.append(correo.mensaje("".join(str(cfg["user"]).split()), cfg.get("nombre", "Calidad Kaizen"),
                                       con.loc[u, "email"], asunto, cuerpo, texto,
                                       cc="".join(str(cfg.get("cc", "")).split())))
        try:
            with st.spinner("Enviando…"):
                res = correo.enviar(msgs, cfg["user"], cfg["app_password"])
        except Exception as e:
            st.error(f"No se pudo conectar con Gmail. {correo.explicar_error(e)}")
            return
        st.session_state["al_resultado"] = res
        del st.session_state["al_ok"]
        st.rerun()
    for destino, ok, error in st.session_state.get("al_resultado", []):
        if ok:      # (if/else, no una llamada suelta: la "magia" de Streamlit mostraría su retorno: None)
            st.success(f"✅ Enviado a {destino}")
        else:
            st.error(f"❌ Falló {destino}: {error}")


def control(D):
    st.header("🛠️ Control de datos")
    c1, c2 = st.columns([4, 1.2], vertical_alignment="center")
    c1.caption(f"Datos leídos a las {D['hora']} · se refrescan solos cada {TTL} s o al guardar.")
    if c2.button("🔄 Actualizar ahora", width="stretch"):
        cargar.clear()
        st.rerun()
    resp = opciones(D["cats"], "responsable")
    propias = [u for u in resp if u.lower() in D["claves"] and u.lower() not in ("admin", "usuario")]
    st.info(f"🔑 {len(propias)} de {len(resp)} responsables tienen contraseña propia en 'accesos'; "
            "el resto entra con la clave compartida (fila 'usuario').")
    for a in D["adv"]:
        st.warning(a)
    if not D["adv"]:
        st.success("Sin advertencias de integridad en el plan ni en los seguimientos.")
    st.subheader("Valores del plan que no están en 'catalogos'")
    st.caption("Los datos antiguos se siguen mostrando. Agrega estos valores al catálogo o corrígelos en el plan.")
    filas = []
    for cat, etiqueta in CATS.items():
        validos = set(D["cats"].loc[D["cats"]["categoria"] == cat, "valor"])
        malos = D["acc"][(D["acc"][cat] != "") & ~D["acc"][cat].isin(validos)][cat].value_counts()
        filas += [{"Campo": etiqueta, "Valor": v, "Acciones": int(n)} for v, n in malos.items()]
    if filas:
        st.dataframe(pd.DataFrame(filas), hide_index=True, width="stretch")
    else:
        st.success("Todos los valores del plan están en el catálogo.")


# ============================================================ pestañas que faltan en el Sheet
def asistente(f):
    st.title("🏭 Portal Kaizen · configuración inicial")
    st.warning("Al Google Sheet le faltan estas pestañas: **" + ", ".join(f.faltan) + "**")
    st.markdown("Descarga los archivos, y en el Sheet usa **Archivo → Importar → Subir → Insertar nuevas hojas**. "
                "Cada pestaña debe llamarse igual que el archivo (`catalogos`, `accesos`). Luego recarga esta página.")
    if "catalogos" in f.faltan and "plan_kaizen" in f.datos:
        cat = derivar_catalogos(f.datos["plan_kaizen"])
        st.download_button("⬇️ catalogos.csv (todos los valores actuales del plan)",
                           cat.to_csv(index=False).encode("utf-8-sig"), file_name="catalogos.csv")
        st.caption("El orden de las filas es el orden de las listas. Para quitar una opción, borra su fila.")
    if "accesos" in f.faltan:
        st.download_button("⬇️ accesos.csv (cambia las claves antes de importar)",
                           "usuario,clave\nadmin,CambiaEstaClaveAdmin1\nusuario,CambiaEstaClaveCompartida1\n".encode(),
                           file_name="accesos.csv")
        st.caption("`admin` = clave del administrador · `usuario` = clave compartida para quien no tenga una propia. "
                   "Para dar clave individual, agrega una fila con su usuario (el mismo de la lista 'responsable').")


# ============================================================ arranque
def main():
    try:
        D = cargar()
    except Faltan as f:
        asistente(f)
        st.stop()
    except ValueError as e:
        st.error(f"⚠️ {e}")
        st.stop()
    except Exception as e:
        st.error("⚠️ No se pudo leer Google Sheets. Revisa la conexión e inténtalo de nuevo.")
        st.caption(f"Detalle técnico: {type(e).__name__}")
        st.stop()
    rol, usuario = ingresar(D)
    with st.sidebar:
        st.markdown("### 🏭 Portal Kaizen")
        st.markdown(f"**{'Administrador' if rol == 'admin' else usuario}**")
        c1, c2 = st.columns(2)
        if c1.button("🔄", help="Actualizar datos", width="stretch"):
            cargar.clear()
            st.rerun()
        if c2.button("Salir", width="stretch"):
            st.session_state.clear()
            st.rerun()
        st.caption(f"Datos de las {D['hora']}")
        st.divider()
    if st.session_state.pop("ir", None):        # tras guardar/ir a seguimiento: abre la lista
        st.session_state[f"nav_{'admin' if rol == 'admin' else 'usuario'}"] = LISTA[rol]
    mostrar_aviso()
    vista_admin(D) if rol == "admin" else vista_usuario(D, usuario)


from streamlit.runtime.scriptrunner import get_script_run_ctx  # noqa: E402

if get_script_run_ctx():    # solo dentro de Streamlit (así los tests pueden importar este archivo)
    main()
