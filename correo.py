"""Alertas 
"""
import html
import smtplib
import ssl
from email.message import EmailMessage
from email.utils import formataddr

import pandas as pd

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre",
         "octubre", "noviembre", "diciembre"]
COLOR_PLAZO = {"Atrasado": "#d32f2f", "En riesgo": "#ef6c00", "A tiempo": "#2e7d32",
               "Sin fecha": "#757575", "Cerrado": "#2e7d32"}
ORDEN_PLAZO = {"Atrasado": 0, "En riesgo": 1, "A tiempo": 2, "Sin fecha": 3, "Cerrado": 4}
MAX_FILAS = 15
DIAS_CERRADAS = 30
OSCURO, AMARILLO, FONDO = "#1a1a2e", "#ffd400", "#f4f5f9"


def _e(x):
    return html.escape(str(x))


def _corto(t, n=90):
    t = " ".join(str(t).split())
    return t if len(t) <= n else t[:n - 1] + "…"


def resumen(acc):
    """Conteos de las acciones de una persona (DataFrame ya enriquecido)."""
    abiertas = acc[acc["estado"] != "Cerrado"]
    return {"pendientes": int((acc["estado"] == "No iniciado").sum()),
            "proceso": int((acc["estado"] == "En proceso").sum()),
            "cerradas": int((acc["estado"] == "Cerrado").sum()),
            "atrasadas": int((abiertas["plazo"] == "Atrasado").sum()),
            "riesgo": int((abiertas["plazo"] == "En riesgo").sum())}


def _ordenar(df):
    return df.assign(_r=df["plazo"].map(ORDEN_PLAZO)).sort_values(["_r", "deadline", "id"],
                                                                 na_position="last").drop(columns="_r")


def _fecha(t):
    return t.strftime("%d/%m/%Y") if pd.notna(t) else "—"


def _tarjeta(titulo, valor, pie, color):
    return (f'<td width="33%" style="padding:6px"><table width="100%" cellpadding="0" cellspacing="0" '
            f'style="background:#fff;border-top:4px solid {OSCURO};border-radius:10px"><tr><td align="center" '
            f'style="padding:14px 6px"><div style="font-size:12px;letter-spacing:2px;color:#555;font-weight:bold">'
            f'{_e(titulo)}</div><div style="font-size:40px;font-weight:bold;color:{color};line-height:1.3">'
            f'{valor}</div><div style="font-size:12px;color:#999">{_e(pie)}</div></td></tr></table></td>')


def _nota(texto, color="#555"):
    return (f'<table width="100%" cellpadding="0" cellspacing="0" style="margin:8px 0"><tr>'
            f'<td style="background:#fff;border-left:5px solid {color};padding:12px 16px;color:#444;'
            f'font-size:15px">{texto}</td></tr></table>')


def _tabla(titulo, cabeceras, filas, extra=0):
    th = "".join(f'<th align="{a}" style="padding:9px 10px;font-size:14px">{_e(h)}</th>' for h, a in cabeceras)
    cuerpo = "".join(
        "<tr>" + "".join(f'<td align="{a}" style="padding:9px 10px;border-bottom:1px solid #eee;font-size:14px;'
                         f'{estilo}">{celda}</td>' for (celda, estilo), (_, a) in zip(fila, cabeceras)) + "</tr>"
        for fila in filas)
    mas = (f'<div style="font-size:13px;color:#777;padding:6px 2px">…y {extra} más en el portal.</div>'
           if extra > 0 else "")
    return (f'<h3 style="margin:22px 0 8px;color:{OSCURO};font-size:19px">{_e(titulo)}</h3>'
            f'<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-collapse:collapse">'
            f'<tr style="background:{OSCURO};color:#fff">{th}</tr>{cuerpo}</table>{mas}')


def _estado(plazo):
    return f'<b style="color:{COLOR_PLAZO.get(plazo, "#555")}">{_e(plazo)}</b>', ""


def construir(nombre, acc, hoy, url_portal=""):
    """Devuelve (asunto, html, texto) del reporte de una persona. `acc` = sus acciones."""
    r = resumen(acc)
    mes = f"{MESES[hoy.month - 1].capitalize()} {hoy.year}"
    asunto = (f"Kaizen · {_corto(nombre, 40)}: {r['pendientes']} pendientes, {r['proceso']} en proceso, "
              f"{r['atrasadas']} atrasadas ({mes})")
    pend = _ordenar(acc[acc["estado"] == "No iniciado"])
    proc = _ordenar(acc[acc["estado"] == "En proceso"])
    cerr = acc[acc["estado"] == "Cerrado"]
    recientes = cerr[cerr["fecha_ult"] >= hoy - pd.Timedelta(days=DIAS_CERRADAS)].sort_values(
        "fecha_ult", ascending=False)

    partes = [f'<table width="100%" cellpadding="0" cellspacing="0" style="background:{OSCURO};'
              f'border-bottom:5px solid {AMARILLO}"><tr><td style="padding:26px 26px 18px">'
              f'<div style="font-size:30px;font-weight:bold;color:#fff;letter-spacing:2px">KAIZEN · PLAN DE ACCIÓN</div>'
              f'<div style="font-size:17px;color:#dcdcf0;margin-top:6px">Resultados de <b style="color:#fff">'
              f'{_e(nombre.upper())}</b> — {_e(mes)}</div></td></tr></table>',
              f'<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:16px"><tr>'
              + _tarjeta("PENDIENTES", r["pendientes"], "sin iniciar", "#ef6c00" if r["pendientes"] else OSCURO)
              + _tarjeta("EN PROCESO", r["proceso"], "con avance", "#1565c0" if r["proceso"] else OSCURO)
              + _tarjeta("CERRADAS", r["cerradas"], "en total", "#2e7d32") + "</tr></table>"]
    if r["atrasadas"] or r["riesgo"]:
        partes.append(_nota(f"⚠️ <b>{r['atrasadas']}</b> acción(es) atrasada(s) y <b>{r['riesgo']}</b> en riesgo "
                            f"(vencen en los próximos 7 días). Por favor registra su seguimiento.", "#d32f2f"))
    else:
        partes.append(_nota("✅ Estás al día: no tienes acciones atrasadas ni por vencer esta semana.", "#2e7d32"))

    if url_portal:
        partes.append(
            f'<table width="100%" cellpadding="0" cellspacing="0" style="margin:14px 0 4px"><tr><td align="center">'
            f'<a href="{_e(url_portal)}" style="display:inline-block;background:{AMARILLO};color:{OSCURO};'
            f'font-weight:bold;font-size:16px;text-decoration:none;padding:13px 30px;border-radius:8px">'
            f'Registrar mi seguimiento en el portal →</a><div style="font-size:12px;color:#888;margin-top:6px">'
            f'{_e(url_portal)}</div></td></tr></table>')
    cab_p = [("ID", "left"), ("Hallazgo", "left"), ("Deadline", "center"), ("Estado", "center")]
    if len(pend):
        filas = [[(_e(f"#{x.id}"), ""), (_e(_corto(x.hallazgo)), ""), (_e(_fecha(x.deadline)), ""),
                  _estado(x.plazo)] for x in pend.head(MAX_FILAS).itertuples()]
        partes.append(_tabla(f"Acciones pendientes ({len(pend)})", cab_p, filas, len(pend) - MAX_FILAS))
    cab_e = [("ID", "left"), ("Hallazgo", "left"), ("Avance", "center"), ("Deadline", "center"),
             ("Estado", "center")]
    if len(proc):
        filas = [[(_e(f"#{x.id}"), ""), (_e(_corto(x.hallazgo)), ""), (f"{x.porc_avance:.0f}%", ""),
                  (_e(_fecha(x.deadline)), ""), _estado(x.plazo)] for x in proc.head(MAX_FILAS).itertuples()]
        partes.append(_tabla(f"Acciones en proceso ({len(proc)})", cab_e, filas, len(proc) - MAX_FILAS))
    partes.append(f'<h3 style="margin:22px 0 8px;color:{OSCURO};font-size:19px">Resumen de acciones cerradas</h3>')
    partes.append(_nota(f"Tienes <b>{r['cerradas']}</b> acción(es) cerrada(s) en total; "
                        f"<b>{len(recientes)}</b> en los últimos {DIAS_CERRADAS} días.", "#2e7d32"))
    if len(recientes):
        filas = [[(_e(f"#{x.id}"), ""), (_e(_corto(x.hallazgo)), ""), (_e(_fecha(x.fecha_ult)), "")]
                 for x in recientes.head(10).itertuples()]
        partes.append(_tabla("Cerradas recientemente", [("ID", "left"), ("Hallazgo", "left"), ("Cierre", "center")],
                             filas, len(recientes) - 10))
    if not (len(pend) or len(proc) or r["cerradas"]):
        partes.append(_nota("Aún no tienes acciones asignadas."))
    enlace = (f' Ingresa al portal: <a href="{_e(url_portal)}" style="color:#1565c0">{_e(url_portal)}</a>.'
              if url_portal else "")
    partes.append(f'<p style="color:#888;font-size:13px;margin:22px 0 0">Correo enviado por el área de Calidad '
                  f'(Kaizen). Información confidencial: contiene datos de acciones solo para tu gestión.{enlace}</p>')
    cuerpo = (f'<html><body style="margin:0;background:{FONDO};font-family:Arial,Helvetica,sans-serif">'
              f'<table width="100%" cellpadding="0" cellspacing="0" style="background:{FONDO}"><tr>'
              f'<td align="center" style="padding:20px 8px"><table width="800" cellpadding="0" cellspacing="0" '
              f'style="max-width:100%"><tr><td>{"".join(partes)}</td></tr></table></td></tr></table></body></html>')

    lineas = [f"KAIZEN · Plan de acción — {nombre} — {mes}",
              f"Pendientes: {r['pendientes']} · En proceso: {r['proceso']} · Cerradas: {r['cerradas']} · "
              f"Atrasadas: {r['atrasadas']} · En riesgo: {r['riesgo']}", ""]
    for titulo, df in (("PENDIENTES", pend), ("EN PROCESO", proc)):
        if len(df):
            lineas.append(titulo)
            lineas += [f"  #{x.id} {_corto(x.hallazgo, 70)} — vence {_fecha(x.deadline)} ({x.plazo})"
                       for x in df.head(MAX_FILAS).itertuples()]
            lineas.append("")
    if url_portal:
        lineas.append(f"Registra tu seguimiento en el portal: {url_portal}")
    lineas.append("Correo enviado por el área de Calidad (Kaizen). Información confidencial.")
    return asunto, cuerpo, "\n".join(lineas)


def mensaje(remitente, nombre_remitente, destino, asunto, cuerpo_html, texto, cc=""):
    m = EmailMessage()
    m["Subject"], m["From"], m["To"] = asunto, formataddr((nombre_remitente, remitente)), destino
    if cc and cc.strip().lower() != destino.strip().lower():     # copia al área de Calidad (si es otra dirección)
        m["Cc"] = cc.strip()
    m.set_content(texto)
    m.add_alternative(cuerpo_html, subtype="html")
    return m


def enviar(mensajes, usuario, app_password):
    """Envía por Gmail (SMTP SSL). Devuelve [(destino, ok, error)]. Una falla no detiene las demás."""
    resultados = []
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context(), timeout=30) as s:
        s.login(usuario, app_password.replace(" ", ""))     # Google la muestra con espacios; sin ellos también sirve
        for m in mensajes:
            try:
                s.send_message(m)
                resultados.append((m["To"], True, ""))
            except Exception as e:      # una dirección rechazada no debe frenar al resto
                resultados.append((m["To"], False, str(e)))
    return resultados
