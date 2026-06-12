import os
import re
import json
import base64
import traceback
from contextlib import contextmanager

from dotenv import load_dotenv
load_dotenv()
from typing import Optional, List, Union, Any
from datetime import date, datetime
from io import BytesIO

import pg8000.dbapi as pgdb

import hashlib
import secrets

from fastapi import FastAPI, HTTPException, Request, Response, Query
from fastapi.responses import HTMLResponse, FileResponse, Response as FastResponse, JSONResponse, StreamingResponse
from pydantic import BaseModel

# ── Anthropic ────────────────────────────────────────────────────────────────
try:
    import anthropic as _anthropic
    ANTHROPIC_OK = True
except ImportError:
    ANTHROPIC_OK = False

# ── Excel ─────────────────────────────────────────────────────────────────────
try:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

# ── WeasyPrint (HTML → PDF, requiere GTK/Pango — funciona en Linux/Railway) ───
try:
    import weasyprint as _weasyprint
    WEASYPRINT_OK = True
except Exception:
    WEASYPRINT_OK = False

# ── xhtml2pdf (HTML → PDF, puro Python — funciona en Windows y Linux) ─────────
try:
    from xhtml2pdf import pisa as _pisa
    XHTML2PDF_OK = True
except ImportError:
    XHTML2PDF_OK = False

# ── PDF (ReportLab — fallback) ────────────────────────────────────────────────
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib.colors import HexColor, white, black
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
        Image as RLImage, HRFlowable, KeepTogether,
    )
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

# ── DB config ─────────────────────────────────────────────────────────────────
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

DB_HOST     = os.environ.get("DB_HOST",     "localhost")
DB_PORT     = int(os.environ.get("DB_PORT", "5432"))
DB_NAME     = os.environ.get("DB_NAME",     "boom_ofertas")
DB_USER     = os.environ.get("DB_USER",     "postgres")
DB_PASSWORD = os.environ.get("DB_PASSWORD", "natalia2281*")

# ── SMTP config ───────────────────────────────────────────────────────────────
SMTP_HOST     = os.environ.get("SMTP_HOST",     "smtp.office365.com")
SMTP_PORT     = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER     = os.environ.get("SMTP_USER",     "")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")
SMTP_FROM     = os.environ.get("SMTP_FROM",     "nvargas@boomlts.com")

NOTIF_TO = [
    "cnavarro@boomlts.com", "ingenieria@boomlts.com", "jsarmiento@boomlts.com",
    "hseboom@boomlts.com",  "rromerro@boomlts.com",   "analistadocumental@boomlts.com",
    "analistatracking@boomlts.com", "proyectos@boomlts.com",
]
NOTIF_CC = [
    "mjamis@boomlts.com", "operaciones@boomlts.com",
    "bborrego@boomlts.com", "comercial@boomlts.com",
]

print(f"[DB] host={DB_HOST} port={DB_PORT} db={DB_NAME} user={DB_USER}")

# ── BOOM comercial info ───────────────────────────────────────────────────────
COMERCIAL_INFO = {
    "Natalia Vargas":   ("Ejecutiva Comercial",   "nvargas@boomlts.com"),
    "Boris Borrego":    ("Gerente General",        "bborrego@boomlts.com"),
    "Willington Ortiz": ("Director Comercial",    "comercial@boomlts.com"),
}


def _logo_src() -> str:
    try:
        logo_path = os.path.join(os.path.dirname(__file__), "templates", "boom_logo.b64")
        with open(logo_path, "r") as f:
            return f.read().strip()
    except Exception:
        return ""


def _fmt_cop(v) -> str:
    try:
        return "${:,}".format(int(v)).replace(",", ".")
    except Exception:
        return "$0"


def _fmt_ref(num) -> str:
    n = str(num).zfill(6)
    return f"{n[:2]}-{n[2:]}"


# ── Auth helpers ──────────────────────────────────────────────────────────────
def _hash_pw(password: str) -> str:
    salt = secrets.token_hex(16)
    key = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 200_000)
    return f"{salt}:{key.hex()}"


def _verify_pw(password: str, stored: str) -> bool:
    try:
        salt, key_hex = stored.split(":", 1)
        key = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 200_000)
        return secrets.compare_digest(key.hex(), key_hex)
    except Exception:
        return False


def _logo_for_pdf(height_cm: float = 1.1, max_width_cm: float = 3.5):
    """Return a reportlab Image for the BOOM logo sized to fit the header cell."""
    if not REPORTLAB_OK:
        return None
    try:
        from reportlab.lib.utils import ImageReader
        logo_path = os.path.join(os.path.dirname(__file__), "templates", "boom_logo.b64")
        with open(logo_path, "r") as f:
            b64_str = f.read().strip()
        if "," in b64_str:
            b64_str = b64_str.split(",", 1)[1]
        img_bytes = base64.b64decode(b64_str)
        iw, ih = ImageReader(BytesIO(img_bytes)).getSize()
        aspect = iw / ih if ih else 1
        target_h = height_cm * cm
        target_w = min(target_h * aspect, max_width_cm * cm)
        target_h = target_w / aspect
        return RLImage(BytesIO(img_bytes), width=target_w, height=target_h)
    except Exception as exc:
        print(f"[PDF] logo error: {exc}")
        return None


def _parsear_detalle(texto: str) -> list:
    """Parse a pricing block into a list of equipment dicts.

    Expected format (one block per empty line):
        Nombre del equipo
        • Descripción / configuración
        • $XX.XXX.XXX,oo
    """
    items: list = []
    current: dict | None = None

    for raw in texto.strip().split("\n"):
        line = raw.strip()

        if not line:
            if current is not None:
                items.append(current)
                current = None
            continue

        is_bullet = bool(re.match(r"^[•\-\*•–]", line))
        clean = re.sub(r"^[•\-\*•–]\s*", "", line).strip()

        if not is_bullet:
            # New equipment – save previous block first
            if current is not None:
                items.append(current)
            # Try to extract leading quantity:  "5 Contenedor 40HC"  or  "Contenedor x5"
            qty = 1
            name = line
            m_qty = re.match(r"^(\d+)\s+(.+)$", line)
            if m_qty:
                qty = int(m_qty.group(1))
                name = m_qty.group(2)
            else:
                m_qty2 = re.search(r"\s+[xX]\s*(\d+)$", line)
                if m_qty2:
                    qty = int(m_qty2.group(1))
                    name = line[: m_qty2.start()]
            current = {"equipo": name.strip(), "dimensiones": "", "cant": qty,
                       "config": "", "valor_unit": 0}
        else:
            if current is None:
                current = {"equipo": "", "dimensiones": "", "cant": 1,
                           "config": "", "valor_unit": 0}
            price_m = re.search(r"\$\s*([\d.,]+)", clean)
            if price_m:
                p = price_m.group(1)
                p = re.sub(r",?[oO]{2}$", "", p)   # strip trailing ,oo
                p = p.replace(".", "").replace(",", "")
                try:
                    current["valor_unit"] = int(p)
                except Exception:
                    pass
            else:
                current["config"] = (current["config"] + "\n" + clean).strip()

    if current is not None:
        items.append(current)

    return [i for i in items if i.get("equipo") or i.get("config")]


def _auto_notas(equipos: list, texto_cliente: str = "",
                origen: str = "", destino: str = "") -> str:
    """Apply BOOM Logistics business rules to generate standard technical notes."""
    all_cfg = " ".join(
        (eq.get("config", "") or "") + " " + (eq.get("equipo", "") or "")
        for eq in equipos
    ).lower()
    all_text = all_cfg + " " + texto_cliente.lower()

    # ── Detect vehicle type for stand-by rate ─────────────────────────────
    is_cama5 = any(w in all_text for w in ["5 ejes", "5ejes", "cama baja 5", "cb5"])
    is_cama4 = any(w in all_text for w in ["4 ejes", "4ejes", "cama baja 4", "cb4"])
    is_cama3 = any(w in all_text for w in [
        "3 ejes", "3ejes", "cama baja 3", "cb3", "camabaja", "cama baja",
    ])
    is_cama_alta = any(w in all_text for w in [
        "cama alta", "camalta", "cama-alta", "extensible",
    ])

    if is_cama5:
        standby_valor = "$2.600.000"
        standby_tipo  = "Cama Baja 5 Ejes"
    elif is_cama4:
        standby_valor = "$1.800.000"
        standby_tipo  = "Cama Baja 4 Ejes"
    elif is_cama3:
        standby_valor = "$1.500.000"
        standby_tipo  = "Cama Baja 3 Ejes"
    else:
        standby_valor = "$1.200.000"
        standby_tipo  = "Cama Alta Extensible"

    # ── Detect extra/special service ──────────────────────────────────────
    is_izaje = any(w in all_text for w in [
        "izaje", "grúa", "grua", "modular en operación",
        "modular en operacion", "izar", "montaje de", "pluma",
    ])
    is_extra = any(w in all_text for w in [
        "extradi", "extrapesa", "escolta", "tecnólog", "tecnolog",
        "2 esc", "1 esc", "oversize",
    ])
    is_modular = any(w in all_text for w in ["modular", "spmt", "self-propelled"])

    for m_dim in re.finditer(r'(\d+(?:[.,]\d+)?)\s*m(?:\b|\s)', all_text):
        try:
            if float(m_dim.group(1).replace(",", ".")) > 3.0:
                is_extra = True
                break
        except Exception:
            pass

    tiempos_libres = (
        "12 horas para cargue / 12 horas para descargue" if is_modular
        else "6 horas para cargue / 6 horas para descargue"
    )

    notas = []

    # ── Fijos siempre presentes ────────────────────────────────────────────
    notas.append(f"Origen: {origen}" if origen else "Origen: ")
    notas.append(f"Destino: {destino}" if destino else "Destino: ")
    notas.append("Esquema de seguridad: ")

    if is_izaje:
        notas.append(
            f"Stand-by {standby_tipo}: {standby_valor} COP/día por unidad. "
            "Stand-by de grúa/equipo especializado: valor proporcional a la hora según equipo asignado."
        )
        notas.append(
            "Inicio de operación: Se requiere visita técnica previa al inicio "
            "de labores de izaje/montaje. El tiempo de alistamiento está incluido."
        )
    else:
        notas.append(f"Stand-by {standby_tipo}: {standby_valor} COP/día por unidad en espera.")

    if is_extra:
        notas.append(
            "Esquema extradimensionado/extrapesado: se incluyen 2 escoltas + 2 tecnólogos por despacho."
        )

    notas.append(f"Tiempos libres: {tiempos_libres}.")

    notas.append(
        "Devolución: Las tarifas incluyen el retorno (viaje vacío) del equipo "
        "hasta el punto de origen o patio BOOM."
    )
    notas.append(
        "Permisos: Incluye gestión de permisos de tránsito ante autoridades "
        "competentes (según aplique y reglamentación vigente)."
    )


    return "\n".join(notas)


# ── HTML offer generation ─────────────────────────────────────────────────────

_STANDBY_RATES = [
    (["modular 6","modular6","6 cuna 6","6cuna6"],                        "$8.500.000/día", "12 horas", "12 horas"),
    (["modular 5","modular5","5 cuna 5","5cuna5",
      "modular 4","modular4","4 cuna 4","4cuna4"],                        "$8.500.000/día", "12 horas", "12 horas"),
    (["semi modular","semimodular","2v4","modular"],                       "$2.800.000/día", "12 horas", "12 horas"),
    (["cama baja 5","camabaja5","cb5","5 ejes","60 ton","60ton"],          "$2.600.000/día", "6 horas",  "6 horas"),
    (["cama baja 4","camabaja4","cb4","4 ejes","45 ton","45ton"],          "$1.800.000/día", "8 horas",  "8 horas"),
    (["cama baja 3","camabaja3","cb3","3 ejes","30 ton","30ton",
      "cama baja","camabaja","cama plana"],                                "$1.200.000/día", "6 horas",  "6 horas"),
    (["camión turbo","camion turbo","turbo","sencillo","camioneta"],       "$550.000/día",   "6 horas",  "6 horas"),
    (["cama alta","camalta","patineta","extensible"],                      "$1.200.000/día", "6 horas",  "6 horas"),
]

def _standby_for_equipo(eq_name: str, eq_config: str) -> tuple:
    txt = (eq_name + " " + eq_config).lower()
    for keywords, rate, libre_c, libre_d in _STANDBY_RATES:
        if any(k in txt for k in keywords):
            return rate, libre_c, libre_d
    return "$1.200.000/día", "6 horas", "6 horas"


def _personal_from_config(config: str) -> str:
    c = config.lower()
    esc = re.search(r"(\d+)\s*escolt", c)
    tec = re.search(r"(\d+)\s*tecn", c)
    parts = []
    if esc:
        parts.append(f"{esc.group(1)} Esc")
    if tec:
        parts.append(f"{tec.group(1)} Tec")
    return " &bull; ".join(parts) if parts else "&mdash;"


def generar_html_oferta(data: dict) -> str:
    ref_fmt       = _fmt_ref(data.get("ref", "260001"))
    cliente       = data.get("cliente", "")
    contacto      = data.get("contacto", "") or ""
    ref_cliente   = data.get("ref_cliente", "") or ""
    cliente_final = data.get("cliente_final", "") or ""
    origen        = data.get("origen", "")
    destino       = data.get("destino", "")
    mes_anio      = data.get("mes_anio", datetime.now().strftime("%b %Y").upper())
    descripcion   = data.get("descripcion", "")
    comercial     = data.get("comercial", "Natalia Vargas")
    cargo_com, email_com = COMERCIAL_INFO.get(comercial, ("Ejecutiva Comercial", "nvargas@boomlts.com"))

    equipos      = data.get("equipos", []) or []
    cargo_items  = data.get("cargo_items", []) or []
    notas_raw    = data.get("notas", "")
    forma_pago   = data.get("forma_pago", "50% anticipo / 50% a 30 días tras radicación de factura")
    vigencia     = data.get("vigencia", 30)
    poliza_carga = data.get("poliza_carga", "Hasta $4.000.000.000 COP")
    poliza_rc    = data.get("poliza_rc",    "Hasta $4.000.000.000 COP")
    excl_raw     = data.get("exclusiones",
        "Permisos de tránsito, operación y pólizas asociadas (a cargo del cliente)\n"
        "Servicios, recursos o actividades no descritos explícitamente en esta oferta")

    logo_src = _logo_src()
    logo_html = (f'<img src="{logo_src}" alt="BOOM Logistics" style="height:42px;width:auto;">'
                 if logo_src else
                 '<span style="color:white;font-weight:bold;font-size:18px;letter-spacing:1px;">BOOM</span>')

    # ── Ref-bar components ────────────────────────────────────────────────────
    ruta = ""
    if origen and destino:
        ruta = f" &nbsp;|&nbsp; {origen.upper()} &#8594; {destino.upper()}"
    elif origen:
        ruta = f" &nbsp;|&nbsp; {origen.upper()}"

    ref_extra_html = ""
    if ref_cliente:
        ref_extra_html = f" &nbsp;|&nbsp; {ref_cliente}"
        if cliente_final:
            ref_extra_html += f" / {cliente_final}"

    desc_bar = f" &nbsp;&mdash;&nbsp; {descripcion}" if descripcion else ""

    # ── Greeting ──────────────────────────────────────────────────────────────
    saludo_nombre = contacto if contacto else f"equipo {cliente}"
    intro_txt = descripcion if descripcion else "la prestación de servicios de logística especializada"

    # ── Detect service flavors ────────────────────────────────────────────────
    all_eq_txt = " ".join(
        (e.get("equipo","") + " " + e.get("config","")).lower() for e in equipos
    ) + " " + notas_raw.lower()
    has_izaje = any(w in all_eq_txt for w in
                    ["izaje","grúa","grua","modular","patineta","izar","montaje"])

    # ── Cargo technical section ────────────────────────────────────────────────
    cargo_section_html = ""
    valid_cargo = [c for c in cargo_items if c.get("descripcion")]
    if valid_cargo:
        show_dims = any(c.get("dimensiones") for c in valid_cargo)
        show_peso = any(c.get("peso")        for c in valid_cargo)
        show_vol  = any(c.get("volumen")     for c in valid_cargo)

        # Spec-grid from first cargo item that has dimensional data
        spec_grid_html = ""
        fc = next((c for c in valid_cargo if c.get("dimensiones") or c.get("peso")), None)
        if fc:
            spec_items = []
            dims_raw = fc.get("dimensiones", "")
            if dims_raw:
                parts = re.split(r'\s*[×xX*]\s*', dims_raw)
                if len(parts) >= 3:
                    spec_items.append(f'<div class="spec-card"><div class="val">{parts[0].strip()}</div><div class="lbl">Largo c/u</div></div>')
                    spec_items.append(f'<div class="spec-card"><div class="val">{parts[1].strip()}</div><div class="lbl">Ancho c/u</div></div>')
                    spec_items.append(f'<div class="spec-card"><div class="val">{parts[2].strip()}</div><div class="lbl">Alto c/u</div></div>')
            if fc.get("peso"):
                spec_items.append(f'<div class="spec-card"><div class="val">{fc["peso"]}</div><div class="lbl">Peso c/u</div></div>')
            elif fc.get("volumen"):
                spec_items.append(f'<div class="spec-card"><div class="val">{fc["volumen"]}</div><div class="lbl">Volumen</div></div>')
            if spec_items:
                spec_grid_html = '<div class="spec-grid">' + "".join(spec_items) + '</div>'

        # Cargo table headers
        cargo_hdrs = '<th style="text-align:left;">Commodity</th><th>Cant.</th>'
        if show_dims: cargo_hdrs += '<th>Dimensiones</th>'
        if show_peso: cargo_hdrs += '<th>Peso c/u</th>'
        if show_vol:  cargo_hdrs += '<th style="text-align:right;">Volumen</th>'

        # Cargo table rows
        cargo_rows_html = ""
        for ci in valid_cargo:
            desc   = ci.get("descripcion","")
            tipo   = ci.get("tipo","")
            cant_c = ci.get("cant", 1)
            dims   = ci.get("dimensiones","")
            peso   = ci.get("peso","")
            vol    = ci.get("volumen","")
            orig_d = ci.get("origen_detalle","")
            dest_d = ci.get("destino_detalle","")
            sub_sp = []
            if tipo: sub_sp.append(tipo)
            if orig_d or dest_d:
                sub_sp.append((orig_d or "") + (" &#8594; " if orig_d and dest_d else "") + (dest_d or ""))
            sub_html = (f'<br><span style="font-size:10px;color:#666;">{" &bull; ".join(sub_sp)}</span>'
                        if sub_sp else "")
            cargo_rows_html += f"""
      <tr>
        <td style="text-align:left;"><strong>{desc}</strong>{sub_html}</td>
        <td>{cant_c}</td>
        {'<td>' + (dims or '&mdash;') + '</td>' if show_dims else ''}
        {'<td>' + (peso or '&mdash;') + '</td>' if show_peso else ''}
        {'<td style="text-align:right;">' + (vol or '&mdash;') + '</td>' if show_vol else ''}
      </tr>"""

        cargo_section_html = f"""
  <div class="section-title">1. DETALLE T&Eacute;CNICO DE LA CARGA</div>
  {spec_grid_html}
  <div class="table-scroll">
  <table class="det">
    <thead><tr>{cargo_hdrs}</tr></thead>
    <tbody>{cargo_rows_html}
    </tbody>
  </table>
  </div>"""

    # ── Dynamic section numbering ─────────────────────────────────────────────
    sec = 2 if cargo_section_html else 1

    # ── Equipment rows: Equipo | Cant | Tarifa c/u | Total ────────────────────
    total = 0
    eq_rows_html      = ""
    summary_rows_html = ""
    standby_rows_html = ""

    for e in equipos:
        cant   = int(e.get("cant", 1) or 1)
        v_unit = int(e.get("valor_unit", 0) or 0)
        sub    = cant * v_unit
        total += sub
        eq_name = e.get("equipo", "")
        config  = (e.get("config") or "")

        # Config sub-line (strip newlines, keep as small bullets)
        config_clean = config.strip().replace("\n", " &bull; ")
        tipo_tag = (f'<br><span style="font-size:10px;color:#666;">{config_clean}</span>'
                    if config_clean else "")

        # ITR inclusion box
        is_itr = "itr" in eq_name.lower()

        unit_str  = _fmt_cop(v_unit) if v_unit else "&mdash;"
        total_str = _fmt_cop(sub)    if sub    else "&mdash;"

        eq_rows_html += f"""
      <tr>
        <td style="text-align:left;"><strong>{eq_name}</strong>{tipo_tag}</td>
        <td>{cant}</td>
        <td>{unit_str}</td>
        <td>{total_str}</td>
      </tr>"""

        if v_unit:
            label = f"{eq_name} &mdash; {cant} &times; {unit_str}" if cant > 1 else f"{eq_name} &mdash; {unit_str}"
            summary_rows_html += f"""
      <tr>
        <td style="text-align:left;font-size:12px;"><strong>{label}</strong></td>
        <td style="text-align:right;">{total_str}</td>
      </tr>"""

        # Stand-by
        sb_rate, sb_libre_c, sb_libre_d = _standby_for_equipo(eq_name, config)
        standby_rows_html += f"""
      <tr>
        <td style="text-align:left;"><strong>{eq_name}</strong></td>
        <td>{sb_libre_c}</td>
        <td>{sb_libre_d}</td>
        <td>{sb_rate}</td>
      </tr>"""

    ruta_fase = (f" &mdash; {origen.upper()} &#8594; {destino.upper()}"
                 if origen and destino else
                 (f" &mdash; {origen.upper()}" if origen else ""))

    # Summary table only when > 1 equipment line (or always if there are any values)
    summary_html = ""
    if summary_rows_html and total:
        summary_html = f"""
  <div class="table-scroll" style="margin-top:12px;">
  <table class="det">
    <tbody>
      {summary_rows_html}
      <tr class="total-row">
        <td><strong>TOTAL OFERTA</strong></td>
        <td><strong>{_fmt_cop(total)} COP</strong></td>
      </tr>
    </tbody>
  </table>
  </div>"""
    elif total:
        # Single item — still show total row
        summary_html = ""   # total already shown in eq_rows_html last column

    # ── Notes → clean bullet list ─────────────────────────────────────────────
    skip_pfx = ("origen:", "destino:", "stand-by", "standby", "tiempos libres")
    note_items = [
        ln for ln in (n.strip() for n in notas_raw.split("\n") if n.strip())
        if not any(ln.lower().startswith(p) for p in skip_pfx)
    ]
    notes_section_html = ""
    if note_items:
        notes_li = "".join(f"    <li>{n}</li>\n" for n in note_items)
        notes_section_html = f"""
  <div class="section-title">{sec + 1}. NOTAS T&Eacute;CNICAS DE OPERACI&Oacute;N</div>
  <ul class="notas">
{notes_li}  </ul>"""

    standby_note = "* Las horas adicionales ser&aacute;n cobradas proporcionalmente seg&uacute;n tarifa establecida."
    if has_izaje:
        standby_note += " El cobro del equipo de izaje inicia desde la llegada al sitio designado."

    # ── Exclusiones ───────────────────────────────────────────────────────────
    excl_items = [e.strip() for e in excl_raw.split("\n") if e.strip()]
    excl_html  = "".join(f"    <li>{e}</li>\n" for e in excl_items)

    cond_num = sec + (2 if note_items else 1)
    excl_num = cond_num + 1

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Oferta {ref_fmt} | {cliente}</title>
<style>
*{{box-sizing:border-box;}}
body{{font-family:Arial,sans-serif;font-size:13px;color:#1B2A4A;margin:0;padding:16px;background:#f4f4f4;}}
.wrapper{{max-width:860px;margin:0 auto;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.1);}}
.header-bar{{background:#1B2A4A;padding:14px 20px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;}}
.header-bar img{{height:42px;width:auto;}}
.header-info{{color:#fff;font-size:11px;line-height:1.6;text-align:right;}}
.header-info strong{{font-size:12px;display:block;}}
.ref-bar{{background:#E8601C;padding:8px 20px;}}
.ref-bar p{{margin:0;color:#fff;font-size:11px;font-weight:bold;}}
.body{{padding:20px;}}
.greeting p{{font-size:13px;line-height:1.6;margin:0 0 10px 0;}}
.section-title{{background:#1B2A4A;color:#fff;font-size:12px;font-weight:bold;padding:7px 12px;margin:20px 0 8px 0;border-radius:3px;}}
.fase-title{{background:#2d4a7a;color:#fff;font-size:11px;font-weight:bold;padding:5px 12px;margin:10px 0 4px 0;border-radius:3px;}}
.spec-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:10px 0;}}
.spec-card{{background:#f0f4fa;border-radius:5px;padding:9px 8px;text-align:center;}}
.spec-card .val{{font-size:14px;font-weight:bold;color:#E8601C;}}
.spec-card .lbl{{font-size:10px;color:#666;margin-top:2px;}}
.incl-box{{background:#e8f5e9;border-left:4px solid #43a047;padding:8px 14px;font-size:12px;line-height:1.7;margin:6px 0 10px 0;}}
.note-box{{background:#fff8e1;border-left:4px solid #f9a825;padding:8px 14px;font-size:12px;line-height:1.7;margin:8px 0;}}
.table-scroll{{width:100%;overflow-x:auto;}}
table.det{{width:100%;border-collapse:collapse;font-size:12px;}}
table.det th{{background:#1B2A4A;color:#fff;padding:7px 8px;font-size:11px;text-align:center;white-space:nowrap;}}
table.det th:first-child{{text-align:left;}}
table.det td{{padding:7px 8px;border-bottom:1px solid #e5e5e5;vertical-align:middle;font-size:12px;}}
table.det td:not(:first-child){{text-align:center;}}
table.det td:last-child{{text-align:right;font-weight:bold;}}
table.det tr:nth-child(even) td{{background:#f9f9f9;}}
.subtotal-row td{{background:#e0e8f5!important;font-weight:bold;padding:7px 8px;border:none!important;color:#1B2A4A;}}
.subtotal-row td:last-child{{text-align:right;}}
.total-row td{{background:#E8601C!important;color:#fff;padding:9px 10px;font-weight:bold;font-size:13px;border:none!important;}}
.total-row td:last-child{{text-align:right;}}
ul.notas{{margin:0;padding-left:18px;}}
ul.notas li{{margin-bottom:7px;font-size:13px;line-height:1.6;}}
table.cond{{width:100%;border-collapse:collapse;font-size:13px;}}
table.cond td{{padding:7px 10px;border-bottom:1px solid #e5e5e5;vertical-align:top;line-height:1.5;}}
table.cond td:first-child{{font-weight:bold;width:40%;background:#f5f5f5;}}
.footer{{border-top:1px solid #e0e0e0;margin-top:22px;padding-top:12px;font-size:13px;color:#444;}}
.firma-nombre{{font-weight:bold;color:#1B2A4A;font-size:14px;margin:0 0 2px 0;}}
.firma-cargo{{color:#E8601C;font-size:13px;margin:2px 0;}}
.pie{{color:#aaa;font-size:11px;margin-top:10px;border-top:1px solid #eee;padding-top:8px;text-align:center;}}
</style>
</head>
<body>
<div class="wrapper">
<div class="header-bar">
  {logo_html}
  <div class="header-info">
    <strong>BOOM LOGISTICS COLOMBIA S.A.S.</strong>
    Soluciones de Transporte Especializado
  </div>
</div>
<div class="ref-bar">
  <p>REF: {ref_fmt} &nbsp;|&nbsp; {cliente.upper()}{ruta} &nbsp;|&nbsp; {mes_anio}{ref_extra_html}{desc_bar}</p>
</div>

<div class="body">
  <div class="greeting">
    <p>Hola {saludo_nombre},</p>
    <p>En atenci&oacute;n a su solicitud, BOOM Logistics Colombia S.A.S. presenta a continuaci&oacute;n su propuesta para <strong>{intro_txt}</strong>.</p>
  </div>

  {cargo_section_html}

  <!-- ===== PROPUESTA ECONÓMICA ===== -->
  <div class="section-title">{sec}. PROPUESTA ECON&Oacute;MICA CON EQUIPO</div>

  <div class="fase-title">TRANSPORTE{ruta_fase}</div>
  <div class="table-scroll">
  <table class="det">
    <thead><tr>
      <th style="text-align:left;">Servicio / Equipo</th>
      <th>Cant.</th>
      <th>Tarifa c/u</th>
      <th style="text-align:right;">Total</th>
    </tr></thead>
    <tbody>
      {eq_rows_html}
    </tbody>
  </table>
  </div>

  {summary_html}

  <!-- Stand-by -->
  <p style="font-size:12px;font-weight:bold;color:#1B2A4A;margin:16px 0 5px 0;">STAND-BY</p>
  <div class="table-scroll">
  <table class="det">
    <thead><tr>
      <th style="text-align:left;">Equipo</th>
      <th>Tiempo libre cargue</th>
      <th>Tiempo libre descargue</th>
      <th style="text-align:right;">Stand-By (por d&iacute;a)</th>
    </tr></thead>
    <tbody>
      {standby_rows_html}
    </tbody>
  </table>
  </div>
  <p style="font-size:11px;color:#666;margin:4px 0 0 2px;">{standby_note}</p>

  {notes_section_html}

  <!-- ===== CONDICIONES COMERCIALES ===== -->
  <div class="section-title">{cond_num}. CONDICIONES COMERCIALES</div>
  <table class="cond">
    <tr><td>Forma de pago</td><td>{forma_pago}</td></tr>
    <tr><td>Moneda</td><td>Pesos colombianos (COP)</td></tr>
    <tr><td>Vigencia</td><td>{vigencia} d&iacute;as calendario a partir de la fecha de emisi&oacute;n</td></tr>
    <tr><td>P&oacute;liza de carga</td><td>{poliza_carga}</td></tr>
    <tr><td>P&oacute;liza RCE</td><td>{poliza_rc}</td></tr>
  </table>

  <!-- ===== EXCLUSIONES ===== -->
  <div class="section-title">{excl_num}. EXCLUSIONES</div>
  <ul class="notas">
{excl_html}  </ul>

  <!-- ===== FIRMA ===== -->
  <div class="footer">
    <p style="margin:0 0 10px 0;">Quedamos atentos a sus comentarios.</p>
    <p style="margin:0 0 2px 0;">Cordialmente,</p>
    <p class="firma-nombre">{comercial}</p>
    <p class="firma-cargo">{cargo_com}</p>
    <p style="margin:2px 0;">BOOM Logistics Colombia S.A.S.</p>
    <p style="margin:2px 0;">{email_com}</p>
    <p class="pie">BOOM LOGISTICS S.A.S. &nbsp;|&nbsp; Oferta v&aacute;lida por {vigencia} d&iacute;as &nbsp;|&nbsp; Ref: {ref_fmt} | {cliente.upper()}</p>
  </div>
</div>
</div>
</body>
</html>"""


# ── PDF helpers ───────────────────────────────────────────────────────────────
def _detect_service_type(equipos: list, notas: str = "") -> dict:
    """Detects service characteristics for dynamic PDF generation."""
    all_text = " ".join(
        (eq.get("equipo", "") + " " + eq.get("config", "")).lower()
        for eq in equipos
    ) + " " + notas.lower()

    has_izaje = any(kw in all_text for kw in [
        "izaje", "grúa", "grua", "modular", "patineta", "izamiento", "montaje con grúa"
    ])
    has_transport = any(kw in all_text for kw in [
        "transporte", "cama baja", "cama alta", "flete", "contenedor"
    ])
    has_security = any(kw in all_text for kw in ["escolta", "tecnólogo", "tecnologo"])

    phase_nums = set()
    for eq in equipos:
        m = re.search(r"fase\s*(\d+)", eq.get("equipo", "").lower())
        if m:
            phase_nums.add(int(m.group(1)))

    return {
        "has_izaje": has_izaje,
        "has_transport": has_transport,
        "has_security": has_security,
        "is_multiphase": len(phase_nums) > 0,
        "phases": sorted(phase_nums),
    }


def _split_notas(notas_raw: str) -> tuple:
    """Splits notas into (izaje_lines, other_lines)."""
    izaje_kws = ["stand-by", "standby", "/hora", "hora adicional", "proporcional",
                 "cobro inicia", "inicio de tiempo", "horas adicionales serán"]
    izaje, other = [], []
    for n in [n.strip() for n in notas_raw.split("\n") if n.strip()]:
        if any(kw in n.lower() for kw in izaje_kws):
            izaje.append(n)
        else:
            other.append(n)
    return izaje, other


# ── PDF offer generation ──────────────────────────────────────────────────────
def generar_pdf_oferta(data: dict) -> bytes:
    if not REPORTLAB_OK:
        raise RuntimeError("reportlab no está instalado")

    # ── Palette ───────────────────────────────────────────────────────────────
    NAVY    = HexColor("#1B2A4A")
    ORANGE  = HexColor("#E8601C")
    LGRAY   = HexColor("#F7F8FA")
    BORDER  = HexColor("#D8DCE4")
    FGRAY   = HexColor("#EEF0F4")
    MED     = HexColor("#5A6373")
    ACCENT  = HexColor("#F0F4FF")   # very light blue tint for alt rows

    # ── Data ──────────────────────────────────────────────────────────────────
    ref_fmt      = _fmt_ref(data.get("ref", "260001"))
    cliente      = data.get("cliente", "")
    contacto     = data.get("contacto", "")
    email_cl     = data.get("email_cliente", "")
    ref_cliente  = data.get("ref_cliente", "") or ""
    cliente_final= data.get("cliente_final", "") or ""
    origen       = data.get("origen", "")
    destino      = data.get("destino", "")
    mes_anio     = data.get("mes_anio", datetime.now().strftime("%b %Y").upper())
    descripcion  = data.get("descripcion", "")
    comercial    = data.get("comercial", "Natalia Vargas")
    cargo_str, email_com = COMERCIAL_INFO.get(comercial, ("Ejecutiva Comercial", "nvargas@boomlts.com"))
    cargo_items  = data.get("cargo_items", []) or []
    equipos      = data.get("equipos", []) or []
    notas_raw    = data.get("notas", "")
    forma_pago   = data.get("forma_pago", "50% anticipo / 50% a 30 días tras radicación de factura")
    vigencia     = data.get("vigencia", 30)
    poliza_carga = data.get("poliza_carga", "Hasta $4.000.000.000 COP por despacho")
    poliza_rc    = data.get("poliza_rc", "Hasta $4.000.000.000 COP")
    resolucion   = data.get("resolucion", "Carga extradimensionada/extrapesada incluida")
    excl_raw     = data.get("exclusiones", "")
    fecha_str    = data.get("fecha", "") or datetime.now().strftime("%d/%m/%Y")

    svc = _detect_service_type(equipos, notas_raw)
    izaje_notas, other_notas = _split_notas(notas_raw)

    # ── Document ──────────────────────────────────────────────────────────────
    buf = BytesIO()
    ML, MR, MT, MB = 1.6*cm, 1.6*cm, 1.4*cm, 1.9*cm
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=ML, rightMargin=MR, topMargin=MT, bottomMargin=MB,
        title=f"Oferta {ref_fmt} — {cliente}",
    )
    W = A4[0] - ML - MR

    # ── Page footer callback ──────────────────────────────────────────────────
    def _page_footer(canvas, doc):
        canvas.saveState()
        pw = A4[0]
        y  = 0.65 * cm
        canvas.setStrokeColor(BORDER)
        canvas.setLineWidth(0.4)
        canvas.line(ML, y + 0.28*cm, pw - MR, y + 0.28*cm)
        canvas.setFont("Helvetica", 6.5)
        canvas.setFillColor(MED)
        canvas.drawString(ML, y,
            f"BOOM Logistics Colombia S.A.S.  ·  NIT 900.548.985-7  ·  Ref. {ref_fmt}")
        canvas.drawRightString(pw - MR, y,
            f"Página {doc.page}  ·  {cliente.upper()}")
        canvas.restoreState()

    # ── Style factory ─────────────────────────────────────────────────────────
    def ps(name, **kw):
        d = dict(fontName="Helvetica", fontSize=8.5, leading=12,
                 textColor=NAVY, spaceAfter=0, spaceBefore=0)
        d.update(kw)
        return ParagraphStyle(name, **d)

    s_body    = ps("body",    fontSize=9,   leading=14)
    s_white   = ps("white",   textColor=white, fontName="Helvetica-Bold", fontSize=8.5)
    s_cell    = ps("cell",    fontSize=8,   leading=12)
    s_cell_c  = ps("cell_c",  fontSize=8,   leading=12, alignment=TA_CENTER)
    s_cell_r  = ps("cell_r",  fontSize=8,   leading=12, alignment=TA_RIGHT)
    s_cell_w  = ps("cell_w",  fontSize=8,   leading=12, textColor=white,
                   fontName="Helvetica-Bold", alignment=TA_CENTER)
    s_cell_wr = ps("cell_wr", fontSize=8,   leading=12, textColor=white,
                   fontName="Helvetica-Bold", alignment=TA_RIGHT)
    s_lbl     = ps("lbl",     fontSize=8,   leading=12, fontName="Helvetica-Bold")
    s_val     = ps("val",     fontSize=8,   leading=12, textColor=MED)
    s_foot    = ps("foot",    fontSize=7.5, leading=11, textColor=MED, alignment=TA_CENTER)

    SP  = lambda n: Spacer(1, n)   # compact spacer shorthand

    def note_para(text):
        m = re.match(r'^([^:]+:)\s*(.*)', text.strip(), re.DOTALL)
        st = ps("nota", fontSize=8, leading=13, leftIndent=10, spaceAfter=3)
        if m:
            return Paragraph(f"<font color='#E8601C'>▸</font> <b>{m.group(1)}</b> {m.group(2)}", st)
        return Paragraph(f"<font color='#E8601C'>▸</font> {text}", st)

    # ── Section title ─────────────────────────────────────────────────────────
    sec = [0]
    def sec_title(title: str):
        sec[0] += 1
        t = Table(
            [[Paragraph(f"<b>{sec[0]}. {title.upper()}</b>",
                        ps("st", fontName="Helvetica-Bold", fontSize=8.5,
                           textColor=white, leading=12))]],
            colWidths=[W],
        )
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), NAVY),
            ("LEFTPADDING",   (0, 0), (-1, -1), 10),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
            ("TOPPADDING",    (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        return t

    # ── Standard table style ─────────────────────────────────────────────────
    def _eq_table_style(data_rows, n_total_span):
        ts = [
            ("BACKGROUND",    (0, 0), (-1, 0),  NAVY),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING",   (0, 0), (-1, -1), 7),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 7),
            ("TOPPADDING",    (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("GRID",          (0, 0), (-1, -2), 0.4, BORDER),
            ("BACKGROUND",    (0, -1), (-1, -1), ORANGE),
            ("SPAN",          (0, -1), (n_total_span, -1)),
            ("ALIGN",         (0, -1), (0, -1), "RIGHT"),
            ("TOPPADDING",    (0, -1), (-1, -1), 7),
            ("BOTTOMPADDING", (0, -1), (-1, -1), 7),
        ]
        for i in range(1, len(data_rows) - 1):
            ts.append(("BACKGROUND", (0, i), (-1, i), ACCENT if i % 2 == 0 else white))
        return TableStyle(ts)

    story = []

    # ── Header: logo + company ────────────────────────────────────────────────
    logo = _logo_for_pdf(1.1, max_width_cm=3.4)
    company_p = Paragraph(
        '<font name="Helvetica-Bold" size="8.5" color="white">BOOM LOGISTICS COLOMBIA S.A.S.</font><br/>'
        '<font size="6.5" color="#4FC3D8">NIT: 900.548.985-7</font><br/>'
        '<font size="6.5" color="white">Soluciones de Transporte Especializado</font>',
        ps("hdr_r", textColor=white, fontSize=8, leading=10, alignment=TA_RIGHT),
    )
    left_cell = logo if logo else Paragraph("<b>BOOM</b>",
        ps("boom_fb", textColor=white, fontName="Helvetica-Bold", fontSize=15))
    hdr = Table([[left_cell, company_p]], colWidths=[W * 0.22, W * 0.78])
    hdr.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), NAVY),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",         (0, 0), (0,  0),  "LEFT"),
        ("ALIGN",         (1, 0), (1,  0),  "RIGHT"),
        ("LEFTPADDING",   (0, 0), (-1, -1), 12),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 12),
        ("TOPPADDING",    (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
    ]))
    story.append(hdr)

    # ── Orange identifier bar ─────────────────────────────────────────────────
    ruta = ""
    if origen and destino:
        ruta = f"  ·  {origen.upper()} → {destino.upper()}"
    elif origen:
        ruta = f"  ·  {origen.upper()}"
    ref_extra = f"  ·  {ref_cliente}" if ref_cliente else ""
    if ref_extra and cliente_final:
        ref_extra += f" / {cliente_final}"
    ref_t = Table(
        [[Paragraph(
            f"<b>OFERTA  {ref_fmt}  ·  {cliente.upper()}{ruta}  ·  {mes_anio}{ref_extra}</b>",
            ps("refb", textColor=white, fontSize=8.5, leading=12)
        )]],
        colWidths=[W],
    )
    ref_t.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, -1), ORANGE),
        ("LEFTPADDING",  (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
        ("TOPPADDING",   (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 6),
    ]))
    story.append(ref_t)
    story.append(SP(8))

    # ── Client info block ────────────────────────────────────────────────────
    ruta_display = f"{origen}  →  {destino}" if origen and destino else (origen or destino or "—")
    _fd = fecha_str if isinstance(fecha_str, str) else str(fecha_str)

    if svc["is_multiphase"]:
        svc_label = "Proyecto especializado — multifase"
    elif svc["has_izaje"] and svc["has_transport"]:
        svc_label = "Transporte e izaje especializado"
    elif svc["has_izaje"]:
        svc_label = "Operación de izaje especializado"
    else:
        svc_label = "Transporte especializado"

    info_rows = [
        [Paragraph("<b>Para:</b>",       s_lbl), Paragraph(cliente,        s_val),
         Paragraph("<b>Referencia:</b>", s_lbl), Paragraph(ref_fmt,        s_val)],
        [Paragraph("<b>Atención:</b>",   s_lbl), Paragraph(contacto or "—", s_val),
         Paragraph("<b>Fecha:</b>",      s_lbl), Paragraph(_fd,            s_val)],
        [Paragraph("<b>Ruta:</b>",       s_lbl), Paragraph(ruta_display,   s_val),
         Paragraph("<b>Vigencia:</b>",   s_lbl), Paragraph(f"{vigencia} días", s_val)],
        [Paragraph("<b>Servicio:</b>",   s_lbl), Paragraph(svc_label,      s_val),
         Paragraph("<b>Moneda:</b>",     s_lbl), Paragraph("COP",          s_val)],
    ]
    n_base = len(info_rows)
    if descripcion:
        info_rows.append([
            Paragraph(
                f"<b>Descripción:</b>  "
                f"<font color='#5A6373'>{descripcion[:220]}</font>",
                ps("desc_v", fontSize=8, leading=12, textColor=NAVY)
            ),
            "", "", "",
        ])
    info_t = Table(info_rows, colWidths=[W*0.15, W*0.38, W*0.15, W*0.32])
    info_ts = [
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING",   (0, 0), (-1, -1), 7),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 7),
        ("GRID",          (0, 0), (-1, -1), 0.4, BORDER),
        ("BACKGROUND",    (0, 0), (0, n_base - 1), FGRAY),
        ("BACKGROUND",    (2, 0), (2, n_base - 1), FGRAY),
    ]
    if descripcion:
        info_ts += [
            ("SPAN",       (0, n_base), (3, n_base)),
            ("BACKGROUND", (0, n_base), (3, n_base), ACCENT),
        ]
    info_t.setStyle(TableStyle(info_ts))
    story.append(info_t)
    story.append(SP(10))

    # ── Helper: render a section block (title + content) keeping together ──────
    def _section(title, content_flowables):
        return KeepTogether([sec_title(title), SP(3)] + content_flowables + [SP(8)])

    # ── Detect which cargo columns actually have data ─────────────────────────
    def _col_has(items, *keys):
        return any(any(ci.get(k) for k in keys) for ci in items)

    show_dim = _col_has(cargo_items, "dimensiones")
    show_peso = _col_has(cargo_items, "peso")
    show_vol  = _col_has(cargo_items, "volumen")
    show_orig = _col_has(cargo_items, "origen_detalle")
    show_dest = _col_has(cargo_items, "destino_detalle")
    has_detail = show_dim or show_peso or show_vol

    # ── TABLE 1: Physical cargo — full detail or compact ──────────────────────
    if cargo_items:
        if has_detail:
            # ── Full detail table (7 cols, only show populated columns) ───────
            active_cols = ["descripcion"]
            hdr_labels  = ["Descripción"]
            col_ws      = [W * 0.24]
            if show_dim:
                active_cols.append("dimensiones"); hdr_labels.append("Dimensiones"); col_ws.append(W * 0.17)
            if show_peso:
                active_cols.append("peso");        hdr_labels.append("Peso");        col_ws.append(W * 0.10)
            if show_vol:
                active_cols.append("volumen");     hdr_labels.append("Volumen");     col_ws.append(W * 0.11)
            # distribute remaining width to Origen + Destino
            used = sum(col_ws)
            rem  = W - used
            if show_orig:
                active_cols.append("origen_detalle");  hdr_labels.append("Origen");  col_ws.append(rem * 0.45)
            if show_dest:
                active_cols.append("destino_detalle"); hdr_labels.append("Destino"); col_ws.append(rem * 0.55 if show_orig else rem)

            c_hdr = [Paragraph(f"<b>{lbl}</b>", s_cell_w) for lbl in hdr_labels]
            c_data = [c_hdr]
            for ci in cargo_items:
                row = []
                desc = ci.get("descripcion", "") or ""
                tipo = ci.get("tipo", "") or ""
                row.append(Paragraph(
                    f"{desc}<br/><font size='6.5' color='#888'>{tipo}</font>" if tipo else desc,
                    s_cell))
                for col in active_cols[1:]:
                    val = (ci.get(col) or "").replace("\n", "<br/>")
                    row.append(Paragraph(val, s_cell))
                c_data.append(row)

            c_t = Table(c_data, colWidths=col_ws, repeatRows=1)
            c_ts = [
                ("BACKGROUND",    (0, 0), (-1, 0),  NAVY),
                ("VALIGN",        (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING",   (0, 0), (-1, -1), 6),
                ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
                ("TOPPADDING",    (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("GRID",          (0, 0), (-1, -1), 0.4, BORDER),
                ("ALIGN",         (1, 0), (-1, 0),  "CENTER"),
                ("ALIGN",         (1, 1), (-1, -1), "LEFT"),
            ]
            for i in range(1, len(c_data)):
                c_ts.append(("BACKGROUND", (0, i), (-1, i), ACCENT if i % 2 == 0 else white))
            c_t.setStyle(TableStyle(c_ts))
            story.append(_section("DETALLE TÉCNICO DE LA CARGA", [c_t]))

        else:
            # ── Compact combined table: Servicio/Descripción | Equipo | Origen | Destino ──
            # Merge cargo info into the economic table (rendered later as one unified table)
            pass  # handled below via _use_combined flag

    _use_combined = cargo_items and not has_detail

    # ── Security callout ──────────────────────────────────────────────────────
    if svc["has_security"]:
        sec_box = Table(
            [[Paragraph(
                "<b>ESQUEMA DE SEGURIDAD:</b>  2 Escoltas + 2 Tecnólogos incluidos "
                "— carga extradimensionada / extrapesada (ancho &gt; 3.00 m).",
                ps("sec_n", textColor=NAVY, fontSize=8, leading=12)
            )]],
            colWidths=[W],
        )
        sec_box.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), HexColor("#FFF7F3")),
            ("TOPPADDING",    (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ("LEFTPADDING",   (0, 0), (-1, -1), 12),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 8),
            ("BOX",           (0, 0), (-1, -1), 0.4, BORDER),
            ("LINEBELOW",     (0, 0), (-1, -1), 2.5, ORANGE),
        ]))
        story.append(sec_box)
        story.append(SP(8))

    # ── TABLE 2: Economic proposal (or combined if no cargo detail) ──────────
    if _use_combined:
        # ── Combined: Servicio/Descripción | Equipo/Config | Origen | Valor ──
        cb_hdr = [
            Paragraph("<b>Servicio / Descripción</b>", s_cell_w),
            Paragraph("<b>Equipo / Contenedor</b>",    s_cell_w),
            Paragraph("<b>Origen</b>",                 s_cell_w),
            Paragraph("<b>Valor</b>",                  s_cell_w),
        ]
        cw_cb = [W*0.36, W*0.24, W*0.20, W*0.20]
        cb_data = [cb_hdr]
        total_oferta = 0
        for i, eq in enumerate(equipos):
            cant   = int(eq.get("cant", 1) or 1)
            v_unit = int(eq.get("valor_unit", 0) or 0)
            sub    = cant * v_unit
            total_oferta += sub
            cfg = (eq.get("config") or "").strip()
            # match cargo_item if available
            ci = cargo_items[i] if i < len(cargo_items) else {}
            cargo_desc = ci.get("descripcion", "") or ""
            cargo_tipo = ci.get("tipo", "") or ""
            orig_val   = (ci.get("origen_detalle") or origen or "").strip()
            # service cell: equipo name bold + cargo description small
            svc_parts = f"<b>{eq.get('equipo','')}</b>"
            if cargo_desc:
                svc_parts += f"<br/><font size='7' color='#888'>{cargo_desc}"
                if cargo_tipo:
                    svc_parts += f" — {cargo_tipo}"
                svc_parts += "</font>"
            cb_data.append([
                Paragraph(svc_parts, s_cell),
                Paragraph(cfg, s_cell),
                Paragraph(orig_val, s_cell),
                Paragraph(f"<b>{_fmt_cop(sub)}</b>" if sub else "—", s_cell_r),
            ])
        # extra cargo items without matching equipo
        for ci in cargo_items[len(equipos):]:
            dest_val = (ci.get("destino_detalle") or destino or "").strip()
            cb_data.append([
                Paragraph(f"<font color='#888'>{ci.get('descripcion','')}</font>", s_cell),
                Paragraph("—", s_cell_c),
                Paragraph(dest_val, s_cell),
                Paragraph("—", s_cell_c),
            ])
        # total row
        ruta_total = f"{origen.upper()} → {destino.upper()}" if origen and destino else (origen or destino or "")
        cb_data.append([
            Paragraph("<b>TOTAL OFERTA</b>",
                      ps("tot_cb", textColor=white, fontName="Helvetica-Bold",
                         fontSize=10, leading=14, alignment=TA_RIGHT)),
            Paragraph(ruta_total,
                      ps("tot_rt", textColor=white, fontSize=8, leading=12, alignment=TA_CENTER)),
            "",
            Paragraph(f"<b>{_fmt_cop(total_oferta)} COP</b>", s_cell_wr),
        ])
        cb_ts = [
            ("BACKGROUND",    (0, 0), (-1, 0),  NAVY),
            ("VALIGN",        (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING",   (0, 0), (-1, -1), 7),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 7),
            ("TOPPADDING",    (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("GRID",          (0, 0), (-1, -2), 0.4, BORDER),
            ("BACKGROUND",    (0, -1), (-1, -1), ORANGE),
            ("SPAN",          (0, -1), (2, -1)),
            ("ALIGN",         (3, -1), (3, -1), "RIGHT"),
            ("TOPPADDING",    (0, -1), (-1, -1), 7),
            ("BOTTOMPADDING", (0, -1), (-1, -1), 7),
        ]
        for i in range(1, len(cb_data) - 1):
            cb_ts.append(("BACKGROUND", (0, i), (-1, i), ACCENT if i % 2 == 0 else white))
        eq_t = Table(cb_data, colWidths=cw_cb, repeatRows=1)
        eq_t.setStyle(TableStyle(cb_ts))
        story.append(_section("DETALLE TÉCNICO Y ECONÓMICO", [eq_t]))

    else:
        # ── Standard economic table ───────────────────────────────────────────
        e_hdr = [
            Paragraph("<b>Concepto</b>",                s_cell_w),
            Paragraph("<b>Configuración de Equipo</b>", s_cell_w),
            Paragraph("<b>Cant.</b>",                   s_cell_w),
            Paragraph("<b>Valor</b>",                   s_cell_w),
        ]
        cw_e = [W*0.29, W*0.44, W*0.07, W*0.20]

        if svc["is_multiphase"]:
            LIGHT_BLUE = HexColor("#EBF3FF")
            PHASE_BG   = HexColor("#1E3A8A")
            phase_groups = {}
            no_phase_eqs = []
            for eq in equipos:
                _m = re.search(r"fase\s*(\d+)", eq.get("equipo", "").lower())
                if _m:
                    phase_groups.setdefault(int(_m.group(1)), []).append(eq)
                else:
                    no_phase_eqs.append(eq)

            eq_data   = [e_hdr]
            row_types = ["header"]
            grand_total = 0

            for ph in sorted(phase_groups.keys()):
                phase_total = 0
                eq_data.append([
                    Paragraph(f"<b>FASE {ph}</b>",
                              ps(f"ph{ph}h", fontName="Helvetica-Bold", fontSize=8.5,
                                 textColor=white, leading=12)),
                    "", "", "",
                ])
                row_types.append("phase_hdr")
                for eq in phase_groups[ph]:
                    cant   = int(eq.get("cant", 1) or 1)
                    v_unit = int(eq.get("valor_unit", 0) or 0)
                    sub    = cant * v_unit
                    phase_total += sub
                    grand_total += sub
                    cfg = (eq.get("config") or "").replace("\n", "<br/>")
                    eq_data.append([
                        Paragraph(eq.get("equipo", "") or "", s_cell),
                        Paragraph(cfg, s_cell),
                        Paragraph(str(cant), s_cell_c),
                        Paragraph(f"<b>{_fmt_cop(sub)}</b>", s_cell_r),
                    ])
                    row_types.append("item")
                eq_data.append([
                    Paragraph(f"<b>SUBTOTAL FASE {ph}</b>",
                              ps(f"ph{ph}s", fontName="Helvetica-Bold", fontSize=8,
                                 textColor=NAVY, leading=12, alignment=TA_RIGHT)),
                    "", "",
                    Paragraph(f"<b>{_fmt_cop(phase_total)}</b>",
                              ps(f"ph{ph}v", fontName="Helvetica-Bold", fontSize=8,
                                 textColor=NAVY, leading=12, alignment=TA_RIGHT)),
                ])
                row_types.append("subtotal")

            for eq in no_phase_eqs:
                cant   = int(eq.get("cant", 1) or 1)
                v_unit = int(eq.get("valor_unit", 0) or 0)
                sub    = cant * v_unit
                grand_total += sub
                cfg = (eq.get("config") or "").replace("\n", "<br/>")
                eq_data.append([
                    Paragraph(eq.get("equipo", "") or "", s_cell),
                    Paragraph(cfg, s_cell),
                    Paragraph(str(cant), s_cell_c),
                    Paragraph(f"<b>{_fmt_cop(sub)}</b>", s_cell_r),
                ])
                row_types.append("item")

            total_label = "TOTAL PROYECTO"
            eq_data.append([
                Paragraph(f"<b>{total_label}</b>",
                          ps("tot_lp", textColor=white, fontName="Helvetica-Bold",
                             fontSize=10, leading=14, alignment=TA_RIGHT)),
                "", "",
                Paragraph(f"<b>{_fmt_cop(grand_total)} COP</b>", s_cell_wr),
            ])
            row_types.append("total")

            ts_mp = [
                ("VALIGN",        (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING",   (0, 0), (-1, -1), 7),
                ("RIGHTPADDING",  (0, 0), (-1, -1), 7),
                ("TOPPADDING",    (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("GRID",          (0, 0), (-1, -1), 0.4, BORDER),
            ]
            for i, rt in enumerate(row_types):
                if rt == "header":
                    ts_mp.append(("BACKGROUND", (0, i), (-1, i), NAVY))
                elif rt == "phase_hdr":
                    ts_mp.extend([
                        ("BACKGROUND", (0, i), (-1, i), PHASE_BG),
                        ("SPAN",       (0, i), (-1, i)),
                        ("LEFTPADDING",(0, i), (-1, i), 12),
                        ("TOPPADDING", (0, i), (-1, i), 6),
                        ("BOTTOMPADDING", (0, i), (-1, i), 6),
                    ])
                elif rt == "subtotal":
                    ts_mp.extend([
                        ("BACKGROUND", (0, i), (-1, i), LIGHT_BLUE),
                        ("SPAN",       (0, i), (2,  i)),
                        ("ALIGN",      (3, i), (3,  i), "RIGHT"),
                        ("TOPPADDING", (0, i), (-1, i), 6),
                        ("BOTTOMPADDING", (0, i), (-1, i), 6),
                    ])
                elif rt == "total":
                    ts_mp.extend([
                        ("BACKGROUND",    (0, i), (-1, i), ORANGE),
                        ("SPAN",          (0, i), (2,  i)),
                        ("ALIGN",         (0, i), (0,  i), "RIGHT"),
                        ("TOPPADDING",    (0, i), (-1, i), 7),
                        ("BOTTOMPADDING", (0, i), (-1, i), 7),
                    ])
                else:
                    ts_mp.append(("BACKGROUND", (0, i), (-1, i),
                                  ACCENT if i % 2 == 0 else white))
            eq_t = Table(eq_data, colWidths=cw_e, repeatRows=1)
            eq_t.setStyle(TableStyle(ts_mp))

        else:
            eq_data = [e_hdr]
            total_oferta = 0
            for eq in equipos:
                cant   = int(eq.get("cant", 1) or 1)
                v_unit = int(eq.get("valor_unit", 0) or 0)
                sub    = cant * v_unit
                total_oferta += sub
                cfg = (eq.get("config") or "").replace("\n", "<br/>")
                eq_data.append([
                    Paragraph(eq.get("equipo", "") or "", s_cell),
                    Paragraph(cfg, s_cell),
                    Paragraph(str(cant), s_cell_c),
                    Paragraph(f"<b>{_fmt_cop(sub)}</b>", s_cell_r),
                ])
            eq_data.append([
                Paragraph("<b>TOTAL OFERTA</b>",
                          ps("tot_l", textColor=white, fontName="Helvetica-Bold",
                             fontSize=10, leading=14, alignment=TA_RIGHT)),
                "", "",
                Paragraph(f"<b>{_fmt_cop(total_oferta)} COP</b>", s_cell_wr),
            ])
            eq_t = Table(eq_data, colWidths=cw_e, repeatRows=1)
            eq_t.setStyle(_eq_table_style(eq_data, 2))

        story.append(_section("PROPUESTA ECONÓMICA", [eq_t]))

    # ── Stand-by table ────────────────────────────────────────────────────────
    if equipos:
        sb_hdr = [
            Paragraph("<b>Equipo</b>",                       s_cell_w),
            Paragraph("<b>Tiempo libre cargue</b>",          s_cell_w),
            Paragraph("<b>Tiempo libre descargue</b>",       s_cell_w),
            Paragraph("<b>Stand-By (por día)</b>",      s_cell_w),
        ]
        sb_data = [sb_hdr]
        for eq in equipos:
            sb_rate, sb_lc, sb_ld = _standby_for_equipo(
                eq.get("equipo", ""), eq.get("config", "")
            )
            sb_data.append([
                Paragraph(eq.get("equipo", "") or "", s_cell),
                Paragraph(sb_lc,   s_cell_c),
                Paragraph(sb_ld,   s_cell_c),
                Paragraph(f"<b>{sb_rate}</b>", s_cell_r),
            ])
        sb_ts = [
            ("BACKGROUND",    (0, 0), (-1, 0),  NAVY),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING",   (0, 0), (-1, -1), 7),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 7),
            ("TOPPADDING",    (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("GRID",          (0, 0), (-1, -1), 0.4, BORDER),
        ]
        for i in range(1, len(sb_data)):
            sb_ts.append(("BACKGROUND", (0, i), (-1, i), ACCENT if i % 2 == 0 else white))
        sb_t = Table(sb_data, colWidths=[W*0.34, W*0.22, W*0.22, W*0.22], repeatRows=1)
        sb_t.setStyle(TableStyle(sb_ts))
        sb_note = Paragraph(
            "* Las horas adicionales serán cobradas proporcionalmente según tarifa establecida."
            + (" El cobro del equipo de izaje inicia desde la llegada al sitio designado."
               if svc["has_izaje"] else ""),
            ps("sb_note", fontSize=7.5, leading=11, textColor=MED)
        )
        story.append(_section("STAND-BY", [sb_t, SP(4), sb_note]))

    # ── Izaje conditions ──────────────────────────────────────────────────────
    if svc["has_izaje"] and izaje_notas:
        story.append(_section("CONDICIONES TÉCNICAS DE IZAJE",
                               [note_para(n) for n in izaje_notas]))

    # ── Other technical notes ─────────────────────────────────────────────────
    if other_notas:
        story.append(_section("NOTAS TÉCNICAS DE OPERACIÓN",
                               [note_para(n) for n in other_notas]))

    # ── Commercial conditions ─────────────────────────────────────────────────
    cond_rows = [
        ("Forma de pago",   forma_pago),
        ("Póliza de carga", poliza_carga),
        ("Póliza RC",       poliza_rc),
        ("Resolución",      resolucion),
    ]
    cond_left  = [
        [Paragraph(f"<b>{lbl}</b>", ps("cl", fontName="Helvetica-Bold", fontSize=8,
                                        textColor=NAVY, leading=12)),
         Paragraph(val, ps("cv", fontSize=8, leading=12, textColor=MED))]
        for lbl, val in cond_rows
    ]
    cond_t = Table(cond_left, colWidths=[W*0.30, W*0.70])
    cond_ts = [
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 8),
        ("GRID",          (0, 0), (-1, -1), 0.4, BORDER),
    ]
    for i in range(len(cond_left)):
        cond_ts.append(("BACKGROUND", (0, i), (0, i), FGRAY))
        cond_ts.append(("BACKGROUND", (1, i), (1, i), ACCENT if i % 2 == 0 else white))
    cond_t.setStyle(TableStyle(cond_ts))
    story.append(_section("CONDICIONES COMERCIALES", [cond_t]))

    # ── Exclusions ────────────────────────────────────────────────────────────
    excls = [e.strip() for e in excl_raw.split("\n") if e.strip()]
    if excls:
        story.append(_section("EXCLUSIONES", [note_para(e) for e in excls]))

    # ── Signature block ───────────────────────────────────────────────────────
    story.append(SP(4))
    story.append(HRFlowable(width=W, color=BORDER, thickness=0.5))
    story.append(SP(10))
    # Two-column: each cell is a single Paragraph combining multiple lines via <br/>
    left_txt = (
        f"<b>{comercial}</b><br/>"
        f"<font color='#E8601C'><b>{cargo_str}</b></font><br/>"
        f"<font size='8'>BOOM Logistics Colombia S.A.S.</font><br/>"
        f"<font size='8' color='#5A6373'>{email_com}</font>"
    )
    right_txt = ""
    sig_t = Table(
        [[Paragraph(left_txt,  ps("sig_l", fontSize=9, leading=14, textColor=NAVY)),
          Paragraph(right_txt, ps("sig_r", fontSize=8.5, leading=13, textColor=MED))]],
        colWidths=[W * 0.48, W * 0.52],
    )
    sig_t.setStyle(TableStyle([
        ("VALIGN",       (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING",   (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 0),
    ]))
    story.append(sig_t)

    doc.build(story,
              onFirstPage=_page_footer,
              onLaterPages=_page_footer)
    return buf.getvalue()


def _extraer_info(texto: str) -> dict:
    result = {}
    tl = texto.lower()
    m = re.search(r"[\w.\-]+@[\w.\-]+\.\w+", texto)
    if m:
        result["email_cliente"] = m.group(0)
    for pat in [
        r"(?:empresa|cliente|para|de|señores?)[:\s]+([A-ZÁÉÍÓÚÑ][A-Za-záéíóúñ\s&.,]+?(?:S\.A\.S?\.?|LTDA\.?|S\.A\.)?)",
        r"^([A-ZÁÉÍÓÚÑ][A-Za-záéíóúñ\s]{4,35})\s*[\n,]",
    ]:
        m2 = re.search(pat, texto, re.IGNORECASE | re.MULTILINE)
        if m2:
            result["cliente"] = m2.group(1).strip().rstrip(",.")
            break
    for pat in [
        r"(?:desde|origen|salida)[:\s]+([A-Za-záéíóúñ\s]+?)(?:\s+hasta|\s+a\s|\s+hacia|\.|,|\n)",
        r"(?:puerto de|ciudad de)\s+([A-Za-záéíóúñ\s]+?)(?:\s+a\s|\.|,|\n)",
    ]:
        m3 = re.search(pat, texto, re.IGNORECASE)
        if m3:
            result["origen"] = m3.group(1).strip()
            break
    for pat in [r"(?:hasta|destino|hacia|a)\s+([A-Za-záéíóúñ\s,]+?)(?:\.|,|\n|$)"]:
        m4 = re.search(pat, texto, re.IGNORECASE)
        if m4:
            v = m4.group(1).strip()
            if len(v) > 2:
                result["destino"] = v
            break
    if any(x in tl for x in ["izaje", "grúa", "grua", "izamiento", "montaje"]):
        result["tipo"] = "SPOT MIXTO" if "transporte" in tl else "SPOT IZAJE"
    elif any(x in tl for x in ["transporte", "contenedor", "carga", "flete"]):
        result["tipo"] = "SPOT TRANSPORTE"
    frases = [f.strip() for f in re.split(r"[.\n]", texto) if len(f.strip()) > 25]
    if frases:
        result["descripcion"] = frases[0][:220]
    return result


# ── DB init ───────────────────────────────────────────────────────────────────
def _ensure_db():
    try:
        admin = pgdb.connect(host=DB_HOST, port=DB_PORT, database="postgres",
                              user=DB_USER, password=DB_PASSWORD)
        admin.autocommit = True
        cur = admin.cursor()
        cur.execute("SELECT 1 FROM pg_database WHERE datname = %s", (DB_NAME,))
        if not cur.fetchone():
            cur.execute(f'CREATE DATABASE "{DB_NAME}"')
            print(f"[DB] Base de datos '{DB_NAME}' creada.")
        admin.close()
    except Exception as e:
        print(f"[DB] Advertencia al crear BD: {e}")
    try:
        conn = pgdb.connect(host=DB_HOST, port=DB_PORT, database=DB_NAME,
                             user=DB_USER, password=DB_PASSWORD)
        conn.autocommit = True
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS ofertas (
                id bigint generated always as identity primary key,
                num text unique not null,
                mes text, fecha date, cliente text,
                realizada text, formalizada text, unidad text, tipo text,
                valor bigint default 0, estado text default 'ENVIADO',
                respuesta text, facturacion text,
                general text, seguimiento text,
                mes_aceptado text, fecha_facturacion date,
                valor_facturado bigint, no_factura text,
                created_at timestamptz default now()
            )
        """)
        # Migraciones: agrega columnas si la tabla ya existía sin ellas
        cur.execute("ALTER TABLE ofertas ADD COLUMN IF NOT EXISTS facturacion text")
        cur.execute("ALTER TABLE ofertas ADD COLUMN IF NOT EXISTS sector text")
        cur.execute("ALTER TABLE ofertas ADD COLUMN IF NOT EXISTS pdf_data jsonb")
        print("[DB] Tabla 'ofertas' lista.")

        cur.execute("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id        bigserial primary key,
                username  varchar(50) unique not null,
                nombre    varchar(100) not null,
                password_hash text not null,
                rol       varchar(20) not null default 'viewer',
                activo    boolean not null default true,
                creado_en timestamptz default now()
            )
        """)
        _admin_hash = _hash_pw("Boom2025*")
        cur.execute("""
            INSERT INTO usuarios (username, nombre, password_hash, rol)
            VALUES ('admin', 'Administrador', %s, 'admin')
            ON CONFLICT (username) DO NOTHING
        """, (_admin_hash,))
        # Migración: área en usuarios
        cur.execute("ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS area varchar(60)")

        cur.execute("""
            CREATE TABLE IF NOT EXISTS areas (
                id        bigserial primary key,
                nombre    varchar(100) unique not null,
                descripcion text,
                icono     varchar(10) default '🏢',
                activo    boolean not null default true,
                created_at timestamptz default now()
            )
        """)
        cur.execute("""
            INSERT INTO areas (nombre, descripcion, icono)
            VALUES ('Comercial', 'Área comercial — gestión de ofertas', '💼')
            ON CONFLICT (nombre) DO NOTHING
        """)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS area_permisos (
                area_id   bigint references areas(id) on delete cascade,
                modulo    varchar(50) not null,
                activo    boolean not null default true,
                primary key (area_id, modulo)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS notificaciones (
                id          bigserial primary key,
                oferta_id   bigint,
                oferta_num  text,
                cliente     text,
                origen      text,
                destino     text,
                valor       bigint default 0,
                leida       boolean default false,
                created_at  timestamptz default now()
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS osi (
                id           bigserial primary key,
                numero_osi   text unique not null,
                oferta_id    bigint,
                oferta_num   text,
                fecha        date default current_date,
                responsable  text,
                equipo       text,
                cliente      text,
                origen       text,
                destino      text,
                valor        bigint default 0,
                estado       text default 'PROGRAMADO',
                notas        text,
                created_at   timestamptz default now()
            )
        """)
        # Feature 3: OSI nuevas columnas
        cur.execute("ALTER TABLE osi ADD COLUMN IF NOT EXISTS fecha_despacho date")
        cur.execute("ALTER TABLE osi ADD COLUMN IF NOT EXISTS conductor text")
        cur.execute("ALTER TABLE osi ADD COLUMN IF NOT EXISTS placa text")
        cur.execute("ALTER TABLE osi ADD COLUMN IF NOT EXISTS observaciones text")

        # Feature 4: Historial de cambios
        cur.execute("""
            CREATE TABLE IF NOT EXISTS oferta_historial (
                id          bigserial primary key,
                oferta_id   bigint,
                oferta_num  text,
                campo       text,
                valor_ant   text,
                valor_nuevo text,
                usuario     text,
                created_at  timestamptz default now()
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS sessions (
                token      text primary key,
                user_id    bigint not null,
                username   text not null,
                nombre     text not null,
                rol        text not null,
                expires_at timestamptz not null,
                created_at timestamptz default now()
            )
        """)
        conn.close()
        print("[DB] Tablas 'areas', 'area_permisos', 'notificaciones', 'osi', 'oferta_historial' y 'sessions' listas.")
    except Exception as e:
        print(f"[DB] Error: {e}")
        raise

_ensure_db()

# ── Session helpers (DB-backed + memory cache) ────────────────────────────────
_sessions: dict = {}   # token -> {id, username, nombre, rol}

def _session_load_from_db():
    """Load all non-expired sessions from DB into memory cache on startup."""
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute(
                "SELECT token, user_id, username, nombre, rol FROM sessions "
                "WHERE expires_at > now()"
            )
            rows = fetchall(cur)
        for r in rows:
            _sessions[r["token"]] = {
                "id": r["user_id"], "username": r["username"],
                "nombre": r["nombre"], "rol": r["rol"],
            }
        print(f"[AUTH] {len(rows)} sesión(es) activa(s) cargada(s) desde DB.")
    except Exception as e:
        print(f"[AUTH] No se pudieron cargar sesiones desde DB: {e}")

def _session_save(token: str, user: dict, max_age_s: int = 86400 * 7):
    _sessions[token] = user
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute(
                """INSERT INTO sessions (token, user_id, username, nombre, rol, expires_at)
                   VALUES (%s, %s, %s, %s, %s, now() + interval '%s seconds')
                   ON CONFLICT (token) DO UPDATE
                   SET expires_at = now() + interval '%s seconds'""",
                (token, user["id"], user["username"], user["nombre"], user["rol"],
                 max_age_s, max_age_s)
            )
    except Exception as e:
        print(f"[AUTH] Error guardando sesión en DB: {e}")

def _session_delete(token: str):
    _sessions.pop(token, None)
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM sessions WHERE token = %s", (token,))
    except Exception as e:
        print(f"[AUTH] Error eliminando sesión de DB: {e}")

def _session_get(token: str) -> dict | None:
    """Check memory first, then DB (and repopulate memory on DB hit)."""
    if not token:
        return None
    if token in _sessions:
        return _sessions[token]
    # Fallback: look up DB (handles restarts where memory was cleared)
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute(
                "SELECT user_id, username, nombre, rol FROM sessions "
                "WHERE token = %s AND expires_at > now()",
                (token,)
            )
            row = fetchone(cur)
        if row:
            user = {"id": row["user_id"], "username": row["username"],
                    "nombre": row["nombre"], "rol": row["rol"]}
            _sessions[token] = user   # repopulate cache
            return user
    except Exception as e:
        print(f"[AUTH] Error consultando sesión en DB: {e}")
    return None

_session_load_from_db()


# ── Email notification ────────────────────────────────────────────────────────
def _enviar_notificacion_osi(oferta_row: dict, pdf_data: dict):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    if not SMTP_USER or not SMTP_PASSWORD:
        print("[EMAIL] SMTP_USER/SMTP_PASSWORD no configurados — email omitido")
        return

    num     = oferta_row.get("num", "")
    cliente = oferta_row.get("cliente", "")
    valor   = oferta_row.get("valor", 0) or 0
    origen  = pdf_data.get("origen", "") if pdf_data else ""
    destino = pdf_data.get("destino", "") if pdf_data else ""
    equipos = pdf_data.get("equipos", []) if pdf_data else []
    cargo_items = pdf_data.get("cargo_items", []) if pdf_data else []
    notas   = pdf_data.get("notas", "") if pdf_data else ""
    forma_pago = pdf_data.get("forma_pago", "50% anticipo / 50% a 30 días tras radicación de factura") if pdf_data else ""
    vigencia   = pdf_data.get("vigencia", 30) if pdf_data else 30

    # ── Tabla carga
    if cargo_items:
        rows_c = "".join(f"""<tr><td style='padding:6px 10px;border-bottom:1px solid #e5e5e5;'>{c.get('descripcion','')}</td>
          <td style='padding:6px 10px;border-bottom:1px solid #e5e5e5;text-align:center;'>{c.get('cant',1)}</td>
          <td style='padding:6px 10px;border-bottom:1px solid #e5e5e5;text-align:center;'>{c.get('dimensiones','')}</td>
          <td style='padding:6px 10px;border-bottom:1px solid #e5e5e5;text-align:center;'>{c.get('peso','')}</td></tr>"""
            for c in cargo_items if c.get('descripcion'))
        tabla_carga = f"""<table style='width:100%;border-collapse:collapse;font-size:13px;margin-top:8px;'>
          <thead><tr style='background:#1B2A4A;color:#fff;'>
            <th style='padding:7px 10px;text-align:left;'>Descripción</th>
            <th style='padding:7px 10px;'>Cant.</th>
            <th style='padding:7px 10px;'>Dimensiones</th>
            <th style='padding:7px 10px;'>Peso</th>
          </tr></thead><tbody>{rows_c}</tbody></table>"""
    else:
        tabla_carga = "<p style='color:#666;font-size:13px;'>Ver detalle en oferta adjunta.</p>"

    # ── Tabla equipos
    total = sum(int(e.get("cant",1) or 1) * int(e.get("valor_unit",0) or 0) for e in equipos)
    rows_e = "".join(f"""<tr><td style='padding:6px 10px;border-bottom:1px solid #e5e5e5;'>{e.get('equipo','')}</td>
      <td style='padding:6px 10px;border-bottom:1px solid #e5e5e5;text-align:center;'>{e.get('cant',1)}</td>
      <td style='padding:6px 10px;border-bottom:1px solid #e5e5e5;'>{e.get('config','')}</td>
      <td style='padding:6px 10px;border-bottom:1px solid #e5e5e5;text-align:right;font-weight:bold;'>{_fmt_cop(int(e.get('cant',1) or 1)*int(e.get('valor_unit',0) or 0))}</td></tr>"""
        for e in equipos if e.get('equipo'))
    total_row = f"""<tr style='background:#E8601C;color:#fff;'><td colspan='3' style='padding:8px 10px;font-weight:bold;'>TOTAL OFERTA</td>
      <td style='padding:8px 10px;text-align:right;font-weight:bold;'>{_fmt_cop(total)} COP</td></tr>"""
    tabla_equipos = f"""<table style='width:100%;border-collapse:collapse;font-size:13px;margin-top:8px;'>
      <thead><tr style='background:#1B2A4A;color:#fff;'>
        <th style='padding:7px 10px;text-align:left;'>Equipo</th>
        <th style='padding:7px 10px;'>Cant.</th>
        <th style='padding:7px 10px;text-align:left;'>Esquema de Seguridad</th>
        <th style='padding:7px 10px;text-align:right;'>Tarifa Neto (COP)</th>
      </tr></thead><tbody>{rows_e}{total_row}</tbody></table>"""

    # ── Stand-by desde notas
    sb_lines = [l.strip() for l in notas.split("\n") if "stand-by" in l.lower()]
    sb_html = "".join(f"<li style='margin-bottom:4px;'>{l}</li>" for l in sb_lines)
    if sb_html:
        sb_html = f"<ul style='margin:6px 0 0 0;padding-left:18px;font-size:13px;'>{sb_html}</ul>"

    # ── Condiciones operativas desde notas
    op_lines = [l.strip() for l in notas.split("\n")
                if l.strip() and "stand-by" not in l.lower()
                and not l.strip().lower().startswith(("origen:","destino:"))]
    op_html = "".join(f"<li style='margin-bottom:4px;'>{l}</li>" for l in op_lines)
    if op_html:
        op_html = f"<ul style='margin:6px 0 0 0;padding-left:18px;font-size:13px;'>{op_html}</ul>"

    html_body = f"""<!DOCTYPE html><html><head><meta charset='UTF-8'></head>
<body style='font-family:Arial,sans-serif;font-size:14px;color:#1B2A4A;margin:0;padding:20px;background:#f4f4f4;'>
<div style='max-width:700px;margin:0 auto;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.1);'>
  <div style='background:#1B2A4A;padding:14px 20px;display:flex;align-items:center;justify-content:space-between;'>
    <span style='color:#fff;font-weight:700;font-size:18px;letter-spacing:1px;'>BOOM</span>
    <div style='color:#fff;font-size:11px;text-align:right;'><strong style='font-size:12px;display:block;'>BOOM LOGISTICS COLOMBIA S.A.S.</strong>Soluciones de Transporte Especializado</div>
  </div>
  <div style='background:#E8601C;padding:8px 20px;'>
    <p style='margin:0;color:#fff;font-size:11px;font-weight:bold;'>🔔 NOTIFICACIÓN DE INICIO DE SERVICIO &nbsp;|&nbsp; REF: {num} &nbsp;|&nbsp; {cliente.upper()} &nbsp;|&nbsp; {(origen+' → '+destino).upper() if origen and destino else ''}</p>
  </div>
  <div style='padding:24px;'>
    <p style='margin:0 0 16px 0;'>Cordial Saludo,</p>
    <p style='margin:0 0 20px 0;'>Para conocimiento de todos, a continuación notificamos lo siguiente:</p>

    <div style='background:#f0f4ff;border-left:4px solid #1B2A4A;padding:12px 16px;border-radius:4px;margin-bottom:20px;'>
      <p style='margin:0 0 4px 0;font-weight:700;font-size:15px;'>📄 NOTIFICACIÓN DE INICIO DE SERVICIO</p>
      <p style='margin:2px 0;font-size:13px;'><strong>PARA:</strong> Operaciones, Seguimiento y Control Documental</p>
      <p style='margin:2px 0;font-size:13px;'><strong>REF:</strong> OFERTA MERCANTIL No. {num}</p>
      <p style='margin:2px 0;font-size:13px;'><strong>CLIENTE:</strong> {cliente.upper()}</p>
      <p style='margin:2px 0;font-size:13px;'><strong>ESTADO:</strong> 🟢 PROGRAMADO</p>
    </div>

    <p style='background:#1B2A4A;color:#fff;font-weight:700;font-size:12px;padding:7px 12px;border-radius:3px;margin:16px 0 8px 0;'>1. INFORMACIÓN DEL PROYECTO</p>
    <ul style='margin:0;padding-left:18px;font-size:13px;line-height:1.8;'>
      <li><strong>Origen:</strong> {origen or "A confirmar con cliente"}</li>
      <li><strong>Destino:</strong> {destino or "A confirmar con cliente"}</li>
      <li><strong>Fecha de cargue:</strong> A confirmar con cliente</li>
    </ul>

    <p style='background:#1B2A4A;color:#fff;font-weight:700;font-size:12px;padding:7px 12px;border-radius:3px;margin:16px 0 8px 0;'>2. DETALLE TÉCNICO DE LA CARGA</p>
    {tabla_carga}

    <p style='background:#1B2A4A;color:#fff;font-weight:700;font-size:12px;padding:7px 12px;border-radius:3px;margin:16px 0 8px 0;'>3. CONFIGURACIÓN DE TRANSPORTE Y TARIFA</p>
    {tabla_equipos}

    <p style='background:#1B2A4A;color:#fff;font-weight:700;font-size:12px;padding:7px 12px;border-radius:3px;margin:16px 0 8px 0;'>4. CONDICIONES OPERATIVAS Y SEGUROS</p>
    {sb_html}{op_html if op_html else "<p style='font-size:13px;color:#666;margin:6px 0 0 0;'>Póliza de carga y RCE hasta $4.000.000.000 COP. GPS incluido.</p>"}

    <p style='background:#1B2A4A;color:#fff;font-weight:700;font-size:12px;padding:7px 12px;border-radius:3px;margin:16px 0 8px 0;'>5. CONDICIONES COMERCIALES</p>
    <ul style='margin:0;padding-left:18px;font-size:13px;line-height:1.8;'>
      <li><strong>Forma de pago:</strong> {forma_pago}</li>
      <li><strong>Vigencia:</strong> {vigencia} días calendario</li>
    </ul>

    <p style='margin:20px 0 0 0;border-top:1px solid #e0e0e0;padding-top:14px;font-size:13px;color:#666;'>Lo anterior para lo correspondiente.</p>
    <p style='margin:8px 0 2px 0;font-weight:bold;color:#1B2A4A;'>Natalia Vargas</p>
    <p style='margin:2px 0;color:#E8601C;font-size:13px;'>Ejecutiva Comercial</p>
    <p style='margin:2px 0;font-size:12px;color:#666;'>BOOM Logistics Colombia S.A.S.</p>
  </div>
</div>
</body></html>"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"🔔 NOTIFICACIÓN DE INICIO DE SERVICIO — REF: {num} — {cliente.upper()}"
    msg["From"]    = SMTP_FROM
    msg["To"]      = ", ".join(NOTIF_TO)
    msg["Cc"]      = ", ".join(NOTIF_CC)
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=15) as server:
            server.ehlo()
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            all_recipients = NOTIF_TO + NOTIF_CC
            server.sendmail(SMTP_FROM, all_recipients, msg.as_string())
        print(f"[EMAIL] Notificación OSI enviada para oferta {num}")
    except Exception as exc:
        print(f"[EMAIL] Error al enviar notificación: {exc}")


def new_conn():
    conn = pgdb.connect(host=DB_HOST, port=DB_PORT, database=DB_NAME,
                         user=DB_USER, password=DB_PASSWORD)
    conn.autocommit = True
    return conn


@contextmanager
def get_conn():
    conn = new_conn()
    try:
        yield conn
    finally:
        conn.close()


def _cols(cursor):
    return [col[0] for col in cursor.description]


def fetchall(cursor):
    if not cursor.description:
        return []
    cols = _cols(cursor)
    return [_serialize(dict(zip(cols, row))) for row in cursor.fetchall()]


def fetchone(cursor):
    if not cursor.description:
        return None
    cols = _cols(cursor)
    row = cursor.fetchone()
    return _serialize(dict(zip(cols, row))) if row is not None else None


def _serialize(d: dict) -> dict:
    return {k: v.isoformat() if isinstance(v, (date, datetime)) else v
            for k, v in d.items()}


# ── App ───────────────────────────────────────────────────────────────────────
app = FastAPI(title="BOOM Logistics - Control de Ofertas")

_AUTH_PUBLIC = {"", "/", "/auth/login", "/auth/logout", "/auth/me", "/api/logo"}
_WRITE_METHODS = {"POST", "PUT", "DELETE", "PATCH"}


@app.middleware("http")
async def auth_middleware(request: Request, call_next):
    path = request.url.path
    if path in _AUTH_PUBLIC:
        return await call_next(request)

    token = request.cookies.get("boom_session")
    user = _session_get(token)

    if not user:
        return JSONResponse({"detail": "No autenticado"}, status_code=401)

    request.state.user = user

    if user["rol"] == "viewer" and request.method in _WRITE_METHODS:
        return JSONResponse({"detail": "Acceso de solo lectura"}, status_code=403)

    if path.startswith("/api/usuarios") and request.method in _WRITE_METHODS and user["rol"] != "admin":
        return JSONResponse({"detail": "Solo administradores pueden gestionar usuarios"}, status_code=403)

    return await call_next(request)


# ── Pydantic models ───────────────────────────────────────────────────────────
class OfertaCreate(BaseModel):
    num: Optional[str] = None
    mes: Optional[str] = None
    fecha: Optional[str] = None
    cliente: Optional[str] = None
    realizada: Optional[str] = None
    formalizada: Optional[str] = None
    unidad: Optional[str] = None
    tipo: Optional[str] = None
    sector: Optional[str] = None
    valor: Optional[int] = 0
    estado: Optional[str] = "ENVIADO"
    respuesta: Optional[str] = None
    facturacion: Optional[str] = None
    general: Optional[str] = None
    seguimiento: Optional[str] = None
    mes_aceptado: Optional[str] = None
    fecha_facturacion: Optional[str] = None
    valor_facturado: Optional[int] = None
    no_factura: Optional[str] = None
    pdf_data: Optional[dict] = None


class OfertaUpdate(BaseModel):
    mes: Optional[str] = None
    fecha: Optional[str] = None
    cliente: Optional[str] = None
    realizada: Optional[str] = None
    formalizada: Optional[str] = None
    unidad: Optional[str] = None
    tipo: Optional[str] = None
    sector: Optional[str] = None
    valor: Optional[int] = None
    estado: Optional[str] = None
    respuesta: Optional[str] = None
    facturacion: Optional[str] = None
    general: Optional[str] = None
    seguimiento: Optional[str] = None
    mes_aceptado: Optional[str] = None
    fecha_facturacion: Optional[str] = None
    valor_facturado: Optional[int] = None
    no_factura: Optional[str] = None
    pdf_data: Optional[dict] = None


class EquipoItem(BaseModel):
    equipo: Optional[str] = ""
    dimensiones: Optional[str] = ""
    cant: Optional[int] = 0
    config: Optional[str] = ""
    valor_unit: Optional[int] = 0


class CargoItem(BaseModel):
    descripcion: Optional[str] = ""
    tipo: Optional[str] = ""
    cant: Optional[int] = 1
    dimensiones: Optional[str] = ""
    peso: Optional[str] = ""
    volumen: Optional[str] = ""
    origen_detalle: Optional[str] = ""
    destino_detalle: Optional[str] = ""


class OfertaHtml(BaseModel):
    ref: Optional[str] = None
    cliente: Optional[str] = None
    contacto: Optional[str] = None
    email_cliente: Optional[str] = None
    ref_cliente: Optional[str] = None
    cliente_final: Optional[str] = None
    origen: Optional[str] = None
    destino: Optional[str] = None
    mes_anio: Optional[str] = None
    descripcion: Optional[str] = None
    comercial: Optional[str] = "Natalia Vargas"
    cargo_items: Optional[List[CargoItem]] = []
    equipos: Optional[List[EquipoItem]] = []
    notas: Optional[str] = ""
    forma_pago: Optional[str] = "50% anticipo / 50% a 30 días tras radicación de factura"
    vigencia: Optional[int] = 30
    poliza_carga: Optional[str] = "Hasta $4.000.000.000 COP por despacho"
    poliza_rc: Optional[str] = "Hasta $4.000.000.000 COP"
    resolucion: Optional[str] = "Carga extradimensionada/extrapesada incluida"
    exclusiones: Optional[str] = (
        "Permisos de tránsito, operación y pólizas asociadas (a cargo del cliente)\n"
        "Servicios, recursos o actividades no descritos explícitamente en esta oferta"
    )


class LoginBody(BaseModel):
    username: str
    password: str


class UsuarioCreate(BaseModel):
    username: str
    nombre: str
    password: str
    rol: str = "viewer"


class UsuarioUpdate(BaseModel):
    nombre: Optional[str] = None
    password: Optional[str] = None
    rol: Optional[str] = None
    activo: Optional[bool] = None


class TextoCliente(BaseModel):
    texto: str


class ChatMsg(BaseModel):
    role: str
    content: Union[str, List[Any]]


class ChatOfertaBody(BaseModel):
    messages: List[ChatMsg]


class ParseDetalle(BaseModel):
    texto: str


class AutoNotasRequest(BaseModel):
    equipos: Optional[List[EquipoItem]] = []
    texto_cliente: Optional[str] = ""
    origen: Optional[str] = ""
    destino: Optional[str] = ""


# ── Claude API extraction ─────────────────────────────────────────────────────
BOOM_SYSTEM_PROMPT = """ERES AUTOMATIZADOR DE OFERTAS COMERCIALES BOOM LOGISTICS.

PRINCIPIO FUNDAMENTAL: NO HAGAS PREGUNTAS. INTERPRETA DIRECTAMENTE COMO NATY LO HACE.
Si hay ambigüedad mínima, ASUME el estándar. Solo pregunta si falta información crítica.

CUANDO RECIBAS DATOS DE OFERTA (correo, mensaje, números):
1. GENERA EL JSON DE OFERTA INMEDIATAMENTE sin esperar confirmación
2. Incluye Stand-By obligatorio (con nota de horas proporcionales)
3. Aplica estructura fija BOOM
4. valor_unit: entero sin separadores (ej: 19500000). Sin precio → 0

ESTRUCTURA OFERTA (NO VARIAR):
- Header: azul #1B2A4A, logo base64, barra #E8601C
- Secciones: Técnico → Económico → Notas → Condiciones → Exclusiones → Firma
- Firma: "Natalia Vargas / Ejecutiva Comercial" (SIEMPRE)
- Stand-by: OBLIGATORIO
- Póliza: "Hasta $4.000.000.000 COP"
- Pago: "50% anticipo / 50% a 30 días" (salvo se indique otra cosa)

TARIFAS STAND-BY 2026:
- Cama Baja 3 ejes: $1.200.000/día (6h libre)
- Cama Baja 4 ejes: $1.800.000/día (8h libre)
- Cama Alta/Patineta 3 ejes: $1.200.000/día
- Semi Modular 5 ejes: $2.800.000/día
- Modular 4 Cuna 4: $8.500.000/día
- Modular 6 Cuna 6: $8.500.000/día
- Camión Turbo/Sencillo: $550.000/día

NOTAS OBLIGATORIAS EN "notas":
- Siempre: "Origen: [origen]\nDestino: [destino]\nEsquema de seguridad: \nStand-by [equipo]: $X.XXX.XXX/día. Las horas adicionales serán cobradas proporcionalmente según tarifa establecida.\nTiempos libres: 6 horas para cargue / 6 horas para descargue."
- Con izaje: agregar "El cobro del equipo de izaje inicia desde la llegada al sitio designado."
- Con skidding: agregar "Póliza de montaje: Para la operación de skidding se hace necesario expedir póliza específica de montaje."
- Modular/multi-punto: tiempos libres 12h en lugar de 6h
- Carga > 3.00 m ancho o extrapesada: agregar "2 Escoltas + 2 Tecnólogos" en config del equipo

REFERENCIAS: 26-0XXX (número consecutivo de 5 dígitos)

VOZ NATY:
- Formal, directo, sin bullets en textos
- ALL CAPS: códigos (26-0720), referencias (SOL #673)
- Cierres: "Quedamos atentos a cualquier inquietud"
- Párrafos cortos, concisión extrema

SOLO PREGUNTA SI:
- Falta cliente
- Falta equipo principal
- Algo es técnicamente inconsistente (ej: peso 100 ton en Cama Alta)
CUANDO DUDES → GENERA CON ESTÁNDARES, no pidas confirmación.
"""

def _extraer_info_claude(texto: str) -> dict:
    if not ANTHROPIC_OK:
        raise RuntimeError("Instala el paquete 'anthropic': pip install anthropic")
    if not ANTHROPIC_API_KEY:
        raise RuntimeError("Configura ANTHROPIC_API_KEY en el archivo .env")

    client = _anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    user_prompt = f"""Extrae del siguiente texto la información para la oferta BOOM y devuelve SOLO un JSON válido (sin markdown, sin texto extra):

TEXTO:
{texto}

JSON a devolver:
{{"cliente":"","contacto":"","email_cliente":"","ref_cliente":"","cliente_final":"","origen":"","destino":"","descripcion":"","cargo_items":[{{"descripcion":"","tipo":"","cant":1,"dimensiones":"","peso":"","volumen":"","origen_detalle":"","destino_detalle":""}}],"equipos":[{{"equipo":"","config":"","cant":1,"valor_unit":0}}],"notas":"Origen: ...\nDestino: ...\nEsquema de seguridad: \nStand-by [equipo]: $X.XXX.XXX/día. Las horas adicionales serán cobradas proporcionalmente según tarifa establecida.\nTiempos libres: 6 horas para cargue / 6 horas para descargue.","forma_pago":"50% anticipo / 50% a 30 días tras radicación de factura","vigencia":30}}"""

    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=2000,
        system=BOOM_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_prompt}],
    )
    raw = message.content[0].text.strip()
    if raw.startswith("```"):
        lines = raw.split("\n")
        inner = lines[1:-1] if lines[-1].strip() == "```" else lines[1:]
        raw = "\n".join(inner)
    return json.loads(raw.strip())


# ── Chat IA multi-turno ───────────────────────────────────────────────────────
CHAT_SYSTEM_PROMPT = BOOM_SYSTEM_PROMPT + """

════════════════════════════════════════
MODO CHAT — INSTRUCCIONES
════════════════════════════════════════
Eres el asistente personal de Natalia Vargas en BOOM Logistics. Conversas directamente con ella.
Responde SIEMPRE en español, de forma breve y profesional.

FORMATO DE RESPUESTA:
1. Responde conversacionalmente (1-4 oraciones según la complejidad).
2. Si tienes suficiente información para llenar el formulario de oferta, incluye al FINAL este bloque:

<<<DATOS>>>
{"cliente":"...","contacto":"...","email_cliente":"...","ref_cliente":"...","cliente_final":"...","origen":"...","destino":"...","descripcion":"...","cargo_items":[{"descripcion":"","tipo":"","cant":1,"dimensiones":"","peso":"","volumen":"","origen_detalle":"","destino_detalle":""}],"equipos":[{"equipo":"","config":"","cant":1,"valor_unit":0}],"notas":"...","forma_pago":"...","vigencia":30}
<<<FIN>>>

3. Si el usuario pide un ajuste puntual (ej: "cambia la vigencia", "agrega un escolta"), incluye el bloque JSON completo con el ajuste aplicado.
4. Si falta información clave, pídela conversacionalmente sin incluir el bloque JSON.
5. valor_unit: entero sin separadores (ej: 19500000). Sin precio → 0.
6. Aplica SIEMPRE las reglas de negocio BOOM: stand-by, escoltas, tiempos libres.
"""


def _chat_oferta(messages: list) -> dict:
    if not ANTHROPIC_OK:
        raise RuntimeError("Instala el paquete 'anthropic': pip install anthropic")
    if not ANTHROPIC_API_KEY:
        raise RuntimeError("Configura ANTHROPIC_API_KEY en el archivo .env")

    client = _anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    api_messages = [{"role": m.role, "content": m.content} for m in messages]

    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=2000,
        system=CHAT_SYSTEM_PROMPT,
        messages=api_messages,
    )
    raw = message.content[0].text.strip()

    fields = None
    reply = raw
    start = raw.find("<<<DATOS>>>")
    end = raw.find("<<<FIN>>>")
    if start != -1 and end != -1:
        json_str = raw[start + len("<<<DATOS>>>"):end].strip()
        reply = raw[:start].strip()
        try:
            fields = json.loads(json_str)
        except Exception:
            fields = None

    return {"reply": reply, "fields": fields}


# ── Routes ────────────────────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
def index():
    return FileResponse("templates/index.html")


@app.get("/api/logo")
def get_logo():
    return {"src": _logo_src()}


@app.get("/api/consecutivo")
def get_consecutivo():
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT COALESCE(MAX(CAST(num AS INTEGER)), 260000) + 1 AS next FROM ofertas")
            return {"consecutivo": fetchone(cur)["next"]}
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


@app.get("/api/ofertas")
def list_ofertas():
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM ofertas ORDER BY CAST(num AS INTEGER) DESC")
            return fetchall(cur)
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


@app.get("/api/ofertas/{oferta_id}")
def get_oferta(oferta_id: int):
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM ofertas WHERE id = %s", (oferta_id,))
            row = fetchone(cur)
            if row is None:
                raise HTTPException(404, "Oferta no encontrada")
            return row
    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


@app.post("/api/ofertas", status_code=201)
def create_oferta(oferta: OfertaCreate):
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            if not oferta.num:
                cur.execute("SELECT COALESCE(MAX(CAST(num AS INTEGER)), 260000) + 1 AS next FROM ofertas")
                oferta.num = str(fetchone(cur)["next"])
            pdf_json = json.dumps(oferta.pdf_data) if oferta.pdf_data else None
            cur.execute(
                """INSERT INTO ofertas
                   (num,mes,fecha,cliente,realizada,formalizada,unidad,tipo,sector,
                    valor,estado,respuesta,facturacion,general,seguimiento,mes_aceptado,
                    fecha_facturacion,valor_facturado,no_factura,pdf_data)
                   VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                   RETURNING *""",
                (oferta.num, oferta.mes, oferta.fecha or None, oferta.cliente,
                 oferta.realizada, oferta.formalizada, oferta.unidad, oferta.tipo,
                 oferta.sector,
                 oferta.valor or 0, oferta.estado or "ENVIADO", oferta.respuesta,
                 oferta.facturacion,
                 oferta.general, oferta.seguimiento, oferta.mes_aceptado,
                 oferta.fecha_facturacion or None, oferta.valor_facturado, oferta.no_factura,
                 pdf_json),
            )
            return fetchone(cur)
    except pgdb.DatabaseError as e:
        msg = str(e)
        if "unique" in msg.lower() or "duplicate" in msg.lower():
            raise HTTPException(409, "El número de oferta ya existe")
        traceback.print_exc()
        raise HTTPException(500, msg)
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


@app.patch("/api/ofertas/{oferta_id}")
def update_oferta(oferta_id: int, oferta: OfertaUpdate, request: Request):
    try:
        fields = {k: v for k, v in oferta.dict().items() if v is not None}
        if not fields:
            raise HTTPException(400, "No hay campos para actualizar")

        # Fields tracked for history
        TRACKED = {"respuesta", "estado", "valor", "facturacion", "mes_aceptado",
                   "seguimiento", "no_factura", "valor_facturado"}

        with get_conn() as conn:
            cur = conn.cursor()
            # Fetch current state before update — all tracked fields
            cur.execute("SELECT estado, respuesta, valor, facturacion, mes_aceptado, "
                        "seguimiento, no_factura, valor_facturado, num FROM ofertas WHERE id = %s",
                        (oferta_id,))
            prev = fetchone(cur)
            if prev is None:
                raise HTTPException(404, "Oferta no encontrada")

        if "pdf_data" in fields and isinstance(fields["pdf_data"], dict):
            fields["pdf_data"] = json.dumps(fields["pdf_data"])
        set_clause = ", ".join(f"{k} = %s" for k in fields)
        values = list(fields.values()) + [oferta_id]

        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute(f"UPDATE ofertas SET {set_clause} WHERE id = %s RETURNING *", values)
            row = fetchone(cur)
            if row is None:
                raise HTTPException(404, "Oferta no encontrada")

        # Feature 4: Insert historial for tracked fields that changed
        usuario = ""
        try:
            usuario = request.state.user.get("nombre", "") if hasattr(request.state, "user") else ""
        except Exception:
            pass
        oferta_num = prev.get("num", "")
        hist_entries = []
        for campo in TRACKED:
            if campo in fields:
                ant = str(prev.get(campo) or "")
                nuevo = str(fields[campo] or "")
                if ant != nuevo:
                    hist_entries.append((oferta_id, oferta_num, campo, ant, nuevo, usuario))
        if hist_entries:
            try:
                with get_conn() as conn_h:
                    cur_h = conn_h.cursor()
                    for entry in hist_entries:
                        cur_h.execute(
                            """INSERT INTO oferta_historial
                               (oferta_id, oferta_num, campo, valor_ant, valor_nuevo, usuario)
                               VALUES (%s, %s, %s, %s, %s, %s)""",
                            entry
                        )
            except Exception as he:
                print(f"[HISTORIAL] Error guardando historial: {he}")

        # ── Trigger notification when respuesta changes to ACEPTADA ──────────
        nueva_respuesta = (fields.get("respuesta") or "").upper()
        prev_respuesta  = (prev.get("respuesta") or "").upper()
        if nueva_respuesta == "ACEPTADA" and prev_respuesta != "ACEPTADA":
            import threading
            pdf_payload = row.get("pdf_data")
            if isinstance(pdf_payload, str):
                try:
                    pdf_payload = json.loads(pdf_payload)
                except Exception:
                    pdf_payload = {}
            pdf_payload = pdf_payload or {}

            # Persist notification + OSI in DB
            try:
                with get_conn() as conn2:
                    cur2 = conn2.cursor()
                    cur2.execute("""
                        INSERT INTO notificaciones (oferta_id, oferta_num, cliente, origen, destino, valor)
                        VALUES (%s,%s,%s,%s,%s,%s)
                    """, (oferta_id, row.get("num",""), row.get("cliente",""),
                          pdf_payload.get("origen",""), pdf_payload.get("destino",""),
                          row.get("valor",0) or 0))

                    # Auto-generate OSI number
                    cur2.execute("SELECT COUNT(*)+1 AS n FROM osi")
                    n = (fetchone(cur2) or {}).get("n", 1)
                    numero_osi = f"OSI-{datetime.now().year}-{str(n).zfill(4)}"
                    equipo_str = "; ".join(
                        e.get("equipo","") for e in (pdf_payload.get("equipos") or []) if e.get("equipo")
                    )
                    cur2.execute("""
                        INSERT INTO osi (numero_osi, oferta_id, oferta_num, responsable, equipo,
                                         cliente, origen, destino, valor)
                        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
                        ON CONFLICT (numero_osi) DO NOTHING
                    """, (numero_osi, oferta_id, row.get("num",""),
                          request.state.user.get("nombre","") if hasattr(request.state,"user") else "",
                          equipo_str, row.get("cliente",""),
                          pdf_payload.get("origen",""), pdf_payload.get("destino",""),
                          row.get("valor",0) or 0))
            except Exception as db_exc:
                print(f"[OSI] Error guardando notificación/OSI: {db_exc}")

            # Send email in background
            threading.Thread(
                target=_enviar_notificacion_osi,
                args=(row, pdf_payload),
                daemon=True
            ).start()

        return row
    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


@app.delete("/api/ofertas/{oferta_id}")
def delete_oferta(oferta_id: int):
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM ofertas WHERE id = %s RETURNING id", (oferta_id,))
            row = cur.fetchone()
            if row is None:
                raise HTTPException(404, "Oferta no encontrada")
            return {"deleted": oferta_id}
    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


@app.get("/api/ofertas/{oferta_id}/pdf")
def download_oferta_pdf(oferta_id: int):
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT pdf_data, num, cliente FROM ofertas WHERE id = %s", (oferta_id,))
            row = fetchone(cur)
        if row is None:
            raise HTTPException(404, "Oferta no encontrada")
        stored = row.get("pdf_data")
        if not stored:
            raise HTTPException(404, "Esta oferta no tiene datos de PDF guardados. Regenera el PDF desde Generar Oferta.")
        if isinstance(stored, str):
            payload = json.loads(stored)
        else:
            payload = stored
        payload["equipos"] = [e for e in (payload.get("equipos") or []) if e.get("equipo") or e.get("cant")]
        payload["cargo_items"] = [c for c in (payload.get("cargo_items") or []) if c.get("descripcion") or c.get("dimensiones")]
        pdf_bytes = generar_pdf_oferta(payload)
        ref_fmt = _fmt_ref(row.get("num") or oferta_id)
        cliente_slug = re.sub(r"[^a-zA-Z0-9]", "_", (row.get("cliente") or "BOOM"))[:20]
        filename = f"Oferta_{ref_fmt}_{cliente_slug}.pdf"
        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


@app.post("/api/extraer-info")
def extraer_info(body: TextoCliente):
    try:
        return _extraer_info(body.texto)
    except Exception as e:
        raise HTTPException(500, str(e))


@app.post("/api/extraer-info-claude")
def extraer_info_claude(body: TextoCliente):
    try:
        return _extraer_info_claude(body.texto)
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


@app.post("/api/chat-oferta")
def chat_oferta_endpoint(body: ChatOfertaBody):
    try:
        return _chat_oferta(body.messages)
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


_MAX_HISTORY   = 8       # turnos máximos de historial enviados a la API
_MAX_TEXT_CHARS = 6000   # caracteres máximos por bloque de texto adjunto

def _trim_api_messages(messages: list) -> list:
    """Recorta el historial a los últimos _MAX_HISTORY mensajes y trunca bloques de texto grandes."""
    msgs = messages[-_MAX_HISTORY:] if len(messages) > _MAX_HISTORY else messages

    def _trim_content(c):
        if isinstance(c, str):
            return c[:_MAX_TEXT_CHARS] + "\n[…truncado]" if len(c) > _MAX_TEXT_CHARS else c
        if isinstance(c, list):
            out = []
            for blk in c:
                if isinstance(blk, dict) and blk.get("type") == "text":
                    t = blk.get("text", "")
                    if len(t) > _MAX_TEXT_CHARS:
                        blk = {**blk, "text": t[:_MAX_TEXT_CHARS] + "\n[…truncado]"}
                out.append(blk)
            return out
        return c

    return [{"role": m["role"], "content": _trim_content(m["content"])} for m in msgs]


@app.post("/api/chat-oferta-stream")
def chat_oferta_stream(body: ChatOfertaBody):
    """Streaming version — devuelve tokens SSE en tiempo real."""
    def generate():
        if not ANTHROPIC_OK or not ANTHROPIC_API_KEY:
            yield f"data: {json.dumps({'err': 'API no configurada'})}\n\n"
            return
        try:
            client = _anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
            api_messages = _trim_api_messages(
                [{"role": m.role, "content": m.content} for m in body.messages]
            )

            full_text  = ""
            datos_seen = False

            with client.messages.stream(
                model="claude-haiku-4-5-20251001",
                max_tokens=2000,
                system=CHAT_SYSTEM_PROMPT,
                messages=api_messages,
            ) as stream:
                for chunk in stream.text_stream:
                    full_text += chunk
                    if not datos_seen:
                        if "<<<DATOS>>>" in full_text:
                            datos_seen = True
                            # emitir solo el texto previo al marcador
                            pre = full_text[:full_text.index("<<<DATOS>>>")]
                            # ya emitimos parte del texto en chunks anteriores;
                            # el frontend acumuló todo — no emitir más texto
                        else:
                            yield f"data: {json.dumps({'t': chunk})}\n\n"

            # Parsear campos al terminar
            fields = None
            reply  = full_text
            s = full_text.find("<<<DATOS>>>")
            e2 = full_text.find("<<<FIN>>>")
            if s != -1 and e2 != -1:
                reply = full_text[:s].strip()
                try:
                    fields = json.loads(full_text[s + len("<<<DATOS>>>"):e2].strip())
                except Exception:
                    fields = None

            yield f"data: {json.dumps({'done': True, 'reply': reply, 'fields': fields})}\n\n"

        except Exception as exc:
            traceback.print_exc()
            yield f"data: {json.dumps({'err': str(exc)})}\n\n"

    return StreamingResponse(
        generate(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.post("/api/parsear-detalle")
def parsear_detalle(body: ParseDetalle):
    try:
        return {"equipos": _parsear_detalle(body.texto)}
    except Exception as e:
        raise HTTPException(500, str(e))


@app.post("/api/auto-notas")
def api_auto_notas(data: AutoNotasRequest):
    try:
        equipos = [e.dict() for e in (data.equipos or [])]
        return {"notas": _auto_notas(
            equipos,
            data.texto_cliente or "",
            data.origen or "",
            data.destino or "",
        )}
    except Exception as e:
        raise HTTPException(500, str(e))


_PDF_PAGE_CSS = """
<style>
@page { size: A4; margin: 1.5cm 1.5cm 2cm 1.5cm; }
@media print {
  body { padding: 0 !important; background: white !important; }
  .wrapper { box-shadow: none !important; border-radius: 0 !important; max-width: 100% !important; }
}
</style>"""

# Ajustes para xhtml2pdf (no soporta CSS Grid ni Google Fonts externos)
_XHTML2PDF_CSS_FIXES = """
<style>
/* Reemplaza grid por tabla para spec-card */
.spec-grid { display: table; width: 100%; border-spacing: 6px; }
.spec-card  { display: table-cell; background: #f0f4fa; border-radius: 5px;
              padding: 9px 8px; text-align: center; width: 25%; }
.spec-card .val { font-size: 13px; font-weight: bold; color: #E8601C; }
.spec-card .lbl { font-size: 10px; color: #666; }
/* Fuentes del sistema como fallback */
body, table, p, ul, li { font-family: Helvetica, Arial, sans-serif !important; }
</style>"""


def _inject_pdf_css(html_str: str, extra_css: str = "") -> str:
    """Elimina fuentes externas e inyecta estilos de impresión."""
    # Quitar link de Google Fonts (no disponible sin red en xhtml2pdf)
    html_str = re.sub(r'<link[^>]+fonts\.googleapis\.com[^>]*>', '', html_str)
    return html_str.replace("</head>", _PDF_PAGE_CSS + extra_css + "\n</head>", 1)


def _html_to_pdf_bytes(html_str: str) -> bytes:
    """Convierte HTML → PDF. Prioridad: WeasyPrint > xhtml2pdf."""
    if WEASYPRINT_OK:
        prepared = _inject_pdf_css(html_str)
        return _weasyprint.HTML(string=prepared).write_pdf()

    if XHTML2PDF_OK:
        prepared = _inject_pdf_css(html_str, _XHTML2PDF_CSS_FIXES)
        buf = BytesIO()
        result = _pisa.CreatePDF(prepared, dest=buf, encoding="utf-8")
        if result.err:
            raise RuntimeError(f"xhtml2pdf: {result.err}")
        return buf.getvalue()

    raise RuntimeError("No hay motor PDF disponible. Instala weasyprint o xhtml2pdf.")


@app.post("/api/generar-pdf")
def generar_pdf(data: OfertaHtml):
    try:
        payload = data.dict()
        payload["equipos"] = [e for e in (payload.get("equipos") or [])
                               if e.get("equipo") or e.get("cant")]
        payload["cargo_items"] = [c for c in (payload.get("cargo_items") or [])
                                   if c.get("descripcion") or c.get("dimensiones")]

        ref_fmt = _fmt_ref(data.ref or "260001")
        cliente_slug = re.sub(r"[^a-zA-Z0-9]", "_", (data.cliente or "BOOM"))[:20]
        filename = f"Oferta_{ref_fmt}_{cliente_slug}.pdf"

        if WEASYPRINT_OK or XHTML2PDF_OK:
            # PDF generado desde el mismo HTML del correo — formato idéntico ✓
            html_str  = generar_html_oferta(payload)
            pdf_bytes = _html_to_pdf_bytes(html_str)
        else:
            # Fallback: PDF ReportLab (formato anterior)
            pdf_bytes = generar_pdf_oferta(payload)

        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


@app.post("/api/generar-html")
def generar_html(data: OfertaHtml):
    try:
        payload = data.dict()
        payload["equipos"] = [e for e in (payload.get("equipos") or [])
                               if e.get("equipo") or e.get("cant")]
        payload["cargo_items"] = [c for c in (payload.get("cargo_items") or [])
                                   if c.get("descripcion") or c.get("dimensiones")]
        html_str = generar_html_oferta(payload)
        ref_fmt = _fmt_ref(data.ref or "260001")
        cliente_slug = re.sub(r"[^a-zA-Z0-9]", "_", (data.cliente or "BOOM"))[:20]
        filename = f"Oferta_{ref_fmt}_{cliente_slug}.html"
        return Response(
            content=html_str.encode("utf-8"),
            media_type="text/html",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


@app.get("/api/stats")
def get_stats():
    try:
        with get_conn() as conn:
            def q(sql, params=None):
                cur = conn.cursor(); cur.execute(sql, params or ()); return fetchall(cur)
            def q1(sql, params=None):
                cur = conn.cursor(); cur.execute(sql, params or ()); return fetchone(cur)

            total     = q1("SELECT COUNT(*) AS total FROM ofertas")["total"]
            aceptadas = q1("SELECT COUNT(*) AS total FROM ofertas WHERE UPPER(respuesta)='ACEPTADA'")["total"]
            seguim    = q1("SELECT COUNT(*) AS total FROM ofertas WHERE UPPER(respuesta)='EN SEGUIMIENTO'")["total"]
            valor     = q1("SELECT COALESCE(SUM(valor),0) AS total FROM ofertas")["total"]
            meses_es  = {"January":"ENERO","February":"FEBRERO","March":"MARZO",
                          "April":"ABRIL","May":"MAYO","June":"JUNIO",
                          "July":"JULIO","August":"AGOSTO","September":"SEPTIEMBRE",
                          "October":"OCTUBRE","November":"NOVIEMBRE","December":"DICIEMBRE"}
            mes_es    = meses_es.get(datetime.now().strftime("%B"), "")
            este_mes  = q1("SELECT COUNT(*) AS total FROM ofertas WHERE mes=%s", (mes_es,))["total"]
            return {
                "total": total,
                "aceptadas": aceptadas,
                "tasa_aceptacion": round(aceptadas*100/total, 1) if total else 0,
                "seguimiento": seguim,
                "valor_total": valor,
                "este_mes": este_mes,
                "por_tipo":      q("SELECT tipo, COUNT(*) AS cnt FROM ofertas GROUP BY tipo ORDER BY cnt DESC"),
                "por_cliente":   q("SELECT cliente, COUNT(*) AS cnt FROM ofertas GROUP BY cliente ORDER BY cnt DESC LIMIT 10"),
                "por_mes":       q("SELECT mes, COUNT(*) AS cnt FROM ofertas WHERE mes IS NOT NULL GROUP BY mes"),
                "por_comercial": q("SELECT realizada, COUNT(*) AS cnt FROM ofertas WHERE realizada IS NOT NULL GROUP BY realizada ORDER BY cnt DESC"),
                "ultimas":       q("SELECT * FROM ofertas ORDER BY CAST(num AS INTEGER) DESC LIMIT 8"),
            }
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


# ── Auth endpoints ────────────────────────────────────────────────────────────
@app.post("/auth/login")
def login(body: LoginBody, response: Response):
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute(
                "SELECT id, username, nombre, password_hash, rol, activo FROM usuarios WHERE username = %s",
                (body.username.lower().strip(),)
            )
            row = fetchone(cur)
        if not row or not row.get("activo") or not _verify_pw(body.password, row["password_hash"]):
            raise HTTPException(401, "Usuario o contraseña incorrectos")
        token = secrets.token_hex(32)
        user_data = {"id": row["id"], "username": row["username"],
                     "nombre": row["nombre"], "rol": row["rol"]}
        _session_save(token, user_data, max_age_s=86400 * 7)
        response.set_cookie("boom_session", token, httponly=True, samesite="lax",
                            max_age=86400 * 7)
        return {"nombre": row["nombre"], "rol": row["rol"], "username": row["username"]}
    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


@app.post("/auth/logout")
def logout(request: Request, response: Response):
    token = request.cookies.get("boom_session")
    if token:
        _session_delete(token)
    response.delete_cookie("boom_session")
    return {"ok": True}


@app.get("/auth/me")
def me(request: Request):
    token = request.cookies.get("boom_session")
    user = _session_get(token)
    if not user:
        raise HTTPException(401, "No autenticado")
    return user


# ── Gestión de usuarios ───────────────────────────────────────────────────────
@app.get("/api/usuarios")
def list_usuarios(request: Request):
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, username, nombre, rol, activo, creado_en FROM usuarios ORDER BY id")
            return fetchall(cur)
    except Exception as e:
        raise HTTPException(500, str(e))


@app.post("/api/usuarios", status_code=201)
def create_usuario(body: UsuarioCreate, request: Request):
    if body.rol not in ("admin", "comercial", "operaciones", "viewer"):
        raise HTTPException(400, "Rol inválido")
    if request.state.user["rol"] != "admin":
        raise HTTPException(403, "Solo administradores pueden crear usuarios")
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO usuarios (username, nombre, password_hash, rol) VALUES (%s,%s,%s,%s) RETURNING id",
                (body.username.lower().strip(), body.nombre, _hash_pw(body.password), body.rol)
            )
            row = fetchone(cur)
            return {"id": row["id"], "username": body.username.lower(), "nombre": body.nombre,
                    "rol": body.rol, "activo": True}
    except Exception as e:
        if "unique" in str(e).lower():
            raise HTTPException(409, f'El usuario "{body.username}" ya existe')
        raise HTTPException(500, str(e))


@app.put("/api/usuarios/{uid}")
def update_usuario(uid: int, body: UsuarioUpdate, request: Request):
    if request.state.user["rol"] != "admin":
        raise HTTPException(403, "Solo administradores")
    sets, params = [], []
    if body.nombre   is not None: sets.append("nombre = %s");        params.append(body.nombre)
    if body.password is not None: sets.append("password_hash = %s"); params.append(_hash_pw(body.password))
    if body.rol      is not None:
        if body.rol not in ("admin", "comercial", "operaciones", "viewer"):
            raise HTTPException(400, "Rol inválido")
        sets.append("rol = %s"); params.append(body.rol)
    if body.activo   is not None: sets.append("activo = %s");        params.append(body.activo)
    if not sets:
        raise HTTPException(400, "Nada que actualizar")
    params.append(uid)
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute(f"UPDATE usuarios SET {', '.join(sets)} WHERE id = %s", params)
        return {"ok": True}
    except Exception as e:
        raise HTTPException(500, str(e))


@app.delete("/api/usuarios/{uid}")
def delete_usuario(uid: int, request: Request):
    if request.state.user["rol"] != "admin":
        raise HTTPException(403, "Solo administradores")
    if uid == request.state.user["id"]:
        raise HTTPException(400, "No puedes desactivar tu propia cuenta")
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("UPDATE usuarios SET activo = FALSE WHERE id = %s", (uid,))
        return {"ok": True}
    except Exception as e:
        raise HTTPException(500, str(e))


# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO TI
# ══════════════════════════════════════════════════════════════════════════════

class AreaCreate(BaseModel):
    nombre: str
    descripcion: Optional[str] = ""
    icono: Optional[str] = "🏢"

class AreaUpdate(BaseModel):
    nombre: Optional[str] = None
    descripcion: Optional[str] = None
    icono: Optional[str] = None
    activo: Optional[bool] = None

MODULOS_DISPONIBLES = ["dashboard", "generar", "control", "aprobadas", "operaciones", "tarifario"]


def _require_admin(request: Request):
    if request.state.user["rol"] != "admin":
        raise HTTPException(403, "Solo administradores TI")


@app.get("/api/ti/estado")
def ti_estado(request: Request):
    _require_admin(request)
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) AS total FROM ofertas")
            total_ofertas = fetchone(cur)["total"]
            cur.execute("SELECT COUNT(*) AS total FROM ofertas WHERE estado = 'APROBADO'")
            aprobadas = fetchone(cur)["total"]
            cur.execute("SELECT COUNT(*) AS total FROM usuarios WHERE activo = TRUE")
            usuarios_activos = fetchone(cur)["total"]
            cur.execute("SELECT COUNT(*) AS total FROM areas WHERE activo = TRUE")
            areas_activas = fetchone(cur)["total"]
        sesiones_activas = len(_sessions)
        return {
            "db": "ok",
            "sesiones_activas": sesiones_activas,
            "total_ofertas": total_ofertas,
            "ofertas_aprobadas": aprobadas,
            "usuarios_activos": usuarios_activos,
            "areas_activas": areas_activas,
            "sesiones": [
                {"nombre": v["nombre"], "rol": v["rol"]}
                for v in _sessions.values()
            ]
        }
    except Exception as e:
        return {"db": "error", "detalle": str(e)}


@app.get("/api/ti/areas")
def ti_list_areas(request: Request):
    _require_admin(request)
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT * FROM areas ORDER BY id")
        areas = fetchall(cur)
        for a in areas:
            cur.execute("SELECT COUNT(*) AS n FROM usuarios WHERE area = %s AND activo = TRUE", (a["nombre"],))
            a["usuarios_count"] = fetchone(cur)["n"]
            cur.execute("SELECT modulo, activo FROM area_permisos WHERE area_id = %s", (a["id"],))
            permisos = {r["modulo"]: r["activo"] for r in fetchall(cur)}
            a["permisos"] = {m: permisos.get(m, True) for m in MODULOS_DISPONIBLES}
        return areas


@app.post("/api/ti/areas", status_code=201)
def ti_create_area(body: AreaCreate, request: Request):
    _require_admin(request)
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO areas (nombre, descripcion, icono) VALUES (%s,%s,%s) RETURNING *",
                (body.nombre.strip(), body.descripcion or "", body.icono or "🏢")
            )
            area = fetchone(cur)
            for m in MODULOS_DISPONIBLES:
                cur.execute(
                    "INSERT INTO area_permisos (area_id, modulo, activo) VALUES (%s,%s,%s) ON CONFLICT DO NOTHING",
                    (area["id"], m, True)
                )
            area["permisos"] = {m: True for m in MODULOS_DISPONIBLES}
            area["usuarios_count"] = 0
            return area
    except Exception as e:
        if "unique" in str(e).lower():
            raise HTTPException(409, f'El área "{body.nombre}" ya existe')
        raise HTTPException(500, str(e))


@app.put("/api/ti/areas/{area_id}")
def ti_update_area(area_id: int, body: AreaUpdate, request: Request):
    _require_admin(request)
    sets, params = [], []
    if body.nombre      is not None: sets.append("nombre = %s");      params.append(body.nombre.strip())
    if body.descripcion is not None: sets.append("descripcion = %s"); params.append(body.descripcion)
    if body.icono       is not None: sets.append("icono = %s");       params.append(body.icono)
    if body.activo      is not None: sets.append("activo = %s");      params.append(body.activo)
    if not sets:
        raise HTTPException(400, "Nada que actualizar")
    params.append(area_id)
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute(f"UPDATE areas SET {', '.join(sets)} WHERE id = %s", params)
        return {"ok": True}
    except Exception as e:
        raise HTTPException(500, str(e))


@app.delete("/api/ti/areas/{area_id}")
def ti_delete_area(area_id: int, request: Request):
    _require_admin(request)
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT nombre FROM areas WHERE id = %s", (area_id,))
            row = fetchone(cur)
            if not row:
                raise HTTPException(404, "Área no encontrada")
            if row["nombre"] == "Comercial":
                raise HTTPException(400, "El área Comercial no se puede eliminar")
            cur.execute("UPDATE usuarios SET area = NULL WHERE area = %s", (row["nombre"],))
            cur.execute("DELETE FROM areas WHERE id = %s", (area_id,))
        return {"ok": True}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, str(e))


@app.put("/api/ti/areas/{area_id}/permisos")
def ti_update_permisos(area_id: int, permisos: dict, request: Request):
    _require_admin(request)
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            for modulo, activo in permisos.items():
                if modulo in MODULOS_DISPONIBLES:
                    cur.execute("""
                        INSERT INTO area_permisos (area_id, modulo, activo)
                        VALUES (%s,%s,%s)
                        ON CONFLICT (area_id, modulo) DO UPDATE SET activo = EXCLUDED.activo
                    """, (area_id, modulo, bool(activo)))
        return {"ok": True}
    except Exception as e:
        raise HTTPException(500, str(e))


@app.put("/api/ti/usuarios/{uid}/area")
def ti_asignar_area(uid: int, body: dict, request: Request):
    _require_admin(request)
    area = body.get("area")
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("UPDATE usuarios SET area = %s WHERE id = %s", (area or None, uid))
        return {"ok": True}
    except Exception as e:
        raise HTTPException(500, str(e))


# ── Notificaciones ────────────────────────────────────────────────────────────
@app.get("/api/notificaciones")
def get_notificaciones(request: Request):
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM notificaciones ORDER BY created_at DESC LIMIT 100")
            return fetchall(cur)
    except Exception as e:
        raise HTTPException(500, str(e))


@app.get("/api/notificaciones/count")
def count_notificaciones(request: Request):
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) AS total FROM notificaciones WHERE leida = false")
            row = fetchone(cur)
            return {"count": row["total"] if row else 0}
    except Exception as e:
        raise HTTPException(500, str(e))


@app.patch("/api/notificaciones/{nid}/leer")
def marcar_leida(nid: int, request: Request):
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("UPDATE notificaciones SET leida = true WHERE id = %s", (nid,))
        return {"ok": True}
    except Exception as e:
        raise HTTPException(500, str(e))


@app.patch("/api/notificaciones/leer-todas")
def marcar_todas_leidas(request: Request):
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("UPDATE notificaciones SET leida = true")
        return {"ok": True}
    except Exception as e:
        raise HTTPException(500, str(e))


# ── OSI ───────────────────────────────────────────────────────────────────────
class OSIUpdate(BaseModel):
    responsable:    Optional[str] = None
    equipo:         Optional[str] = None
    estado:         Optional[str] = None
    notas:          Optional[str] = None
    fecha:          Optional[str] = None
    fecha_despacho: Optional[str] = None
    conductor:      Optional[str] = None
    placa:          Optional[str] = None
    observaciones:  Optional[str] = None


@app.get("/api/osi")
def get_osi(request: Request):
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM osi ORDER BY created_at DESC")
            return fetchall(cur)
    except Exception as e:
        raise HTTPException(500, str(e))


@app.patch("/api/osi/{osi_id}")
def update_osi(osi_id: int, body: OSIUpdate, request: Request):
    try:
        fields = {k: v for k, v in body.dict().items() if v is not None}
        if not fields:
            raise HTTPException(400, "Sin campos")
        set_clause = ", ".join(f"{k} = %s" for k in fields)
        values = list(fields.values()) + [osi_id]
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute(f"UPDATE osi SET {set_clause} WHERE id = %s RETURNING *", values)
            row = fetchone(cur)
            if not row:
                raise HTTPException(404, "OSI no encontrada")
            return row
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, str(e))


# ── Feature 4: Historial de cambios ──────────────────────────────────────────
@app.get("/api/ofertas/{oferta_id}/historial")
def get_oferta_historial(oferta_id: int):
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute(
                "SELECT * FROM oferta_historial WHERE oferta_id = %s ORDER BY created_at DESC",
                (oferta_id,)
            )
            return fetchall(cur)
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


# ── Feature 2: Mini CRM Clientes ──────────────────────────────────────────────
@app.get("/api/clientes/stats")
def get_clientes_stats():
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT
                    cliente,
                    COUNT(*) AS total_ofertas,
                    COUNT(*) FILTER (WHERE UPPER(respuesta) = 'ACEPTADA') AS aceptadas,
                    COUNT(*) FILTER (WHERE UPPER(respuesta) = 'RECHAZADA') AS rechazadas,
                    COUNT(*) FILTER (WHERE UPPER(respuesta) = 'EN SEGUIMIENTO') AS en_seguimiento,
                    ROUND(
                        COUNT(*) FILTER (WHERE UPPER(respuesta) = 'ACEPTADA') * 100.0
                        / NULLIF(COUNT(*), 0), 1
                    ) AS tasa_cierre,
                    COALESCE(SUM(valor), 0) AS valor_total_ofertado,
                    COALESCE(SUM(valor_facturado), 0) AS valor_total_facturado,
                    MAX(fecha) AS ultima_oferta
                FROM ofertas
                WHERE cliente IS NOT NULL AND cliente != ''
                GROUP BY cliente
                ORDER BY total_ofertas DESC
            """)
            return fetchall(cur)
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


# ── Feature 1: Exportar a Excel ───────────────────────────────────────────────
def _make_excel_ofertas(rows: list) -> bytes:
    if not OPENPYXL_OK:
        raise RuntimeError("openpyxl no está instalado. Ejecuta: pip install openpyxl")

    NAVY  = "1B2A4A"
    ALT   = "EEF2FF"

    def _v(val):
        """Convierte cualquier valor a algo seguro para openpyxl."""
        if val is None:
            return ""
        if isinstance(val, (dict, list)):
            return ""
        return val

    wb = Workbook()

    # ── Hoja 1: Ofertas ───────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Ofertas"
    ws.freeze_panes = "A2"

    hdr_fill = PatternFill("solid", fgColor=NAVY)
    alt_fill = PatternFill("solid", fgColor=ALT)
    hdr_font = Font(color="FFFFFF", bold=True, size=10)
    center   = Alignment(horizontal="center", vertical="center", wrap_text=False)

    headers    = ["No. Oferta","Mes","Fecha","Cliente","Realizada por",
                  "Unidad","Tipo","Valor COP","Respuesta","Estado",
                  "Seguimiento","Mes Aceptado","Facturación","Valor Facturado","No. Factura"]
    col_widths = [12, 12, 13, 30, 22, 14, 22, 18, 16, 14, 16, 14, 14, 18, 16]

    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.fill = hdr_fill
        c.font = hdr_font
        c.alignment = center
        ws.column_dimensions[c.column_letter].width = w
    ws.row_dimensions[1].height = 22

    for ri, o in enumerate(rows, 2):
        vals = [
            _v(o.get("num")),      _v(o.get("mes")),     _v(o.get("fecha")),
            _v(o.get("cliente")),  _v(o.get("realizada")),
            _v(o.get("unidad")),   _v(o.get("tipo")),
            o.get("valor") or 0,
            _v(o.get("respuesta")), _v(o.get("estado")),
            _v(o.get("seguimiento")), _v(o.get("mes_aceptado")),
            _v(o.get("facturacion")),
            o.get("valor_facturado") or 0,
            _v(o.get("no_factura")),
        ]
        fill = alt_fill if ri % 2 == 0 else None
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=ri, column=ci, value=v)
            c.alignment = Alignment(vertical="center")
            if fill:
                c.fill = fill

    # ── Hoja 2: Resumen ───────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Resumen")
    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 14
    ws2.column_dimensions["C"].width = 20

    def _hdr2(row, col, text):
        c = ws2.cell(row=row, column=col, value=text)
        c.fill = hdr_fill
        c.font = hdr_font
        c.alignment = center

    def _title2(row, text):
        c = ws2.cell(row=row, column=1, value=text)
        c.font = Font(bold=True, size=11, color=NAVY)

    # Bloque 1 — por respuesta
    _title2(1, "RESUMEN POR RESPUESTA")
    _hdr2(2, 1, "Respuesta")
    _hdr2(2, 2, "Cantidad")

    resp_map: dict = {}
    for o in rows:
        k = (o.get("respuesta") or "Sin respuesta").upper()
        resp_map[k] = resp_map.get(k, 0) + 1
    for ri2, (k, v) in enumerate(sorted(resp_map.items()), 3):
        ws2.cell(row=ri2, column=1, value=k)
        ws2.cell(row=ri2, column=2, value=v)

    # Bloque 2 — por mes
    off = 3 + len(resp_map) + 2
    _title2(off, "RESUMEN POR MES")
    _hdr2(off + 1, 1, "Mes")
    _hdr2(off + 1, 2, "Cantidad")
    _hdr2(off + 1, 3, "Valor Total COP")

    mes_order = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
                 "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
    mes_cnt: dict = {}
    mes_val: dict = {}
    for o in rows:
        m = (o.get("mes") or "").upper()
        if m:
            mes_cnt[m] = mes_cnt.get(m, 0) + 1
            mes_val[m] = mes_val.get(m, 0) + (o.get("valor") or 0)
    for ri2, m in enumerate([m for m in mes_order if m in mes_cnt], off + 2):
        ws2.cell(row=ri2, column=1, value=m)
        ws2.cell(row=ri2, column=2, value=mes_cnt[m])
        ws2.cell(row=ri2, column=3, value=mes_val[m])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


@app.get("/api/exportar/ofertas")
def exportar_ofertas_excel(
    mes:    Optional[str] = Query(None),
    anio:   Optional[int] = Query(None),
    estado: Optional[str] = Query(None),
):
    try:
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM ofertas ORDER BY CAST(num AS INTEGER) DESC")
            rows = fetchall(cur)

        # Filtrar
        if mes:
            rows = [r for r in rows if (r.get("mes") or "").upper() == mes.upper()]
        if anio:
            rows = [r for r in rows if r.get("fecha") and str(r["fecha"])[:4] == str(anio)]
        if estado:
            rows = [r for r in rows if (r.get("respuesta") or "").upper() == estado.upper()
                    or (r.get("estado") or "").upper() == estado.upper()]

        excel_bytes = _make_excel_ofertas(rows)
        fname = "Ofertas_BOOM"
        if mes:
            fname += f"_{mes}"
        if anio:
            fname += f"_{anio}"
        fname += ".xlsx"
        return StreamingResponse(
            BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{fname}"'},
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))


# ── Feature 5: Informe Mensual de Facturación ─────────────────────────────────
@app.get("/api/reportes/facturacion")
def reporte_facturacion(anio: Optional[int] = Query(None)):
    try:
        anio_val = anio or datetime.now().year
        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT
                    mes_aceptado AS mes,
                    COUNT(*) AS total_aceptadas,
                    COUNT(*) FILTER (WHERE no_factura IS NOT NULL AND no_factura != '') AS facturadas,
                    COUNT(*) FILTER (WHERE no_factura IS NULL OR no_factura = '') AS sin_facturar,
                    COALESCE(SUM(valor), 0) AS valor_ofertado,
                    COALESCE(SUM(valor_facturado), 0) AS valor_facturado,
                    ROUND(
                        COALESCE(SUM(valor_facturado), 0) * 100.0
                        / NULLIF(COALESCE(SUM(valor), 0), 0), 1
                    ) AS tasa_facturacion
                FROM ofertas
                WHERE UPPER(respuesta) = 'ACEPTADA'
                  AND mes_aceptado IS NOT NULL AND mes_aceptado != ''
                  AND (
                      fecha_facturacion IS NULL
                      OR EXTRACT(YEAR FROM fecha_facturacion) = %s
                      OR fecha IS NULL
                      OR EXTRACT(YEAR FROM fecha) = %s
                  )
                GROUP BY mes_aceptado
                ORDER BY
                    CASE mes_aceptado
                        WHEN 'ENERO'      THEN 1  WHEN 'FEBRERO'    THEN 2
                        WHEN 'MARZO'      THEN 3  WHEN 'ABRIL'      THEN 4
                        WHEN 'MAYO'       THEN 5  WHEN 'JUNIO'      THEN 6
                        WHEN 'JULIO'      THEN 7  WHEN 'AGOSTO'     THEN 8
                        WHEN 'SEPTIEMBRE' THEN 9  WHEN 'OCTUBRE'    THEN 10
                        WHEN 'NOVIEMBRE'  THEN 11 WHEN 'DICIEMBRE'   THEN 12
                        ELSE 99
                    END
            """, (anio_val, anio_val))
            return fetchall(cur)
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, str(e))

