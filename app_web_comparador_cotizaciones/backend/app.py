
from fastapi import FastAPI, File, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from typing import List
from pathlib import Path
import io, re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

app = FastAPI(title="Comparador Cotizaciones V3.5 - Lector Universal")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# =========================
# UTILIDADES
# =========================

def clean(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()

def norm(s):
    s = clean(s).lower()
    for a, b in [("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u"),("²","2")]:
        s = s.replace(a, b)
    return s

def parse_num(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^0-9,.\-]", "", str(v or "").strip())
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    try:
        return float(s)
    except Exception:
        return None

def format_group(s):
    s = clean(s)
    s = s.replace("MM2","mm²").replace("mm2","mm²").replace("MM²","mm²")
    s = s.replace(",", ".").replace("X", "x")
    return re.sub(r"\s+", "", s)

def detect_group(text):
    t = clean(text)
    m = re.search(r"(\d+\s*[xX]\s*\d+(?:[\.,]\d+)?(?:\s*\+\s*B?\d+)?\s*(?:mm2|MM2|mm²|MM²)?)", t)
    if m:
        return format_group(m.group(1))

    maps = {
        "EC 0210": "2x1mm²", "EC 0215": "2x1.5mm²", "EC 0307": "3x0.75mm²",
        "EC 0410": "4x1mm²", "NF 11500": "1x150mm²", "OF 1210": "12x1mm²",
        "NF 0215": "2x1.5mm²", "NF 0225": "2x2.5mm²", "NF 0315": "3x1.5mm²",
        "NF 0325": "3x2.5mm²", "NF 0425": "4x2.5mm²", "NF 0440": "4x4mm²",
        "OF 0715": "7x1.5mm²", "VK 1160": "1x16mm²", "VK 0125": "1x2.5mm²",
        "1X150": "1x150mm²", "12X1": "12x1mm²", "2X1.5": "2x1.5mm²",
        "2X2.5": "2x2.5mm²", "3X1.5": "3x1.5mm²", "3X2.5": "3x2.5mm²",
        "4X2.5": "4x2.5mm²", "4X4": "4x4mm²", "7X1.5": "7x1.5mm²",
        "1X16": "1x16mm²", "1X2.5": "1x2.5mm²"
    }
    u = t.upper()
    for k, v in maps.items():
        if k in u:
            return v
    return ""

def detect_provider(filename, text):
    low = text.lower()
    if "ingeniería boggio" in low or "ingenieria boggio" in low:
        return "Ingeniería Boggio"
    if "marlew" in low:
        return "Marlew"
    if "ateco cables" in low:
        return "Ateco"

    # proveedor genérico: intenta tomar una línea superior con nombre
    lines = [clean(x) for x in text.splitlines() if clean(x)]
    for line in lines[:12]:
        l = norm(line)
        if any(x in l for x in ["presupuesto", "cotizacion", "cotización", "fecha", "cliente", "cuit"]):
            continue
        if len(line) >= 4 and len(line) <= 60:
            return line[:60]
    return Path(filename).stem

def detect_quote(filename, text):
    patterns = [
        r"N[úu]mero:\s*([A-Z0-9\-]+)",
        r"PRESUPUESTO\s+([0-9]+)",
        r"Presupuesto:\s*([0-9]+)",
        r"Cotizaci[oó]n\s*[:#]?\s*([A-Z0-9\-]+)",
        r"Oferta\s*[:#]?\s*([A-Z0-9\-]+)",
        r"PR\s*([0-9]+)"
    ]
    for pat in patterns:
        m = re.search(pat, text, re.I)
        if m:
            return m.group(1)
    return Path(filename).stem

def pdf_pages(data):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        return [p.extract_text(x_tolerance=1, y_tolerance=3) or "" for p in pdf.pages]

def make_item(filename, prov, cot, nro, codigo, marca, desc, grupo, unidad, cant, punit, subtotal, notas="", minimo="", entrega="", parser=""):
    subtotal = subtotal if subtotal is not None else (cant or 0) * (punit or 0)
    iva = subtotal * 0.21
    return {
        "archivo": filename,
        "proveedor": prov,
        "cotizacion": cot,
        "nro_item": str(nro),
        "codigo": codigo or "",
        "codigo_interno": "",
        "marca": marca or "",
        "descripcion": clean(desc),
        "formacion": grupo or "",
        "grupo_comparable": grupo or codigo or "",
        "moneda": "USD",
        "unidad": unidad or "u",
        "cantidad_pedida": cant or 0,
        "cantidad_real": cant or 0,
        "precio_unitario": punit or 0,
        "subtotal_sin_iva": subtotal or 0,
        "iva_pct": 21,
        "iva_monto": iva,
        "total_con_iva": subtotal + iva,
        "minimo_compra": minimo,
        "venta_fraccionada": "",
        "entrega": entrega,
        "notas": notas,
        "parser": parser,
        "validado": False
    }

# =========================
# LECTORES ESPECÍFICOS
# =========================

def parse_boggio(filename, pages):
    text = "\n".join(pages)
    cot = detect_quote(filename, text)
    entrega = "15/05/2026" if "Fecha de Entrega: 15/05/2026" in text else "Consultar"
    lines = [clean(x) for x in text.splitlines() if clean(x)]
    blocks, cur = [], []
    start = re.compile(r"^\d+\s+\d{3,6}\s+\[[^\]]+\]\s+", re.I)
    ignore = ("X ","Doc. no","Ingeniería Boggio","Remedios","Argentina -","www.","Cliente:","L De La Torre","Cond. IVA","CUIT:","Comercial:","Vencimiento:","Términos","Referencia","% de","N° Imagen","Subtotal USD","IVA 21%","Total USD","JUANI","PRECIOS NETOS","NETO PAGO","Su pedido","Para más","https:","Fecha de Entrega","Los precios","Las garantias","(*)Confirmar","Ud. fue","Página:")
    for ln in lines:
        if ln.startswith(ignore):
            continue
        if start.match(ln):
            if cur:
                blocks.append(" ".join(cur))
            cur = [ln]
        elif cur:
            cur.append(ln)
    if cur:
        blocks.append(" ".join(cur))

    out = []
    brand_re = r"\b(INDECA|WENTINCK|FONSECA|IMSA|PRYSMIAN)\b"
    for b in blocks:
        cm = re.match(r"^(\d+)\s+(\d{3,6})\s+\[([^\]]+)\]\s+(.+)$", b, re.I)
        if not cm:
            continue
        nro, img, cod, rest = cm.groups()
        m = re.search(brand_re + r"\s*(?:\(\*\))?\s+([\d\.,]+)\s*(m|mt|mts|Unidad|unidad|un|u)?\s+([\d\.,]+)\s+IVA\s*21%\s+([\d\.,]+)", rest, re.I)
        if not m:
            continue
        marca = m.group(1).upper()
        cant = parse_num(m.group(2)) or 1
        unidad = "m" if clean(m.group(3)).lower().startswith("m") else "u"
        punit = parse_num(m.group(4)) or 0
        subt = parse_num(m.group(5)) or cant * punit
        desc = clean(rest[:m.start()]).replace("()", "").strip()
        out.append(make_item(filename, "Ingeniería Boggio", cot, nro, cod, marca, desc, detect_group(desc), unidad, cant, punit, subt, "Confirmar stock" if "(*)" in b else "", "No informado", entrega, "boggio"))
    return out

def parse_marlew(filename, pages):
    text = "\n".join(pages)
    cot = detect_quote(filename, text)
    lines = [clean(x) for x in text.splitlines() if clean(x)]
    starts = [i for i,l in enumerate(lines) if re.match(r"^\d+\s+\d{2,5}\s+MT\s+C[oó]digo:", l, re.I)]
    blocks = [" ".join(lines[s: starts[j+1] if j+1 < len(starts) else len(lines)]) for j, s in enumerate(starts)]
    out = []
    for b in blocks:
        m = re.search(r"^(\d+)\s+(\d{2,5})\s+MT\s+C[oó]digo:\s*(.*?)\s+Formaci[oó]n:\s*([^\s]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)", b, re.I)
        if not m:
            continue
        nro, cant_raw, cod, form, punit_raw, net, total_raw = m.groups()
        cant = parse_num(cant_raw) or 1
        punit = parse_num(punit_raw) or 0
        subt = parse_num(total_raw) or cant * punit
        grupo = detect_group(form)
        low = norm(b)
        entrega = "6/8 semanas" if "6/8 semanas" in low else "A confirmar"
        minimo = ""
        notas = []
        if "minimo de provision" in low:
            minimo = f"Mín. {int(cant)} m"
        if "unica bobina" in low or "no fraccionable" in low:
            minimo = f"Bobina única {int(cant)} m"
            notas.append("No fraccionable")
        if "material en stock" in low:
            notas.append("En stock salvo venta")
        out.append(make_item(filename, "Marlew", cot, nro, cod, "MARLEW", cod, grupo, "m", cant, punit, subt, "; ".join(notas), minimo, entrega, "marlew"))
    return out

def parse_ateco(filename, pages):
    text = "\n".join(pages)
    cot = detect_quote(filename, text)
    entrega = "5 días" if "Plazo de entrega: 5 dias" in text else "Consultar"
    out = []
    for ln in [clean(x) for x in text.splitlines() if clean(x)]:
        m = re.match(r"^([0-9]{4}-[0-9]-[0-9]{6}-[0-9])\s+(\d+)\s+(.+?)\s+sa\s+([\d\.,]+)\s+([\d\.,]+)$", ln, re.I)
        if not m:
            continue
        cod, cant_raw, desc, punit_raw, sub_raw = m.groups()
        cant = parse_num(cant_raw) or 1
        punit = parse_num(punit_raw) or 0
        subt = parse_num(sub_raw) or cant * punit
        out.append(make_item(filename, "Ateco", cot, len(out)+1, cod, "ATECO", desc, detect_group(desc), "m", cant, punit, subt, "Contado anticipo", "No informado", entrega, "ateco"))
    return out

# =========================
# LECTOR UNIVERSAL
# =========================

BAD_WORDS = [
    "cuit", "iva responsable", "ingresos brutos", "domicilio", "tel", "telefono", "mail", "email", "www",
    "cliente", "fecha", "presupuesto", "cotizacion", "cotización", "validez", "condicion", "condición",
    "forma de pago", "plazo de pago", "total", "subtotal", "observacion", "observación", "pagina", "página",
    "banco", "cbu", "alias", "vendedor", "remito", "factura"
]

TECH_WORDS = [
    "cable", "caño", "cano", "tubo", "válvula", "valvula", "brida", "codo", "tee", "reducción", "reduccion",
    "cupla", "niple", "curva", "acople", "chapa", "perfil", "aislación", "aislacion", "mm", "pulg", "sch",
    "ansi", "astm", "inox", "acero", "cobre", "pvc", "xlpe", "blindado", "pantalla", "motor", "bomba",
    "sensor", "presostato", "termómetro", "termometro", "manómetro", "manometro", "tablero", "borne",
    "interruptor", "contactor", "disyuntor"
]

def looks_like_item(line):
    l = norm(line)
    if len(line) < 15:
        return False
    if any(b in l for b in BAD_WORDS):
        return False
    nums = re.findall(r"\d+(?:[\.,]\d+)?", line)
    if len(nums) < 2:
        return False
    if any(w in l for w in TECH_WORDS):
        return True
    if detect_group(line):
        return True
    return False

def parse_universal(filename, pages):
    text = "\n".join(pages)
    prov = detect_provider(filename, text)
    cot = detect_quote(filename, text)
    lines = [clean(x) for x in text.splitlines() if clean(x)]
    out = []
    buffer = []

    def flush_candidate(candidate):
        if not looks_like_item(candidate):
            return
        nums = re.findall(r"\d+(?:[\.,]\d+)?", candidate)
        parsed = [parse_num(x) for x in nums if parse_num(x) is not None]
        if len(parsed) < 2:
            return

        # heurística universal:
        # último número = subtotal o precio
        # anteúltimo = precio unitario
        # primer número razonable = cantidad
        subtotal = parsed[-1]
        punit = parsed[-2] if len(parsed) >= 2 else subtotal
        cant = round(subtotal / punit, 4) if punit and punit > 0 else 1

        # si detecta cantidad al inicio, usa esa
        m_qty = re.match(r"^\D*(\d+(?:[\.,]\d+)?)\s*(m|mt|mts|un|u|unidad|kg)?\b", candidate, re.I)
        if m_qty:
            q = parse_num(m_qty.group(1))
            if q and q > 0:
                cant = q

        grupo = detect_group(candidate)
        codigo = ""
        m_code = re.search(r"\b([A-Z]{1,4}[- ]?\d{2,8}|[0-9]{3,8}[-/][0-9A-Z\-]+)\b", candidate, re.I)
        if m_code:
            codigo = clean(m_code.group(1))

        unidad = "m" if re.search(r"\b(m|mt|mts|metro|metros)\b", candidate, re.I) else "u"

        out.append(make_item(
            filename, prov, cot, len(out)+1, codigo, "", candidate,
            grupo, unidad, cant, punit, subtotal, "Detectado por lector universal: revisar/validar",
            "A confirmar", "A confirmar", "universal"
        ))

    # intenta unir líneas partidas: una descripción técnica puede estar cortada en 2-3 líneas
    for line in lines:
        l = norm(line)
        if any(b in l for b in BAD_WORDS):
            if buffer:
                flush_candidate(" ".join(buffer))
                buffer = []
            continue

        if looks_like_item(line):
            if buffer:
                flush_candidate(" ".join(buffer))
                buffer = []
            flush_candidate(line)
        else:
            # si parece continuación técnica, la acumula
            if any(w in l for w in TECH_WORDS) or detect_group(line):
                buffer.append(line)
            elif buffer:
                buffer.append(line)
                if len(buffer) >= 3:
                    flush_candidate(" ".join(buffer))
                    buffer = []

    if buffer:
        flush_candidate(" ".join(buffer))

    # elimina duplicados muy parecidos
    unique = []
    seen = set()
    for it in out:
        key = (it["descripcion"][:80], it["precio_unitario"], it["subtotal_sin_iva"])
        if key not in seen:
            seen.add(key)
            unique.append(it)
    return unique

# =========================
# COMPARACIÓN Y EXPORT
# =========================

def compare_items(items):
    df = pd.DataFrame(items)
    if df.empty:
        return [], "No hay ítems para comparar."
    df["grupo_comparable"] = df["grupo_comparable"].fillna("").astype(str).str.strip()
    df = df[df["grupo_comparable"] != ""]
    comps, summary = [], []
    for group, g in df.groupby("grupo_comparable"):
        if g["proveedor"].astype(str).str.lower().nunique() < 2:
            continue
        g = g.sort_values("precio_unitario")
        best = g.iloc[0]
        offers = []
        for _, r in g.iterrows():
            d = r.to_dict()
            d["dif_unit_vs_mejor"] = float(r["precio_unitario"] - best["precio_unitario"])
            d["dif_total_vs_mejor"] = float(r["total_con_iva"] - best["total_con_iva"])
            d["recomendado_precio"] = float(r["precio_unitario"]) == float(best["precio_unitario"])
            offers.append(d)
        comps.append({
            "grupo_comparable": group,
            "mejor_proveedor": best["proveedor"],
            "moneda": best["moneda"],
            "mejor_precio_unitario": float(best["precio_unitario"]),
            "mejor_total_con_iva": float(best["total_con_iva"]),
            "ofertas": offers
        })
        summary.append(f"{group}: mejor precio unitario {best['proveedor']} — USD {best['precio_unitario']:.2f}. Validar mínimos, entrega y equivalencia técnica.")
    return comps, "\n".join(summary) if summary else "No hay grupos comparables con 2 o más proveedores. Revisá/normalizá la columna 'Grupo comparable'."

def build_excel(items, comps, summary):
    wb = Workbook()
    thin = Side(style="thin", color="D9D9D9")
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)

    ws = wb.active
    ws.title = "Items Detectados"
    headers = ["Proveedor","Cotización","N°","Código","Marca","Descripción","Formación","Grupo comparable","Moneda","Cant. pedida","Cant. real","Unidad","P. unit.","Subtotal s/IVA","IVA %","IVA monto","Total c/IVA","Mínimo","Entrega","Notas","Parser"]
    for c,h in enumerate(headers,1):
        cell = ws.cell(1,c,h)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = Border(left=thin,right=thin,top=thin,bottom=thin)

    for r,it in enumerate(items,2):
        vals = [it.get("proveedor",""),it.get("cotizacion",""),it.get("nro_item",""),it.get("codigo",""),it.get("marca",""),it.get("descripcion",""),it.get("formacion",""),it.get("grupo_comparable",""),it.get("moneda",""),it.get("cantidad_pedida",0),it.get("cantidad_real",0),it.get("unidad",""),it.get("precio_unitario",0),it.get("subtotal_sin_iva",0),it.get("iva_pct",0),it.get("iva_monto",0),it.get("total_con_iva",0),it.get("minimo_compra",""),it.get("entrega",""),it.get("notas",""),it.get("parser","")]
        for c,v in enumerate(vals,1):
            cell = ws.cell(r,c,v)
            cell.border = Border(left=thin,right=thin,top=thin,bottom=thin)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if c in [13,14,16,17]:
                cell.number_format = '$ #,##0.00'

    for i,w in enumerate([18,18,8,18,14,55,16,18,10,14,14,10,14,16,10,14,16,20,18,36,16],1):
        ws.column_dimensions[get_column_letter(i)].width = w

    if items:
        tab = Table(displayName="ItemsDetectados", ref=f"A1:U{len(items)+1}")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tab)

    ws2 = wb.create_sheet("Comparaciones")
    h2 = ["Grupo","Proveedor","Descripción","Cant. real","P. unit.","Subtotal","IVA","Total","Dif. unit.","Dif. total","Estado"]
    for c,h in enumerate(h2,1):
        cell = ws2.cell(1,c,h)
        cell.fill = fill
        cell.font = font
        cell.border = Border(left=thin,right=thin,top=thin,bottom=thin)

    row = 2
    for comp in comps:
        for off in comp.get("ofertas", []):
            vals = [comp.get("grupo_comparable",""),off.get("proveedor",""),off.get("descripcion",""),off.get("cantidad_real",0),off.get("precio_unitario",0),off.get("subtotal_sin_iva",0),off.get("iva_monto",0),off.get("total_con_iva",0),off.get("dif_unit_vs_mejor",0),off.get("dif_total_vs_mejor",0),"MEJOR PRECIO" if off.get("recomendado_precio") else "Alternativa"]
            for c,v in enumerate(vals,1):
                cell = ws2.cell(row,c,v)
                cell.border = Border(left=thin,right=thin,top=thin,bottom=thin)
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                if c in [5,6,7,8,9,10]:
                    cell.number_format = '$ #,##0.00'
                if c == 11 and v == "MEJOR PRECIO":
                    cell.fill = PatternFill("solid", fgColor="C6EFCE")
            row += 1

    for i,w in enumerate([16,18,55,14,14,16,14,16,14,14,18],1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    ws3 = wb.create_sheet("Resumen")
    ws3["A1"] = "Resumen Ejecutivo"
    ws3["A1"].font = Font(size=16, bold=True)
    ws3["A3"] = summary or "Sin resumen."
    ws3["A3"].alignment = Alignment(wrap_text=True, vertical="top")
    ws3.column_dimensions["A"].width = 120

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# =========================
# ENDPOINTS
# =========================

@app.post("/api/analyze")
async def analyze(files: List[UploadFile] = File(...)):
    all_items, raw = [], []
    for f in files:
        data = await f.read()
        filename = f.filename or "archivo"
        if filename.lower().endswith(".pdf"):
            try:
                pages = pdf_pages(data)
                text = "\n".join(pages)
                prov = detect_provider(filename, text)

                if prov == "Ingeniería Boggio":
                    items = parse_boggio(filename, pages)
                elif prov == "Marlew":
                    items = parse_marlew(filename, pages)
                elif prov == "Ateco":
                    items = parse_ateco(filename, pages)
                else:
                    items = parse_universal(filename, pages)

                # si el específico falló, usa universal
                if not items:
                    items = parse_universal(filename, pages)

                all_items.extend(items)
                raw.append({"archivo": filename, "tabla": f"Texto PDF - {prov}", "columns": ["Texto"], "rows": [[line] for line in text.splitlines()[:500]]})
            except Exception as e:
                raw.append({"archivo": filename, "tabla": "ERROR", "columns": ["Error"], "rows": [[str(e)]]})

    return {"items": all_items, "raw_tables": raw, "warnings": ["V3.5: lector universal + validación. Revisar ítems antes de comparar."]}

@app.post("/api/compare")
async def compare(payload: dict):
    comps, summary = compare_items(payload.get("items", []))
    return {"comparisons": comps, "summary": summary}

@app.post("/api/export_excel")
async def export_excel(payload: dict):
    out = build_excel(payload.get("items", []), payload.get("comparisons", []), payload.get("summary", ""))
    return StreamingResponse(out, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=comparativa_cotizaciones.xlsx"})

@app.get("/api/health")
def health():
    return {"status": "ok", "version": "v3.5-lector-universal"}
