
from fastapi import FastAPI, File, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from typing import List
from pathlib import Path
import io, re, csv
import pandas as pd

app = FastAPI(title="Comparador Cotizaciones V3 - Validación")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def clean(s):
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()

def norm(s):
    s = clean(s).lower()
    for a,b in [("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u"),("²","2")]:
        s=s.replace(a,b)
    return s

def parse_num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return None
    s = re.sub(r"[^0-9,.\-]", "", s)
    if not s:
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        if s.count(".") > 1:
            s = s.replace(".", "")
    try:
        return float(s)
    except Exception:
        return None

def format_group(s):
    s = clean(s)
    s = s.replace("MM2","mm²").replace("mm2","mm²").replace("MM²","mm²")
    s = s.replace(",", ".").replace("X","x")
    s = re.sub(r"\s+", "", s)
    return s

def detect_formation(text):
    t = clean(text)
    # 3x2.5+B6mm² / 2x1mm² / 1x150mm²
    m = re.search(r"(\d+\s*[xX]\s*\d+(?:[\.,]\d+)?(?:\s*\+\s*B?\d+)?\s*(?:mm2|MM2|mm²|MM²)?)", t)
    if m:
        return format_group(m.group(1))
    return ""

def detect_provider(filename, text):
    low = text.lower()
    if "ingeniería boggio" in low or "ingenieria boggio" in low:
        return "Ingeniería Boggio"
    if "marlew" in low:
        return "Marlew"
    if "ateco cables" in low or "ateco" in low:
        return "Ateco"
    return Path(filename).stem

def detect_quote(filename, text):
    m = re.search(r"N[úu]mero:\s*([A-Z0-9\-]+)", text, re.I)
    if m: return m.group(1)
    m = re.search(r"PRESUPUESTO\s+([0-9]+)", text, re.I)
    if m: return m.group(1)
    m = re.search(r"Presupuesto:\s*([0-9]+)", text, re.I)
    if m: return m.group(1)
    return Path(filename).stem

def read_pdf_pages(data):
    import pdfplumber
    pages = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages:
            pages.append(page.extract_text(x_tolerance=1, y_tolerance=3) or "")
    return pages

def mk_item(filename, proveedor, quote, nro, codigo, codigo_int, marca, desc, formacion, unidad, qty, unit_price, subtotal, iva_pct, parser, notas="", min_compra="", entrega="", venta_fraccionada=""):
    if subtotal is None:
        subtotal = (qty or 0) * (unit_price or 0)
    iva_monto = subtotal * iva_pct / 100
    total = subtotal + iva_monto
    group = formacion or codigo or codigo_int
    return {
        "id": f"{proveedor}-{quote}-{nro}-{codigo}-{codigo_int}",
        "archivo": filename,
        "proveedor": proveedor,
        "cotizacion": quote,
        "nro_item": str(nro),
        "codigo": codigo or "",
        "codigo_interno": codigo_int or "",
        "marca": marca or "",
        "descripcion": clean(desc),
        "formacion": formacion or "",
        "grupo_comparable": group or "",
        "moneda": "USD",
        "unidad": unidad or "u",
        "cantidad_pedida": qty or 0,
        "cantidad_real": qty or 0,
        "precio_unitario": unit_price or 0,
        "subtotal_sin_iva": subtotal or 0,
        "iva_pct": iva_pct,
        "iva_monto": iva_monto,
        "total_con_iva": total,
        "minimo_compra": min_compra,
        "venta_fraccionada": venta_fraccionada,
        "entrega": entrega,
        "notas": notas,
        "parser": parser,
        "validado": False
    }

def parse_boggio(filename, pages):
    text = "\n".join(pages)
    quote = detect_quote(filename, text)
    lines = [clean(x) for x in text.splitlines() if clean(x)]
    blocks, cur = [], []
    start = re.compile(r"^\d+\s+\d{3,6}\s+\[[^\]]+\]\s+", re.I)
    ignore = ["cliente:", "cond. iva", "cuit:", "comercial:", "vencimiento:", "terminos de pago",
        "referencia de cliente", "doc. no valido", "presupuesto", "ingenieria boggio",
        "remedios de escalada", "www.", "numero:", "fecha:", "iibb:", "pagina:",
        "subtotal usd", "total usd", "su pedido incluye", "fecha de entrega",
        "los precios", "las garantias", "confirmar stock", "ud. fue atendido"]
    for line in lines:
        low = norm(line)
        if any(x in low for x in ignore): continue
        if start.match(line):
            if cur: blocks.append(" ".join(cur))
            cur = [line]
        elif cur:
            cur.append(line)
    if cur: blocks.append(" ".join(cur))

    items=[]
    brands=["INDECA","WENTINCK","FONSECA","IMSA","PRYSMIAN"]
    brand_re=r"\b("+"|".join(brands)+r")\b"
    entrega = "15/05/2026" if "Fecha de Entrega: 15/05/2026" in text else "Consultar"
    for b in blocks:
        cm = re.match(r"^(\d+)\s+(\d{3,6})\s+\[([^\]]+)\]\s+(.+)", b)
        if not cm: continue
        nro,img_code,code,rest = cm.groups()
        m = re.search(brand_re + r"\s+\(?\*?\)?\s*([\d\.,]+)\s*(m|mt|mts|Unidad|unidad|un|u)?\s+([\d\.,]+)\s+IVA\s*21%\s+([\d\.,]+)", rest, re.I)
        if not m: continue
        brand=m.group(1).upper()
        qty=parse_num(m.group(2)) or 1
        unit_raw=clean(m.group(3)) or "u"
        unit="m" if unit_raw.lower().startswith("m") else "u"
        unit_price=parse_num(m.group(4)) or 0
        subtotal=parse_num(m.group(5)) or qty*unit_price
        desc = re.split(r"\b"+re.escape(brand)+r"\b", rest, flags=re.I)[0]
        desc = clean(desc).replace("()","").strip()
        form=detect_formation(desc)
        notas = "Confirmar stock" if "(*)" in b else ""
        items.append(mk_item(filename,"Ingeniería Boggio",quote,nro,code,img_code,brand,desc,form,unit,qty,unit_price,subtotal,21,"boggio_v3",notas,"No informado",entrega,"A confirmar"))
    return items

def parse_marlew(filename, pages):
    text="\n".join(pages)
    quote=detect_quote(filename,text)
    lines=[clean(x) for x in text.splitlines() if clean(x)]
    blocks, cur=[],[]
    start=re.compile(r"^(\d{2,5})\s+MT\s+Formaci[oó]n:", re.I)
    ignore=["presupuesto:", "vendedor:", "telefono:", "e-mail:", "empresa:", "atte.:",
        "proyecto:", "pais:", "prov:", "caso n", "fecha:", "condiciones comerciales",
        "moneda y precio", "plazo de entrega", "forma y plazo de pago", "tolerancias",
        "validez", "terminos y condiciones", "notas importantes", "pag1", "pag2", "pag3", "pag4"]
    for line in lines:
        low=norm(line)
        if any(x in low for x in ignore): continue
        if start.match(line):
            if cur: blocks.append(" ".join(cur))
            cur=[line]
        elif cur:
            cur.append(line)
    if cur: blocks.append(" ".join(cur))

    items=[]
    for idx,b in enumerate(blocks,1):
        qm=re.search(r"^(\d{2,5})\s+MT\s+Formaci[oó]n:", b, re.I)
        if not qm: continue
        qty=parse_num(qm.group(1)) or 1
        fm=re.search(r"Formaci[oó]n:\s*([0-9]+x[0-9\.,]+(?:\+B[0-9]+)?\s*(?:mm2|mm²)?)", b, re.I)
        form=format_group(fm.group(1)) if fm else ""
        code=""
        cm=re.search(r"C[oó]digo:\s*([A-Za-z0-9\s\.\-+xX]+?)\s+Mat\s*N", b, re.I)
        if cm: code=clean(cm.group(1))
        mat=""
        mm=re.search(r"Mat\s*N[º°]?\s*:\s*([0-9]+)", b, re.I)
        if mm: mat=mm.group(1)
        family=""
        fam=re.search(r"(Automatizar|Potenciar|Variforce|Instalar)\s*\|", b, re.I)
        if fam: family=fam.group(1)
        nums=re.findall(r"(?<![A-Za-z])(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})(?![A-Za-z])", b)
        if len(nums)<2: continue
        unit_price=parse_num(nums[-3] if len(nums)>=3 else nums[-2]) or 0
        subtotal=parse_num(nums[-1]) or qty*unit_price
        desc = code or f"{family} {form}".strip()
        if family and family.lower() not in desc.lower(): desc=f"{family} {desc}"
        low=norm(b)
        entrega="6/8 semanas" if "6/8 semanas" in low else "A confirmar"
        min_compra=""
        venta=""
        if "minimo de provision" in low: min_compra=f"Mín. {int(qty)} m"
        if "unica bobina" in low or "no fraccionable" in low:
            min_compra=f"Bobina única {int(qty)} m"
            venta="No fraccionable"
        notas=[]
        if "material en stock" in low: notas.append("En stock salvo venta")
        if "unica bobina" in low: notas.append("Única bobina")
        if "no fraccionable" in low: notas.append("No fraccionable")
        items.append(mk_item(filename,"Marlew",quote,idx,code,mat,"MARLEW",desc,form,"m",qty,unit_price,subtotal,21,"marlew_v3","; ".join(notas),min_compra,entrega,venta or "A confirmar"))
    return items

def parse_ateco(filename, pages):
    text="\n".join(pages)
    quote=detect_quote(filename,text)
    items=[]
    lines=[clean(x) for x in text.splitlines() if clean(x)]
    entrega="5 días" if "Plazo de entrega: 5 dias" in text else "Consultar"
    for line in lines:
        # Código Cantidad Descripción Partida Pr.Unitario Importe
        m=re.match(r"^([0-9]{4}-[0-9]-[0-9]{6}-[0-9])\s+(\d+)\s+(.+?)\s+sa\s+([\d\.,]+)\s+([\d\.,]+)$", line, re.I)
        if not m: continue
        code, qty_raw, desc, unit_raw, subtotal_raw = m.groups()
        qty=parse_num(qty_raw) or 1
        unit_price=parse_num(unit_raw) or 0
        subtotal=parse_num(subtotal_raw) or qty*unit_price
        form=detect_formation(desc)
        items.append(mk_item(filename,"Ateco",quote,len(items)+1,code,"","ATECO",desc,form,"m",qty,unit_price,subtotal,21,"ateco_v3","Contado anticipo","No informado",entrega,"A confirmar"))
    return items

def compare_confirmed(items):
    if not items: return [], "No hay ítems confirmados."
    df=pd.DataFrame(items)
    if df.empty: return [], "No hay ítems confirmados."
    df["grupo_comparable"]=df["grupo_comparable"].fillna("").astype(str).str.strip()
    df=df[df["grupo_comparable"]!=""]
    comps=[]
    summary=[]
    for group,g in df.groupby("grupo_comparable"):
        if g["proveedor"].astype(str).str.lower().nunique()<2: continue
        g=g.sort_values("precio_unitario")
        best=g.iloc[0]
        offers=[]
        for _,r in g.iterrows():
            d=r.to_dict()
            d["dif_unit_vs_mejor"]=float(r["precio_unitario"]-best["precio_unitario"])
            d["dif_total_vs_mejor"]=float(r["total_con_iva"]-best["total_con_iva"])
            d["recomendado_precio"]=bool(float(r["precio_unitario"])==float(best["precio_unitario"]))
            offers.append(d)
        comps.append({"grupo_comparable":group,"mejor_proveedor":best["proveedor"],"moneda":best["moneda"],"mejor_precio_unitario":float(best["precio_unitario"]),"mejor_total_con_iva":float(best["total_con_iva"]),"ofertas":offers})
        summary.append(f"{group}: mejor precio unitario {best['proveedor']} - {best['moneda']} {best['precio_unitario']:.2f}. Validar mínimos, plazos y equivalencia técnica antes de OC.")
    return comps, "\n".join(summary) if summary else "No hay grupos comparables con 2 o más proveedores."

@app.post("/api/analyze")
async def analyze(files: List[UploadFile] = File(...)):
    all_items=[]
    raw_tables=[]
    for f in files:
        data=await f.read()
        filename=f.filename or "archivo"
        if filename.lower().endswith(".pdf"):
            try:
                pages=read_pdf_pages(data)
                text="\n".join(pages)
                provider=detect_provider(filename,text)
                if provider=="Ingeniería Boggio": items=parse_boggio(filename,pages)
                elif provider=="Marlew": items=parse_marlew(filename,pages)
                elif provider=="Ateco": items=parse_ateco(filename,pages)
                else: items=[]
                all_items.extend(items)
                raw_tables.append({"archivo":filename,"tabla":f"Texto PDF - {provider}","columns":["Texto"],"rows":[[line] for line in text.splitlines()[:500]]})
            except Exception as e:
                raw_tables.append({"archivo":filename,"tabla":"ERROR","columns":["Error"],"rows":[[str(e)]]})
    return {"items":all_items,"raw_tables":raw_tables,"warnings":["V3: revise y confirme los ítems antes de comparar.","Los grupos comparables pueden editarse en pantalla.","Se recomienda validar equivalencia técnica antes de emitir OC."]}

@app.post("/api/compare")
async def compare(payload: dict):
    items=payload.get("items",[])
    comps,summary=compare_confirmed(items)
    return {"comparisons":comps,"summary":summary}

@app.get("/api/health")
def health():
    return {"status":"ok","version":"v3-validacion"}
