
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

app = FastAPI(title="Comparador Cotizaciones V5.2 Universal Cables")

app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

def clean(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()

def norm(s):
    s = clean(s).lower()
    for a,b in [("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u"),("ñ","n"),("²","2")]:
        s = s.replace(a,b)
    return s

def parse_num(v):
    if isinstance(v,(int,float)):
        return float(v)
    s = re.sub(r"[^0-9,.\-]","",str(v or "").strip())
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(".","").replace(",",".") if s.rfind(",") > s.rfind(".") else s.replace(",","")
    elif "," in s:
        s = s.replace(".","").replace(",",".")
    elif s.count(".") > 1:
        s = s.replace(".","")
    try:
        return float(s)
    except:
        return None

def pdf_pages(data):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        return [p.extract_text(x_tolerance=1,y_tolerance=3) or "" for p in pdf.pages]

def detect_currency(text):
    t = text.upper()
    if "USD" in t or "U$S" in t or "U$" in t or "DOLAR" in t or "DÓLAR" in t:
        return "USD"
    if "$" in t or "PESOS" in t or "ARS" in t:
        return "ARS"
    return "USD"

def detect_provider(filename,text):
    f = filename.lower()
    l = text.lower()
    if "p037357" in f or "marlew" in l:
        return "Marlew"
    if "nv-0004" in f or "boggio" in l or "ingenieria boggio" in l or "ingeniería boggio" in l:
        return "Ingeniería Boggio"
    if "ateco cables" in l:
        return "Ateco"
    if "provemet" in l or "complemet" in l or "grupoprovemet" in l:
        return "Provemet / Complemet"
    if "ivanar" in l:
        return "IVANAR"
    if "la tornillera" in l:
        return "La Tornillera"
    if "hg confecciones" in l or "ropa de trabajo" in l:
        return "HG Confecciones"
    if f.startswith("cot_") or "cable mallado" in l:
        return "Proveedor Cables"
    return Path(filename).stem

def detect_quote(filename,text):
    for p in [r"COTIZACI[ÓO]N\s*N[°º]?\s*([0-9]+)", r"N[°º]\s*:\s*([0-9]+)", r"PRESUPUESTO\s+([0-9]+)", r"PR\s*([0-9]+)"]:
        m = re.search(p,text,re.I|re.S)
        if m:
            return clean(m.group(1))
    return Path(filename).stem

def cable_formation(text):
    t = clean(text)
    m = re.search(r"(\d+)\s*[xX]\s*(\d+(?:[\.,]\d+)?)(?:\s*\+\s*B?(\d+))?", t)
    if m:
        a = m.group(1)
        b = m.group(2).replace(",",".")
        c = m.group(3)
        return f"{a}x{b}+B{c}" if c else f"{a}x{b}"
    maps = {"EC 0210":"2x1","EC 0215":"2x1.5","EC 0307":"3x0.75","EC 0410":"4x1","NF 11500":"1x150","OF 1210":"12x1","NF 0215":"2x1.5","NF 0225":"2x2.5","NF 0315":"3x1.5","NF 0325":"3x2.5","NF 0425":"4x2.5","NF 0440":"4x4","OF 0715":"7x1.5","VK 1160":"1x16","VK 0125":"1x2.5","BX 0325":"3x2.5+B6"}
    u = t.upper()
    for k,v in maps.items():
        if k in u:
            return v
    return ""

def normalize_group(text):
    cf = cable_formation(text)
    if cf:
        return "CABLE_" + cf
    base = norm(text).upper()
    base = re.sub(r"[^A-Z0-9X/\"\. ]"," ",base)
    return re.sub(r"\s+"," ",base).strip()[:80]

def make_item(filename,prov,cot,nro,codigo,marca,desc,unidad,cant,punit,subtotal,moneda="USD",iva_pct=21,notas="",minimo="",entrega="",parser=""):
    desc = clean(desc)
    grupo = normalize_group(" ".join([codigo or "", desc]))
    subtotal = subtotal if subtotal is not None else (cant or 0) * (punit or 0)
    iva = subtotal * iva_pct / 100
    return {"archivo":filename,"proveedor":prov,"cotizacion":cot,"nro_item":str(nro),"codigo":codigo or "","codigo_interno":"","marca":marca or "","descripcion":desc,"formacion":grupo,"grupo_comparable":grupo,"moneda":moneda,"unidad":unidad or "u","cantidad_pedida":cant or 0,"cantidad_real":cant or 0,"precio_unitario":punit or 0,"subtotal_sin_iva":subtotal or 0,"iva_pct":iva_pct,"iva_monto":iva,"total_con_iva":subtotal+iva,"minimo_compra":minimo,"venta_fraccionada":"","entrega":entrega,"notas":notas,"parser":parser,"validado":False}

def parse_generic_cables(filename,pages):
    text = "\n".join(pages)
    prov = detect_provider(filename,text)
    cot = detect_quote(filename,text)
    moneda = detect_currency(text)
    out = []
    for line in [clean(x) for x in text.splitlines() if clean(x)]:
        if "cable" not in norm(line):
            continue
        if any(x in norm(line) for x in ["telefono","domicilio","cuit","mail","email","cliente","subtotal","total presupuesto"]):
            continue
        m = re.match(r"^([A-Z0-9\-]{3,12})\s+(.+?)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)$", line, re.I)
        if m:
            codigo, desc, qty_raw, punit_raw, subtotal_raw = m.groups()
            cant = parse_num(qty_raw) or 1
            punit = parse_num(punit_raw) or 0
            subtotal = parse_num(subtotal_raw) or cant*punit
            if cant > 0 and punit > 0 and subtotal > 0:
                out.append(make_item(filename,prov,cot,len(out)+1,codigo,"",desc,"m",cant,punit,subtotal,moneda,21,"Detectado por lector cables genérico","A confirmar","A confirmar","generic_cables_v5_2"))
                continue
        nums = [parse_num(x) for x in re.findall(r"\d+(?:[\.,]\d+)?", line)]
        nums = [x for x in nums if x is not None]
        if len(nums) >= 3:
            cant, punit, subtotal = nums[-3], nums[-2], nums[-1]
            if cant <= 0 or punit <= 0 or subtotal <= 0:
                continue
            desc = re.sub(r"\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s*$","",line).strip()
            codigo = ""
            mcode = re.match(r"^([A-Z0-9\-]{3,12})\s+(.+)$", desc, re.I)
            if mcode:
                codigo, desc = mcode.group(1), mcode.group(2)
            out.append(make_item(filename,prov,cot,len(out)+1,codigo,"",desc,"m",cant,punit,subtotal,moneda,21,"Detectado por lector cables genérico","A confirmar","A confirmar","generic_cables_v5_2"))
    return out

def parse_marlew(filename,pages):
    items = parse_generic_cables(filename,pages)
    for it in items:
        it["proveedor"] = "Marlew"
        it["marca"] = "MARLEW"
        it["parser"] = "marlew_generic_v5_2"
    return items

def parse_ateco(filename,pages):
    text = "\n".join(pages)
    cot = detect_quote(filename,text)
    entrega = "5 días" if "Plazo de entrega: 5 dias" in text else "Consultar"
    out = []
    for ln in [clean(x) for x in text.splitlines() if clean(x)]:
        m = re.match(r"^([0-9]{4}-[0-9]-[0-9]{6}-[0-9])\s+(\d+)\s+(.+?)\s+sa\s+([\d\.,]+)\s+([\d\.,]+)$",ln,re.I)
        if not m: 
            continue
        cod,cant_raw,desc,punit_raw,sub_raw = m.groups()
        cant = parse_num(cant_raw) or 1
        punit = parse_num(punit_raw) or 0
        subtotal = parse_num(sub_raw) or cant*punit
        out.append(make_item(filename,"Ateco",cot,len(out)+1,cod,"ATECO",desc,"m",cant,punit,subtotal,"USD",21,"Contado anticipo","No informado",entrega,"ateco_v5_2"))
    return out or parse_generic_cables(filename,pages)

def parse_universal(filename,pages):
    items = parse_generic_cables(filename,pages)
    if items:
        return items
    text = "\n".join(pages)
    prov = detect_provider(filename,text)
    cot = detect_quote(filename,text)
    moneda = detect_currency(text)
    out = []
    for line in [clean(x) for x in text.splitlines() if clean(x)]:
        l = norm(line)
        if any(b in l for b in ["cuit","telefono","domicilio","cliente","fecha","subtotal","total","factura"]):
            continue
        nums = [parse_num(x) for x in re.findall(r"\d+(?:[\.,]\d+)?",line)]
        nums = [x for x in nums if x is not None]
        if len(nums) < 3:
            continue
        if not any(w in l for w in ["cable","tubo","caño","cano","perfil","varilla","tuerca","arandela","camisa","pantalon","buzo"]):
            continue
        cant,punit,subtotal = nums[-3],nums[-2],nums[-1]
        desc = re.sub(r"\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s*$","",line).strip()
        out.append(make_item(filename,prov,cot,len(out)+1,"","",desc,"u",cant,punit,subtotal,moneda,21,"Detectado por lector universal: revisar","A confirmar","A confirmar","universal_v5_2"))
    return out

def compare_items(items):
    for it in items:
        g = normalize_group(" ".join([it.get("codigo",""), it.get("grupo_comparable",""), it.get("descripcion","")]))
        if g:
            it["grupo_comparable"] = g
    df = pd.DataFrame(items)
    if df.empty:
        return [], "No hay ítems para comparar."
    df["grupo_comparable"] = df["grupo_comparable"].fillna("").astype(str).str.strip()
    df = df[df["grupo_comparable"] != ""]
    comps, summary = [], []
    for group,g in df.groupby("grupo_comparable"):
        if g["proveedor"].astype(str).str.lower().nunique() < 2:
            continue
        note = "Monedas distintas: validar tipo de cambio." if g["moneda"].astype(str).nunique() > 1 else ""
        g = g.sort_values("precio_unitario")
        best = g.iloc[0]
        offers = []
        for _,r in g.iterrows():
            d = r.to_dict()
            d["dif_unit_vs_mejor"] = float(r["precio_unitario"] - best["precio_unitario"])
            d["dif_total_vs_mejor"] = float(r["total_con_iva"] - best["total_con_iva"])
            d["recomendado_precio"] = float(r["precio_unitario"]) == float(best["precio_unitario"])
            offers.append(d)
        comps.append({"grupo_comparable":group,"mejor_proveedor":best["proveedor"],"moneda":best["moneda"],"mejor_precio_unitario":float(best["precio_unitario"]),"mejor_total_con_iva":float(best["total_con_iva"]),"nota":note,"ofertas":offers})
        summary.append(f"{group}: mejor precio unitario {best['proveedor']} — {best['moneda']} {best['precio_unitario']:.2f}. {note} Validar equivalencia técnica, mínimos y entrega.")
    if not summary:
        grupos = sorted(df["grupo_comparable"].unique().tolist())[:20]
        return comps, "No hay grupos comparables con 2 o más proveedores. Grupos detectados: " + ", ".join(grupos)
    return comps, "\n".join(summary)

def build_excel(items,comps,summary):
    wb = Workbook()
    thin = Side(style="thin",color="D9D9D9")
    fill = PatternFill("solid",fgColor="1F4E78")
    font = Font(color="FFFFFF",bold=True)
    green = PatternFill("solid",fgColor="C6EFCE")
    ws0 = wb.active
    ws0.title = "Resumen Ejecutivo"
    ws0["A1"] = "Resumen Ejecutivo - Comparativa de Cotizaciones"
    ws0["A1"].font = Font(size=16,bold=True)
    ws0["A3"] = summary or "Sin resumen."
    ws0["A3"].alignment = Alignment(wrap_text=True,vertical="top")
    ws0.column_dimensions["A"].width = 120
    ws0["A6"]="Cantidad de ítems detectados"; ws0["B6"]=len(items)
    ws0["A7"]="Comparaciones encontradas"; ws0["B7"]=len(comps)
    ws0["A8"]="Proveedores detectados"; ws0["B8"]=len(set([i.get("proveedor","") for i in items]))
    ws = wb.create_sheet("Items Detectados")
    headers = ["Proveedor","Cotización","N°","Código","Marca","Descripción","Grupo comparable","Moneda","Cant. pedida","Cant. real","Unidad","P. unit.","Subtotal s/IVA","IVA %","IVA monto","Total c/IVA","Mínimo","Entrega","Notas","Parser"]
    for c,h in enumerate(headers,1):
        cell = ws.cell(1,c,h); cell.fill=fill; cell.font=font; cell.alignment=Alignment(horizontal="center",wrap_text=True); cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
    for r,it in enumerate(items,2):
        vals = [it.get("proveedor",""),it.get("cotizacion",""),it.get("nro_item",""),it.get("codigo",""),it.get("marca",""),it.get("descripcion",""),it.get("grupo_comparable",""),it.get("moneda",""),it.get("cantidad_pedida",0),it.get("cantidad_real",0),it.get("unidad",""),it.get("precio_unitario",0),it.get("subtotal_sin_iva",0),it.get("iva_pct",0),it.get("iva_monto",0),it.get("total_con_iva",0),it.get("minimo_compra",""),it.get("entrega",""),it.get("notas",""),it.get("parser","")]
        for c,v in enumerate(vals,1):
            cell = ws.cell(r,c,v); cell.border=Border(left=thin,right=thin,top=thin,bottom=thin); cell.alignment=Alignment(wrap_text=True,vertical="top")
            if c in [12,13,15,16]: cell.number_format = '$ #,##0.00'
    for i,w in enumerate([20,18,8,18,16,60,24,10,14,14,10,14,16,10,14,16,20,20,42,16],1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws2 = wb.create_sheet("Comparaciones")
    headers2 = ["Grupo","Proveedor","Descripción","Moneda","Cant. real","P. unit.","Subtotal","IVA","Total","Dif. unit.","Dif. total","Estado","Nota"]
    for c,h in enumerate(headers2,1):
        cell = ws2.cell(1,c,h); cell.fill=fill; cell.font=font; cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
    row = 2
    for comp in comps:
        for off in comp.get("ofertas",[]):
            vals = [comp.get("grupo_comparable",""),off.get("proveedor",""),off.get("descripcion",""),off.get("moneda",""),off.get("cantidad_real",0),off.get("precio_unitario",0),off.get("subtotal_sin_iva",0),off.get("iva_monto",0),off.get("total_con_iva",0),off.get("dif_unit_vs_mejor",0),off.get("dif_total_vs_mejor",0),"MEJOR PRECIO" if off.get("recomendado_precio") else "Alternativa",comp.get("nota","")]
            for c,v in enumerate(vals,1):
                cell = ws2.cell(row,c,v); cell.border=Border(left=thin,right=thin,top=thin,bottom=thin); cell.alignment=Alignment(wrap_text=True,vertical="top")
                if c in [6,7,8,9,10,11]: cell.number_format = '$ #,##0.00'
                if c == 12 and v == "MEJOR PRECIO": cell.fill = green
            row += 1
    for i,w in enumerate([24,20,60,10,14,14,16,14,16,14,14,18,36],1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

@app.post("/api/analyze")
async def analyze(files: List[UploadFile] = File(...)):
    all_items, raw = [], []
    for f in files:
        data = await f.read()
        filename = f.filename or "archivo"
        if not filename.lower().endswith(".pdf"):
            continue
        try:
            pages = pdf_pages(data)
            text = "\n".join(pages)
            prov = detect_provider(filename,text)
            if prov == "Marlew": items = parse_marlew(filename,pages)
            elif prov == "Ateco": items = parse_ateco(filename,pages)
            else: items = parse_universal(filename,pages)
            if not items:
                items = parse_universal(filename,pages)
            all_items.extend(items)
            raw.append({"archivo":filename,"tabla":f"Texto PDF - {prov}","columns":["Texto"],"rows":[[line] for line in text.splitlines()[:600]]})
        except Exception as e:
            raw.append({"archivo":filename,"tabla":"ERROR","columns":["Error"],"rows":[[str(e)]]})
    return {"items":all_items,"raw_tables":raw,"warnings":["V5.2: lector universal de cables mejorado."]}

@app.post("/api/compare")
async def compare(payload: dict):
    comps,summary = compare_items(payload.get("items",[]))
    return {"comparisons":comps,"summary":summary}

@app.post("/api/export_excel")
async def export_excel(payload: dict):
    items = payload.get("items",[])
    comps = payload.get("comparisons",[])
    summary = payload.get("summary","")
    if not comps:
        comps,summary = compare_items(items)
    out = build_excel(items,comps,summary)
    return StreamingResponse(out,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",headers={"Content-Disposition":"attachment; filename=comparativa_cotizaciones_v5_2.xlsx"})

@app.get("/api/health")
def health():
    return {"status":"ok","version":"v5.2-universal-cables"}
