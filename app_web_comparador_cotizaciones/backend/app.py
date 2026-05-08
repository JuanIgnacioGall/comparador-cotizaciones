
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

app = FastAPI(title="Comparador Cotizaciones V4 Hibrido")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

def clean(s): return re.sub(r"\s+", " ", str(s or "")).strip()

def norm(s):
    s = clean(s).lower()
    for a,b in [("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u"),("²","2")]:
        s=s.replace(a,b)
    return s

def parse_num(v):
    if isinstance(v,(int,float)): return float(v)
    s=re.sub(r"[^0-9,.\-]","",str(v or "").strip())
    if not s: return None
    if "," in s and "." in s:
        s=s.replace(".","").replace(",",".") if s.rfind(",")>s.rfind(".") else s.replace(",","")
    elif "," in s:
        s=s.replace(".","").replace(",",".")
    elif s.count(".")>1:
        s=s.replace(".","")
    try: return float(s)
    except Exception: return None

def money_currency(text):
    t=text.upper()
    if "USD" in t or "U$" in t or "U$S" in t or "DOLAR" in t: return "USD"
    if "$" in t or "PESOS" in t or "ARS" in t: return "ARS"
    return "USD"

def detect_group(text):
    t=clean(text)
    m=re.search(r"(\d+\s*[xX]\s*\d+(?:[\.,]\d+)?(?:\s*\+\s*B?\d+)?\s*(?:mm2|MM2|mm²|MM²)?)",t)
    if m: return re.sub(r"\s+","",m.group(1).replace("X","x").replace(",",".").replace("MM2","mm²").replace("mm2","mm²"))
    m=re.search(r"((?:\d+\s*)?(?:1/2|1/4|3/4)|\d+|1\s*1/2|1\s*1/4)\s*'?''?\s*SCH\s*(\d+)",t,re.I)
    if m: return clean(m.group(0)).replace("''",'"').replace(" ","")
    m=re.search(r"SCH\s*(\d+).*?Ø\s*([\d\.,]+)",t,re.I)
    if m: return f"SCH{m.group(1)}-Ø{m.group(2).replace(',', '.')}"
    u=t.upper()
    for k,v in {"EC 0210":"2x1mm²","EC 0215":"2x1.5mm²","EC 0307":"3x0.75mm²","EC 0410":"4x1mm²","NF 11500":"1x150mm²","OF 1210":"12x1mm²","NF 0215":"2x1.5mm²","NF 0225":"2x2.5mm²","NF 0315":"3x1.5mm²","NF 0325":"3x2.5mm²","NF 0425":"4x2.5mm²","NF 0440":"4x4mm²","OF 0715":"7x1.5mm²","VK 1160":"1x16mm²","VK 0125":"1x2.5mm²"}.items():
        if k in u: return v
    for key in ["VARILLA ROSCADA","TUERCA HEXAG","ARANDELA LISA","PERFIL UPN"]:
        if key in u: return clean(u.split(" 0,00")[0])[:70]
    if any(w in norm(t) for w in ["camisa","pantalon","pantalón","remera","buzo","logo bordado"]):
        return re.sub(r"\s+"," ",t)[:70]
    return ""

def pdf_pages(data):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        return [p.extract_text(x_tolerance=1,y_tolerance=3) or "" for p in pdf.pages]

def detect_provider(filename,text):
    l=text.lower()
    if "provemet" in l or "complemet" in l: return "Provemet / Complemet"
    if "ivanar" in l: return "IVANAR"
    if "la tornillera" in l: return "La Tornillera"
    if "hg confecciones" in l or "hg ropa de trabajo" in l: return "HG Confecciones"
    if "ingeniería boggio" in l or "ingenieria boggio" in l: return "Ingeniería Boggio"
    if "marlew" in l: return "Marlew"
    if "ateco cables" in l: return "Ateco"
    lines=[clean(x) for x in text.splitlines() if clean(x)]
    for line in lines[:15]:
        nl=norm(line)
        if any(x in nl for x in ["presupuesto","cotizacion","cotización","fecha","cliente","cuit","domicilio"]): continue
        if 4 <= len(line) <= 60: return line[:60]
    return Path(filename).stem

def detect_quote(filename,text):
    for p in [r"COTIZACI[ÓO]N\s*N[°º]?\s*([0-9]+)",r"N[°º]\s*:\s*([0-9]+)",r"PRESUPUESTO\s*Nro\.\s*([0-9\-]+)",r"PRESUPUESTO\s+NUMERO\s*:\s*([0-9]+)",r"Presupuesto:\s*([0-9]+)",r"N[úu]mero:\s*([A-Z0-9\-]+)",r"PR\s*([0-9]+)"]:
        m=re.search(p,text,re.I|re.S)
        if m: return clean(m.group(1))
    return Path(filename).stem

def make_item(filename,prov,cot,nro,codigo,marca,desc,grupo,unidad,cant,punit,subtotal,moneda="USD",iva_pct=21,notas="",minimo="",entrega="",parser=""):
    subtotal=subtotal if subtotal is not None else (cant or 0)*(punit or 0)
    iva=subtotal*iva_pct/100
    return {"archivo":filename,"proveedor":prov,"cotizacion":cot,"nro_item":str(nro),"codigo":codigo or "","codigo_interno":"","marca":marca or "","descripcion":clean(desc),"formacion":grupo or "","grupo_comparable":grupo or codigo or clean(desc)[:60],"moneda":moneda,"unidad":unidad or "u","cantidad_pedida":cant or 0,"cantidad_real":cant or 0,"precio_unitario":punit or 0,"subtotal_sin_iva":subtotal or 0,"iva_pct":iva_pct,"iva_monto":iva,"total_con_iva":subtotal+iva,"minimo_compra":minimo,"venta_fraccionada":"","entrega":entrega,"notas":notas,"parser":parser,"validado":False}

def parse_provemet(filename,pages):
    text="\n".join(pages); cot=detect_quote(filename,text); out=[]; lines=[clean(x) for x in text.splitlines() if clean(x)]
    for i,line in enumerate(lines):
        if not line.upper().startswith("TUBO S/C"): continue
        desc=line
        if i+1<len(lines) and not re.match(r"^\d+\s+[\d\.,]+",lines[i+1]): desc+=" "+lines[i+1]
        data=""
        for j in range(i+1,min(i+6,len(lines))):
            if re.match(r"^\d+\s+[\d\.,]+\s+[\d\.,]+\s+Mt\s+[\d\.,]+",lines[j],re.I):
                data=lines[j]; break
        if not data: data=line
        m=re.search(r"^(\d+)\s+([\d\.,]+)\s+([\d\.,]+)\s+Mt\s+([\d\.,]+)\s+([\d\.,]+)\s+Mt\s+([\d\.,]+)",data,re.I)
        if m:
            nro=int(m.group(1)); cant=parse_num(m.group(3)) or 0; punit=parse_num(m.group(5)) or 0; subtotal=parse_num(m.group(6)) or cant*punit
            out.append(make_item(filename,"Provemet / Complemet",cot,nro,"","PROVEMET",desc,detect_group(desc),"Mt",cant,punit,subtotal,"USD",21,"Precios no incluyen IVA","","En stock","provemet_v4"))
    return out

def parse_ivanar(filename,pages):
    text="\n".join(pages); cot=detect_quote(filename,text); out=[]
    for line in [clean(x) for x in text.splitlines() if clean(x)]:
        m=re.match(r"^\s*([0-9]{4,8})\s+(.+?)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+(UN|KG|MT|M|U)\b",line,re.I)
        if not m: continue
        codigo,desc,cant_raw,punit_raw,subtotal_raw,unidad=m.groups()
        if any(x in norm(desc) for x in ["subtotal","total","iva","documento"]): continue
        cant=parse_num(cant_raw) or 1; punit=parse_num(punit_raw) or 0; subtotal=parse_num(subtotal_raw) or cant*punit
        out.append(make_item(filename,"IVANAR",cot,len(out)+1,codigo,"IVANAR",desc,detect_group(desc),unidad,cant,punit,subtotal,"ARS",21,"","No informado","A confirmar","ivanar_v4"))
    return out

def parse_tornillera(filename,pages):
    text="\n".join(pages); cot=detect_quote(filename,text); out=[]
    for line in [clean(x) for x in text.splitlines() if clean(x)]:
        m=re.match(r"^([\d\.,]+)\s+\*\s+(.+?)\s+[\d\.,]+%\s+([\d\.,]+)\s+([\d\.,]+)$",line,re.I)
        if not m: continue
        cant_raw,desc,punit_raw,subtotal_raw=m.groups()
        cant=parse_num(cant_raw) or 1; punit=parse_num(punit_raw) or 0; subtotal=parse_num(subtotal_raw) or cant*punit
        out.append(make_item(filename,"La Tornillera",cot,len(out)+1,"","LA TORNILLERA",desc,detect_group(desc),"u",cant,punit,subtotal,"ARS",21,"Sujeto a disponibilidad","No informado","24/04/2026","tornillera_v4"))
    return out

def parse_hg(filename,pages):
    text="\n".join(pages); cot=detect_quote(filename,text); out=[]
    for line in [clean(x) for x in text.splitlines() if clean(x)]:
        m=re.match(r"^(.+?)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)$",line)
        if not m: continue
        desc,cant_raw,punit_raw,subtotal_raw=m.groups(); nd=norm(desc)
        if len(desc)<8 or any(x in nd for x in ["subtotal","total","iva","forma de pago","terminos","válido","valido"]): continue
        if not any(x in nd for x in ["camisa","pantalon","pantalón","remera","buzo","logo"]): continue
        cant=parse_num(cant_raw) or 1; punit=parse_num(punit_raw) or 0; subtotal=parse_num(subtotal_raw) or cant*punit
        out.append(make_item(filename,"HG Confecciones",cot,len(out)+1,"","HG",desc,detect_group(desc),"u",cant,punit,subtotal,"ARS",21,"Plazo sujeto a talles/seña","No informado","A confirmar","hg_v4"))
    return out

def parse_boggio(filename,pages):
    text="\n".join(pages); cot=detect_quote(filename,text); entrega="15/05/2026" if "Fecha de Entrega: 15/05/2026" in text else "Consultar"
    lines=[clean(x) for x in text.splitlines() if clean(x)]; blocks=[]; cur=[]; start=re.compile(r"^\d+\s+\d{3,6}\s+\[[^\]]+\]\s+",re.I)
    ignore=("X ","Doc. no","Ingeniería Boggio","Remedios","Argentina -","www.","Cliente:","L De La Torre","Cond. IVA","CUIT:","Comercial:","Vencimiento:","Términos","Referencia","% de","N° Imagen","Subtotal USD","IVA 21%","Total USD","JUANI","PRECIOS NETOS","NETO PAGO","Su pedido","Para más","https:","Fecha de Entrega","Los precios","Las garantias","(*)Confirmar","Ud. fue","Página:")
    for ln in lines:
        if ln.startswith(ignore): continue
        if start.match(ln):
            if cur: blocks.append(" ".join(cur))
            cur=[ln]
        elif cur: cur.append(ln)
    if cur: blocks.append(" ".join(cur))
    out=[]; brand_re=r"\b(INDECA|WENTINCK|FONSECA|IMSA|PRYSMIAN)\b"
    for b in blocks:
        cm=re.match(r"^(\d+)\s+(\d{3,6})\s+\[([^\]]+)\]\s+(.+)$",b,re.I)
        if not cm: continue
        nro,img,cod,rest=cm.groups()
        m=re.search(brand_re+r"\s*(?:\(\*\))?\s+([\d\.,]+)\s*(m|mt|mts|Unidad|unidad|un|u)?\s+([\d\.,]+)\s+IVA\s*21%\s+([\d\.,]+)",rest,re.I)
        if not m: continue
        marca=m.group(1).upper(); cant=parse_num(m.group(2)) or 1; unidad="m" if clean(m.group(3)).lower().startswith("m") else "u"; punit=parse_num(m.group(4)) or 0; subtotal=parse_num(m.group(5)) or cant*punit
        desc=clean(rest[:m.start()]).replace("()","").strip()
        out.append(make_item(filename,"Ingeniería Boggio",cot,nro,cod,marca,desc,detect_group(desc),unidad,cant,punit,subtotal,"USD",21,"Confirmar stock" if "(*)" in b else "","No informado",entrega,"boggio_v4"))
    return out

def parse_marlew(filename,pages):
    text="\n".join(pages); cot=detect_quote(filename,text); lines=[clean(x) for x in text.splitlines() if clean(x)]
    starts=[i for i,l in enumerate(lines) if re.match(r"^\d+\s+\d{2,5}\s+MT\s+C[oó]digo:",l,re.I)]
    blocks=[" ".join(lines[s:starts[j+1] if j+1<len(starts) else len(lines)]) for j,s in enumerate(starts)]
    out=[]
    for b in blocks:
        m=re.search(r"^(\d+)\s+(\d{2,5})\s+MT\s+C[oó]digo:\s*(.*?)\s+Formaci[oó]n:\s*([^\s]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)",b,re.I)
        if not m: continue
        nro,cant_raw,cod,form,punit_raw,net,total_raw=m.groups(); cant=parse_num(cant_raw) or 1; punit=parse_num(punit_raw) or 0; subtotal=parse_num(total_raw) or cant*punit
        low=norm(b); entrega="6/8 semanas" if "6/8 semanas" in low else "A confirmar"; minimo=""; notas=[]
        if "minimo de provision" in low: minimo=f"Mín. {int(cant)} m"
        if "unica bobina" in low or "no fraccionable" in low: minimo=f"Bobina única {int(cant)} m"; notas.append("No fraccionable")
        if "material en stock" in low: notas.append("En stock salvo venta")
        out.append(make_item(filename,"Marlew",cot,nro,cod,"MARLEW",cod,detect_group(form),"m",cant,punit,subtotal,"USD",21,"; ".join(notas),minimo,entrega,"marlew_v4"))
    return out

def parse_ateco(filename,pages):
    text="\n".join(pages); cot=detect_quote(filename,text); entrega="5 días" if "Plazo de entrega: 5 dias" in text else "Consultar"; out=[]
    for ln in [clean(x) for x in text.splitlines() if clean(x)]:
        m=re.match(r"^([0-9]{4}-[0-9]-[0-9]{6}-[0-9])\s+(\d+)\s+(.+?)\s+sa\s+([\d\.,]+)\s+([\d\.,]+)$",ln,re.I)
        if not m: continue
        cod,cant_raw,desc,punit_raw,sub_raw=m.groups(); cant=parse_num(cant_raw) or 1; punit=parse_num(punit_raw) or 0; subtotal=parse_num(sub_raw) or cant*punit
        out.append(make_item(filename,"Ateco",cot,len(out)+1,cod,"ATECO",desc,detect_group(desc),"m",cant,punit,subtotal,"USD",21,"Contado anticipo","No informado",entrega,"ateco_v4"))
    return out

BAD=["cuit","iva responsable","domicilio","telefono","tel.","email","mail","www","cliente","fecha","presupuesto","cotizacion","cotización","validez","condicion","condición","forma de pago","subtotal","total","observacion","observación","pagina","página","banco","cbu","vendedor","factura","documento no valido"]
TECH=["cable","caño","cano","tubo","válvula","valvula","brida","codo","tee","reducción","reduccion","cupla","niple","curva","acople","chapa","perfil","aislación","aislacion","mm","pulg","sch","ansi","astm","inox","acero","cobre","pvc","xlpe","blindado","pantalla","motor","bomba","sensor","presostato","termómetro","manómetro","tablero","borne","interruptor","contactor","disyuntor","varilla","tuerca","arandela"]
def looks_like_item(line):
    l=norm(line)
    if len(line)<12 or any(b in l for b in BAD): return False
    if len(re.findall(r"\d+(?:[\.,]\d+)?",line))<2: return False
    return any(w in l for w in TECH) or bool(detect_group(line))
def parse_universal(filename,pages):
    text="\n".join(pages); prov=detect_provider(filename,text); cot=detect_quote(filename,text); moneda=money_currency(text); out=[]
    for line in [clean(x) for x in text.splitlines() if clean(x)]:
        if not looks_like_item(line): continue
        nums=[parse_num(x) for x in re.findall(r"\d+(?:[\.,]\d+)?",line)]; nums=[x for x in nums if x is not None]
        if len(nums)<2: continue
        subtotal=nums[-1]; punit=nums[-2]; cant=round(subtotal/punit,4) if punit and punit>0 else 1
        qm=re.match(r"^\s*([\d\.,]+)\s*(m|mt|mts|un|u|unidad|kg)?\b",line,re.I)
        if qm:
            q=parse_num(qm.group(1))
            if q and q>0: cant=q
        codigo=""; cm=re.search(r"\b([A-Z]{1,5}[- ]?\d{2,8}|[0-9]{3,8}[-/][0-9A-Z\-]+)\b",line,re.I)
        if cm: codigo=clean(cm.group(1))
        unidad="m" if re.search(r"\b(m|mt|mts|metro|metros)\b",line,re.I) else "u"
        out.append(make_item(filename,prov,cot,len(out)+1,codigo,"",line,detect_group(line),unidad,cant,punit,subtotal,moneda,21,"Detectado por lector universal: revisar","A confirmar","A confirmar","universal_v4"))
    seen=set(); unique=[]
    for it in out:
        key=(it["descripcion"][:80],it["subtotal_sin_iva"])
        if key not in seen:
            seen.add(key); unique.append(it)
    return unique

def compare_items(items):
    df=pd.DataFrame(items)
    if df.empty: return [],"No hay ítems para comparar."
    df["grupo_comparable"]=df["grupo_comparable"].fillna("").astype(str).str.strip(); df=df[df["grupo_comparable"]!=""]
    comps=[]; summary=[]
    for group,g in df.groupby("grupo_comparable"):
        if g["proveedor"].astype(str).str.lower().nunique()<2: continue
        g=g.sort_values("precio_unitario"); best=g.iloc[0]; offers=[]
        for _,r in g.iterrows():
            d=r.to_dict(); d["dif_unit_vs_mejor"]=float(r["precio_unitario"]-best["precio_unitario"]); d["dif_total_vs_mejor"]=float(r["total_con_iva"]-best["total_con_iva"]); d["recomendado_precio"]=float(r["precio_unitario"])==float(best["precio_unitario"]); offers.append(d)
        comps.append({"grupo_comparable":group,"mejor_proveedor":best["proveedor"],"moneda":best["moneda"],"mejor_precio_unitario":float(best["precio_unitario"]),"mejor_total_con_iva":float(best["total_con_iva"]),"ofertas":offers})
        summary.append(f"{group}: mejor precio unitario {best['proveedor']} — {best['moneda']} {best['precio_unitario']:.2f}. Validar equivalencia técnica, mínimos y entrega.")
    return comps,"\n".join(summary) if summary else "No hay grupos comparables con 2 o más proveedores. Revisá/normalizá la columna 'Grupo comparable'."

def build_excel(items,comps,summary):
    wb=Workbook(); thin=Side(style="thin",color="D9D9D9"); fill=PatternFill("solid",fgColor="1F4E78"); font=Font(color="FFFFFF",bold=True)
    ws=wb.active; ws.title="Items Detectados"; headers=["Proveedor","Cotización","N°","Código","Marca","Descripción","Grupo comparable","Moneda","Cant. pedida","Cant. real","Unidad","P. unit.","Subtotal s/IVA","IVA %","IVA monto","Total c/IVA","Mínimo","Entrega","Notas","Parser"]
    for c,h in enumerate(headers,1):
        cell=ws.cell(1,c,h); cell.fill=fill; cell.font=font; cell.alignment=Alignment(horizontal="center",wrap_text=True); cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
    for r,it in enumerate(items,2):
        vals=[it.get("proveedor",""),it.get("cotizacion",""),it.get("nro_item",""),it.get("codigo",""),it.get("marca",""),it.get("descripcion",""),it.get("grupo_comparable",""),it.get("moneda",""),it.get("cantidad_pedida",0),it.get("cantidad_real",0),it.get("unidad",""),it.get("precio_unitario",0),it.get("subtotal_sin_iva",0),it.get("iva_pct",0),it.get("iva_monto",0),it.get("total_con_iva",0),it.get("minimo_compra",""),it.get("entrega",""),it.get("notas",""),it.get("parser","")]
        for c,v in enumerate(vals,1):
            cell=ws.cell(r,c,v); cell.border=Border(left=thin,right=thin,top=thin,bottom=thin); cell.alignment=Alignment(wrap_text=True,vertical="top")
            if c in [12,13,15,16]: cell.number_format='$ #,##0.00'
    for i,w in enumerate([20,18,8,18,16,60,22,10,14,14,10,14,16,10,14,16,20,20,42,16],1): ws.column_dimensions[get_column_letter(i)].width=w
    if items:
        tab=Table(displayName="ItemsDetectados",ref=f"A1:T{len(items)+1}"); tab.tableStyleInfo=TableStyleInfo(name="TableStyleMedium2",showRowStripes=True); ws.add_table(tab)
    ws2=wb.create_sheet("Comparaciones"); h2=["Grupo","Proveedor","Descripción","Cant. real","P. unit.","Subtotal","IVA","Total","Dif. unit.","Dif. total","Estado"]
    for c,h in enumerate(h2,1):
        cell=ws2.cell(1,c,h); cell.fill=fill; cell.font=font; cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
    row=2
    for comp in comps:
        for off in comp.get("ofertas",[]):
            vals=[comp.get("grupo_comparable",""),off.get("proveedor",""),off.get("descripcion",""),off.get("cantidad_real",0),off.get("precio_unitario",0),off.get("subtotal_sin_iva",0),off.get("iva_monto",0),off.get("total_con_iva",0),off.get("dif_unit_vs_mejor",0),off.get("dif_total_vs_mejor",0),"MEJOR PRECIO" if off.get("recomendado_precio") else "Alternativa"]
            for c,v in enumerate(vals,1):
                cell=ws2.cell(row,c,v); cell.border=Border(left=thin,right=thin,top=thin,bottom=thin); cell.alignment=Alignment(wrap_text=True,vertical="top")
                if c in [5,6,7,8,9,10]: cell.number_format='$ #,##0.00'
                if c==11 and v=="MEJOR PRECIO": cell.fill=PatternFill("solid",fgColor="C6EFCE")
            row+=1
    for i,w in enumerate([22,20,60,14,14,16,14,16,14,14,18],1): ws2.column_dimensions[get_column_letter(i)].width=w
    ws3=wb.create_sheet("Resumen"); ws3["A1"]="Resumen Ejecutivo"; ws3["A1"].font=Font(size=16,bold=True); ws3["A3"]=summary or "Sin resumen."; ws3["A3"].alignment=Alignment(wrap_text=True,vertical="top"); ws3.column_dimensions["A"].width=120
    out=io.BytesIO(); wb.save(out); out.seek(0); return out

@app.post("/api/analyze")
async def analyze(files: List[UploadFile] = File(...)):
    all_items=[]; raw=[]
    for f in files:
        data=await f.read(); filename=f.filename or "archivo"
        if not filename.lower().endswith(".pdf"): continue
        try:
            pages=pdf_pages(data); text="\n".join(pages); prov=detect_provider(filename,text)
            parser_map={"Provemet / Complemet":parse_provemet,"IVANAR":parse_ivanar,"La Tornillera":parse_tornillera,"HG Confecciones":parse_hg,"Ingeniería Boggio":parse_boggio,"Marlew":parse_marlew,"Ateco":parse_ateco}
            items=parser_map.get(prov,parse_universal)(filename,pages)
            if not items: items=parse_universal(filename,pages)
            all_items.extend(items); raw.append({"archivo":filename,"tabla":f"Texto PDF - {prov}","columns":["Texto"],"rows":[[line] for line in text.splitlines()[:500]]})
        except Exception as e:
            raw.append({"archivo":filename,"tabla":"ERROR","columns":["Error"],"rows":[[str(e)]]})
    return {"items":all_items,"raw_tables":raw,"warnings":["V4: motor híbrido con parsers por patrón + lector universal. Revisar ítems antes de emitir OC."]}

@app.post("/api/compare")
async def compare(payload: dict):
    comps,summary=compare_items(payload.get("items",[]))
    return {"comparisons":comps,"summary":summary}

@app.post("/api/export_excel")
async def export_excel(payload: dict):
    out=build_excel(payload.get("items",[]),payload.get("comparisons",[]),payload.get("summary",""))
    return StreamingResponse(out,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",headers={"Content-Disposition":"attachment; filename=comparativa_cotizaciones.xlsx"})

@app.get("/api/health")
def health():
    return {"status":"ok","version":"v4-hibrido-compras"}
