
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

app = FastAPI(title="Comparador Cotizaciones V6 Industrial")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

def clean(s): return re.sub(r"\s+"," ",str(s or "")).strip()
def norm(s):
    s=clean(s).lower()
    for a,b in [("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u"),("ñ","n"),("²","2")]: s=s.replace(a,b)
    return s
def parse_num(v):
    if isinstance(v,(int,float)): return float(v)
    s=re.sub(r"[^0-9,.\-]","",str(v or "").strip())
    if not s: return None
    if "," in s and "." in s: s=s.replace(".","").replace(",",".") if s.rfind(",")>s.rfind(".") else s.replace(",","")
    elif "," in s: s=s.replace(".","").replace(",",".")
    elif s.count(".")>1: s=s.replace(".","")
    try: return float(s)
    except: return None

def pdf_pages(data):
    import pdfplumber
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        return [p.extract_text(x_tolerance=1,y_tolerance=3) or "" for p in pdf.pages]

def currency(text):
    t=text.upper()
    if "USD" in t or "U$S" in t or "U$" in t or "DOLAR" in t or "DÓLAR" in t: return "USD"
    if "$" in t or "PESOS" in t or "ARS" in t: return "ARS"
    return "USD"

def quote(filename,text):
    pats=[r"COTIZACI[ÓO]N\s*N[°º]?\s*([0-9]+)",r"N[°º]\s*:\s*([0-9]+)",r"PRESUPUESTO\s+([0-9]+)",r"Presupuesto:\s*([0-9]+)",r"N[úu]mero:\s*([A-Z0-9\-]+)",r"PR\s*([0-9]+)"]
    for p in pats:
        m=re.search(p,text,re.I|re.S)
        if m: return clean(m.group(1))
    return Path(filename).stem

def provider(filename,text):
    f=filename.lower(); l=text.lower()
    if "p037357" in f or "marlew" in l: return "Marlew"
    if "nv-0004" in f or "boggio" in l or "ingenieria boggio" in l or "ingeniería boggio" in l: return "Ingeniería Boggio"
    if "ateco" in l: return "Ateco"
    if "provemet" in l or "complemet" in l: return "Provemet / Complemet"
    if "ivanar" in l: return "IVANAR"
    if "tornillera" in l: return "La Tornillera"
    if "hg confecciones" in l or "ropa de trabajo" in l: return "HG Confecciones"
    if f.startswith("cot_") or "cable" in l: return "Proveedor Cables"
    return Path(filename).stem

CODEMAP={"EC0210":"2x1","EC0215":"2x1.5","EC0307":"3x0.75","EC0410":"4x1","NF11500":"1x150","OF1210":"12x1","NF0215":"2x1.5","NF0225":"2x2.5","NF0315":"3x1.5","NF0325":"3x2.5","NF0425":"4x2.5","NF0440":"4x4","OF0715":"7x1.5","VK1160":"1x16","VK0125":"1x2.5","BX0325":"3x2.5+B6"}

def cable_form(text):
    t=clean(text); u=t.upper().replace(" ","")
    for k,v in CODEMAP.items():
        if k in u: return v
    m=re.search(r"(\d+)\s*[xX]\s*(\d+(?:[\.,]\d+)?)(?:\s*(?:MM2|mm2|MM²|mm²))?(?:\s*\+\s*B?(\d+))?",t)
    if m:
        a=m.group(1); b=m.group(2).replace(",","."); b=re.sub(r"\.0+$","",b); c=m.group(3)
        return f"{a}x{b}+B{c}" if c else f"{a}x{b}"
    m=re.search(r"\b(\d+(?:[\.,]\d+)?)\s*MM\s*C/?PANTALLA",t,re.I)
    if m: return "2x"+m.group(1).replace(",",".")
    return ""

def group(text):
    t=clean(text); cf=cable_form(t)
    if cf: return "CABLE_"+cf
    nt=norm(t).upper()
    if any(x in nt for x in ["TUBO","CANO","CAÑO"]):
        size=sch=diam=""
        m=re.search(r"(\d+\s*1/2|\d+\s*1/4|\d+\s*3/4|\d+/\d+|\d+)\s*''?",t)
        if m: size=re.sub(r"\s+","",m.group(1))
        m=re.search(r"SCH\s*(\d+)",t,re.I)
        if m: sch="SCH"+m.group(1)
        m=re.search(r"Ø\s*([\d\.,]+)",t)
        if m: diam="D"+m.group(1).replace(",",".")
        return "_".join([x for x in ["TUBO",size,sch,diam] if x])
    for key in ["VARILLA ROSCADA","TUERCA HEXAG","ARANDELA LISA","PERFIL UPN"]:
        if key in nt: return re.sub(r"\s+"," ",t.upper()).strip()[:80]
    base=re.sub(r"[^A-Z0-9X/\"\. ]"," ",norm(t).upper())
    return re.sub(r"\s+"," ",base).strip()[:80]

def item(fn,prov,cot,nro,cod,marca,desc,unidad,cant,punit,sub,mon="USD",iva=21,notas="",minimo="",entrega="",parser=""):
    desc=clean(desc); g=group((cod or "")+" "+desc); sub=sub if sub is not None else (cant or 0)*(punit or 0); iva_m=sub*iva/100
    return {"archivo":fn,"proveedor":prov,"cotizacion":cot,"nro_item":str(nro),"codigo":cod or "","codigo_interno":"","marca":marca or "","descripcion":desc,"formacion":g,"grupo_comparable":g,"moneda":mon,"unidad":unidad or "u","cantidad_pedida":cant or 0,"cantidad_real":cant or 0,"precio_unitario":punit or 0,"subtotal_sin_iva":sub or 0,"iva_pct":iva,"iva_monto":iva_m,"total_con_iva":sub+iva_m,"minimo_compra":minimo,"venta_fraccionada":"","entrega":entrega,"notas":notas,"parser":parser,"validado":False}

def parse_table_like(fn,pages,prov=None):
    text="\n".join(pages); prov=prov or provider(fn,text); cot=quote(fn,text); mon=currency(text); out=[]
    for line in [clean(x) for x in text.splitlines() if clean(x)]:
        l=norm(line)
        if any(b in l for b in ["cuit","telefono","domicilio","cliente","fecha","subtotal presupuesto","total presupuesto","factura","iva 21 total"]): continue
        if not any(w in l for w in ["cable","tubo","caño","cano","perfil","varilla","tuerca","arandela","camisa","pantalon","buzo"]): continue
        m=re.match(r"^([A-Z0-9\-\/\.]{3,18})\s+(.+?)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)$",line,re.I)
        if m:
            cod,desc,qr,pr,sr=m.groups(); cant=parse_num(qr) or 1; p=parse_num(pr) or 0; sub=parse_num(sr) or cant*p
            if p>0 and sub>0: out.append(item(fn,prov,cot,len(out)+1,cod,"",desc,"m" if "cable" in l or "tubo" in l else "u",cant,p,sub,mon,21,"Detectado por lector universal","A confirmar","A confirmar","universal_v6")); continue
        nums=[parse_num(x) for x in re.findall(r"\d+(?:[\.,]\d+)?",line)]; nums=[x for x in nums if x is not None]
        if len(nums)>=3:
            cant,p,sub=nums[-3],nums[-2],nums[-1]
            if p<=0 or sub<=0: continue
            desc=re.sub(r"\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s*$","",line).strip()
            cod=""; mc=re.match(r"^([A-Z0-9\-\/\.]{3,18})\s+(.+)$",desc,re.I)
            if mc: cod,desc=mc.group(1),mc.group(2)
            out.append(item(fn,prov,cot,len(out)+1,cod,"",desc,"m" if "cable" in l or "tubo" in l else "u",cant,p,sub,mon,21,"Detectado por lector universal","A confirmar","A confirmar","universal_v6"))
    return out

def parse_ateco(fn,pages):
    text="\n".join(pages); cot=quote(fn,text); out=[]
    for ln in [clean(x) for x in text.splitlines() if clean(x)]:
        m=re.match(r"^([0-9]{4}-[0-9]-[0-9]{6}-[0-9])\s+(\d+)\s+(.+?)\s+sa\s+([\d\.,]+)\s+([\d\.,]+)$",ln,re.I)
        if m:
            cod,qr,desc,pr,sr=m.groups(); cant=parse_num(qr) or 1; p=parse_num(pr) or 0; sub=parse_num(sr) or cant*p
            out.append(item(fn,"Ateco",cot,len(out)+1,cod,"ATECO",desc,"m",cant,p,sub,"USD",21,"Contado anticipo","No informado","5 días","ateco_v6"))
    return out or parse_table_like(fn,pages,"Ateco")

def parse_boggio(fn,pages):
    text="\n".join(pages); cot=quote(fn,text); lines=[clean(x) for x in text.splitlines() if clean(x)]; blocks=[]; cur=[]
    start=re.compile(r"^\d+\s+\d{3,6}\s+\[[^\]]+\]",re.I)
    for ln in lines:
        if start.match(ln):
            if cur: blocks.append(" ".join(cur))
            cur=[ln]
        elif cur: cur.append(ln)
    if cur: blocks.append(" ".join(cur))
    out=[]
    for b in blocks:
        cm=re.match(r"^(\d+)\s+(\d{3,6})\s+\[([^\]]+)\]\s+(.+)$",b,re.I)
        if not cm: continue
        nro,img,cod,rest=cm.groups()
        m=re.search(r"\b(INDECA|WENTINCK|FONSECA|IMSA|PRYSMIAN)\b\s*(?:\(\*\))?\s+([\d\.,]+)\s*(m|mt|mts|Unidad|unidad|un|u)?\s+([\d\.,]+)\s+IVA\s*21%\s+([\d\.,]+)",rest,re.I)
        if not m: continue
        marca=m.group(1).upper(); cant=parse_num(m.group(2)) or 1; unidad="m" if clean(m.group(3)).lower().startswith("m") else "u"; p=parse_num(m.group(4)) or 0; sub=parse_num(m.group(5)) or cant*p
        desc=clean(rest[:m.start()]+" "+marca+" "+rest[m.end():])
        out.append(item(fn,"Ingeniería Boggio",cot,nro,cod,marca,desc,unidad,cant,p,sub,"USD",21,"Confirmar stock" if "(*)" in b else "","No informado","Consultar","boggio_v6"))
    return out or parse_table_like(fn,pages,"Ingeniería Boggio")

def parse_any(fn,pages):
    text="\n".join(pages); prov=provider(fn,text)
    if prov=="Ateco": return parse_ateco(fn,pages)
    if prov=="Ingeniería Boggio": return parse_boggio(fn,pages)
    if prov=="Marlew":
        out=parse_table_like(fn,pages,"Marlew")
        for x in out: x["marca"]="MARLEW"; x["parser"]="marlew_v6"
        return out
    return parse_table_like(fn,pages,prov)

def normalize_items(items):
    for it in items:
        src=" ".join([str(it.get("codigo","")),str(it.get("grupo_comparable","")),str(it.get("descripcion","")),str(it.get("formacion",""))])
        g=group(src)
        if g: it["grupo_comparable"]=g
        if "03492" in it.get("proveedor","") or norm(it.get("proveedor","")).startswith("l de la torre"): it["proveedor"]="Proveedor Cables"
    return items

def compare_items(items):
    items=normalize_items(items); df=pd.DataFrame(items)
    if df.empty: return [],"No hay ítems para comparar."
    df["grupo_comparable"]=df["grupo_comparable"].fillna("").astype(str).str.strip()
    df=df[df["grupo_comparable"]!=""]
    comps=[]; summary=[]
    for gr,g in df.groupby("grupo_comparable"):
        if g["proveedor"].astype(str).str.lower().nunique()<2: continue
        note="Monedas distintas: validar tipo de cambio." if g["moneda"].astype(str).nunique()>1 else ""
        g=g.sort_values("precio_unitario"); best=g.iloc[0]; offers=[]
        for _,r in g.iterrows():
            d=r.to_dict(); d["dif_unit_vs_mejor"]=float(r["precio_unitario"]-best["precio_unitario"]); d["dif_total_vs_mejor"]=float(r["total_con_iva"]-best["total_con_iva"]); d["recomendado_precio"]=float(r["precio_unitario"])==float(best["precio_unitario"]); offers.append(d)
        comps.append({"grupo_comparable":gr,"mejor_proveedor":best["proveedor"],"moneda":best["moneda"],"mejor_precio_unitario":float(best["precio_unitario"]),"mejor_total_con_iva":float(best["total_con_iva"]),"nota":note,"ofertas":offers})
        summary.append(f"{gr}: mejor precio unitario {best['proveedor']} — {best['moneda']} {best['precio_unitario']:.2f}. {note} Validar equivalencia técnica, mínimos y entrega.")
    if not summary:
        counts=df.groupby("grupo_comparable")["proveedor"].nunique().sort_values(ascending=False).to_dict()
        shown=", ".join([f"{k} ({v} prov.)" for k,v in list(counts.items())[:25]])
        return comps,"No hay grupos comparables con 2 o más proveedores. Revisá los grupos. Detectados: "+shown
    return comps,"\n".join(summary)

def build_excel(items,comps,summary):
    wb=Workbook(); thin=Side(style="thin",color="D9D9D9"); fill=PatternFill("solid",fgColor="1F4E78"); font=Font(color="FFFFFF",bold=True); green=PatternFill("solid",fgColor="C6EFCE")
    ws0=wb.active; ws0.title="Resumen Ejecutivo"; ws0["A1"]="Resumen Ejecutivo - Comparativa de Cotizaciones"; ws0["A1"].font=Font(size=16,bold=True); ws0["A3"]=summary or "Sin resumen."; ws0["A3"].alignment=Alignment(wrap_text=True,vertical="top"); ws0.column_dimensions["A"].width=120
    ws0["A6"]="Cantidad de ítems detectados"; ws0["B6"]=len(items); ws0["A7"]="Comparaciones encontradas"; ws0["B7"]=len(comps); ws0["A8"]="Proveedores detectados"; ws0["B8"]=len(set([i.get("proveedor","") for i in items]))
    ws=wb.create_sheet("Items Detectados"); headers=["Proveedor","Cotización","N°","Código","Marca","Descripción","Grupo comparable","Moneda","Cant. pedida","Cant. real","Unidad","P. unit.","Subtotal s/IVA","IVA %","IVA monto","Total c/IVA","Mínimo","Entrega","Notas","Parser"]
    for c,h in enumerate(headers,1):
        cell=ws.cell(1,c,h); cell.fill=fill; cell.font=font; cell.alignment=Alignment(horizontal="center",wrap_text=True); cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
    for r,it in enumerate(items,2):
        vals=[it.get("proveedor",""),it.get("cotizacion",""),it.get("nro_item",""),it.get("codigo",""),it.get("marca",""),it.get("descripcion",""),it.get("grupo_comparable",""),it.get("moneda",""),it.get("cantidad_pedida",0),it.get("cantidad_real",0),it.get("unidad",""),it.get("precio_unitario",0),it.get("subtotal_sin_iva",0),it.get("iva_pct",0),it.get("iva_monto",0),it.get("total_con_iva",0),it.get("minimo_compra",""),it.get("entrega",""),it.get("notas",""),it.get("parser","")]
        for c,v in enumerate(vals,1):
            cell=ws.cell(r,c,v); cell.border=Border(left=thin,right=thin,top=thin,bottom=thin); cell.alignment=Alignment(wrap_text=True,vertical="top")
            if c in [12,13,15,16]: cell.number_format='$ #,##0.00'
    for i,w in enumerate([20,18,8,18,16,65,24,10,14,14,10,14,16,10,14,16,20,20,42,16],1): ws.column_dimensions[get_column_letter(i)].width=w
    ws2=wb.create_sheet("Comparaciones"); h2=["Grupo","Proveedor","Descripción","Moneda","Cant. real","P. unit.","Subtotal","IVA","Total","Dif. unit.","Dif. total","Estado","Nota"]
    for c,h in enumerate(h2,1):
        cell=ws2.cell(1,c,h); cell.fill=fill; cell.font=font; cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
    row=2
    for comp in comps:
        for off in comp.get("ofertas",[]):
            vals=[comp.get("grupo_comparable",""),off.get("proveedor",""),off.get("descripcion",""),off.get("moneda",""),off.get("cantidad_real",0),off.get("precio_unitario",0),off.get("subtotal_sin_iva",0),off.get("iva_monto",0),off.get("total_con_iva",0),off.get("dif_unit_vs_mejor",0),off.get("dif_total_vs_mejor",0),"MEJOR PRECIO" if off.get("recomendado_precio") else "Alternativa",comp.get("nota","")]
            for c,v in enumerate(vals,1):
                cell=ws2.cell(row,c,v); cell.border=Border(left=thin,right=thin,top=thin,bottom=thin); cell.alignment=Alignment(wrap_text=True,vertical="top")
                if c in [6,7,8,9,10,11]: cell.number_format='$ #,##0.00'
                if c==12 and v=="MEJOR PRECIO": cell.fill=green
            row+=1
    for i,w in enumerate([24,20,65,10,14,14,16,14,16,14,14,18,36],1): ws2.column_dimensions[get_column_letter(i)].width=w
    out=io.BytesIO(); wb.save(out); out.seek(0); return out

@app.post("/api/analyze")
async def analyze(files: List[UploadFile]=File(...)):
    all_items=[]; raw=[]
    for f in files:
        data=await f.read(); fn=f.filename or "archivo"
        if not fn.lower().endswith(".pdf"): continue
        try:
            pages=pdf_pages(data); text="\n".join(pages); prov=provider(fn,text); items=parse_any(fn,pages); items=normalize_items(items)
            all_items.extend(items); raw.append({"archivo":fn,"tabla":f"Texto PDF - {prov}","columns":["Texto"],"rows":[[line] for line in text.splitlines()[:700]]})
        except Exception as e:
            raw.append({"archivo":fn,"tabla":"ERROR","columns":["Error"],"rows":[[str(e)]]})
    return {"items":all_items,"raw_tables":raw,"warnings":["V6: normalización técnica y comparación corregidas."]}

@app.post("/api/normalize")
async def normalize(payload: dict):
    return {"items":normalize_items(payload.get("items",[]))}

@app.post("/api/compare")
async def compare(payload: dict):
    comps,summary=compare_items(payload.get("items",[]))
    return {"comparisons":comps,"summary":summary}

@app.post("/api/export_excel")
async def export_excel(payload: dict):
    items=normalize_items(payload.get("items",[])); comps=payload.get("comparisons",[]); summary=payload.get("summary","")
    if not comps: comps,summary=compare_items(items)
    out=build_excel(items,comps,summary)
    return StreamingResponse(out,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",headers={"Content-Disposition":"attachment; filename=comparativa_cotizaciones_v6.xlsx"})

@app.get("/api/health")
def health():
    return {"status":"ok","version":"v6-industrial-fix-total"}
