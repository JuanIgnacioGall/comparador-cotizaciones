
from fastapi import FastAPI, File, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from typing import List
from pathlib import Path
import io, re
import pandas as pd

app = FastAPI(title="Comparador Web de Cotizaciones")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def parse_number(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return None
    s = re.sub(r"[^0-9,.-]", "", s)
    if not s:
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts[-1]) <= 2:
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        parts = s.split(".")
        if len(parts) > 2:
            s = s.replace(".", "")
    try:
        return float(s)
    except:
        return None

def clean(x):
    if x is None:
        return ""
    if str(x).lower() == "nan":
        return ""
    return str(x).strip()

def norm(x):
    return re.sub(r"\s+", " ", str(x or "").lower().strip())

def currency(x):
    x = str(x or "").upper().strip()
    if x in ["USD", "U$S", "US$"]:
        return "USD"
    if x in ["ARS", "$", "PESOS", "PESO"]:
        return "ARS"
    if x in ["EUR", "€"]:
        return "EUR"
    return x or "SIN DEFINIR"

def detect_col(df, keys):
    cols = list(df.columns)
    normalized = [norm(c) for c in cols]
    for k in keys:
        for i, c in enumerate(normalized):
            if k in c:
                return cols[i]
    return None

def looks_header(row):
    text = " ".join(norm(x) for x in row if clean(x))
    keys = ["descripcion", "descripción", "detalle", "producto", "material", "cantidad", "precio", "unitario", "subtotal", "iva", "total", "codigo", "código", "cod"]
    return sum(1 for k in keys if k in text) >= 2

def promote_header(raw):
    raw = raw.dropna(how="all").reset_index(drop=True)
    if raw.empty:
        return raw
    header_idx = None
    for i in range(min(20, len(raw))):
        if looks_header(raw.iloc[i].tolist()):
            header_idx = i
            break
    if header_idx is not None:
        headers = [clean(x) or f"Col_{j+1}" for j, x in enumerate(raw.iloc[header_idx].tolist())]
        df = raw.iloc[header_idx+1:].copy()
        df.columns = headers[:len(df.columns)]
    else:
        df = raw.copy()
        df.columns = [f"Col_{j+1}" for j in range(len(df.columns))]
    df = df.dropna(how="all")
    return df.reset_index(drop=True)

def read_excel(data):
    xls = pd.ExcelFile(io.BytesIO(data))
    out = []
    for sheet in xls.sheet_names:
        raw = pd.read_excel(io.BytesIO(data), sheet_name=sheet, header=None, dtype=object)
        out.append((sheet, promote_header(raw), ""))
    return out

def read_csv(data):
    text = None
    for enc in ["utf-8", "latin-1", "cp1252"]:
        try:
            text = data.decode(enc)
            break
        except:
            pass
    if text is None:
        text = data.decode("utf-8", errors="ignore")
    best = None
    best_cols = 0
    for sep in [";", ",", "\t", "|"]:
        try:
            df = pd.read_csv(io.StringIO(text), sep=sep, dtype=object)
            if len(df.columns) > best_cols:
                best = df
                best_cols = len(df.columns)
        except:
            pass
    if best is None:
        best = pd.DataFrame({"Texto": [l for l in text.splitlines() if l.strip()]})
    return [("CSV", best, text)]

def text_rows(text):
    rows = []
    money_pat = r"(USD|U\$S|US\$|\$|ARS|EUR|€)?\s*([\d]{1,3}(?:[\.,]\d{3})*(?:[\.,]\d{1,2})|[\d]+(?:[\.,]\d{1,2})?)"
    for line in [l.strip() for l in text.splitlines() if l.strip()]:
        low = norm(line)
        if len(line) < 8 or any(x in low for x in ["subtotal", "total general", "cuit", "fecha", "telefono"]):
            continue
        matches = list(re.finditer(money_pat, line, flags=re.I))
        nums = []
        for m in matches:
            n = parse_number(m.group(2))
            if n is not None:
                nums.append((m.group(1) or "", n, m.group(0)))
        if not nums:
            continue
        mon, price, raw_price = nums[-1]
        qty = 1
        q = re.search(r"(cant\.?|cantidad|qty)\s*[:\-]?\s*(\d+[\.,]?\d*)", line, flags=re.I)
        if q:
            qty = parse_number(q.group(2)) or 1
        else:
            q = re.match(r"^\s*(\d+[\.,]?\d*)\s+", line)
            if q:
                possible = parse_number(q.group(1))
                if possible and possible < 100000:
                    qty = possible
        desc = line.replace(raw_price, "")
        desc = re.sub(r"(cant\.?|cantidad|qty)\s*[:\-]?\s*\d+[\.,]?\d*", "", desc, flags=re.I)
        desc = re.sub(r"^\s*\d+[\.,]?\d*\s+", "", desc).strip(" -|:")
        if len(desc) < 3:
            continue
        rows.append({"Descripcion": desc, "Cantidad": qty, "Precio Unitario": price, "Moneda": currency(mon)})
    return rows

def read_pdf(data):
    out = []
    all_text = ""
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for pageno, page in enumerate(pdf.pages, 1):
                text = page.extract_text() or ""
                all_text += "\n" + text
                tables = page.extract_tables() or []
                for tidx, table in enumerate(tables, 1):
                    if table and len(table) >= 2:
                        df = promote_header(pd.DataFrame(table))
                        out.append((f"PDF_P{pageno}_T{tidx}", df, text))
    except Exception as e:
        out.append(("ERROR_PDF", pd.DataFrame({"Error": [str(e)]}), str(e)))
    if not out:
        rows = text_rows(all_text)
        if rows:
            out.append(("PDF_TEXTO", pd.DataFrame(rows), all_text))
        else:
            out.append(("PDF_SIN_TABLA", pd.DataFrame({"Observacion": ["No se detectó tabla utilizable. Puede ser PDF escaneado."]}), all_text))
    return out

def standardize(df, filename, sheet, text):
    proveedor = Path(filename).stem.replace("_", " ").replace("-", " ")
    col_codigo = detect_col(df, ["codigo", "código", "cod", "sku", "referencia"])
    col_desc = detect_col(df, ["descripcion", "descripción", "detalle", "producto", "material", "concepto", "texto"])
    col_cant = detect_col(df, ["cantidad", "cant", "qty"])
    col_moneda = detect_col(df, ["moneda"])
    col_unit = detect_col(df, ["precio unitario", "p.unit", "p unit", "unitario", "precio"])
    col_subtotal = detect_col(df, ["subtotal", "sin iva", "neto", "importe"])
    col_iva = detect_col(df, ["iva"])
    col_total = detect_col(df, ["total", "con iva"])
    rows = []
    iva_default = 10.5 if ("10,5" in text or "10.5" in text) else 21
    for _, r in df.iterrows():
        codigo = clean(r.get(col_codigo)) if col_codigo else ""
        desc = clean(r.get(col_desc)) if col_desc else ""
        if not desc and not codigo:
            parts = [clean(v) for v in r.tolist() if clean(v) and parse_number(v) is None]
            desc = " ".join(parts[:4])
        if not desc and not codigo:
            continue
        cant = parse_number(r.get(col_cant)) if col_cant else None
        cant = cant if cant and cant > 0 else 1
        mon = currency(r.get(col_moneda)) if col_moneda else "SIN DEFINIR"
        unit = parse_number(r.get(col_unit)) if col_unit else None
        subtotal = parse_number(r.get(col_subtotal)) if col_subtotal else None
        iva = parse_number(r.get(col_iva)) if col_iva else iva_default
        total = parse_number(r.get(col_total)) if col_total else None
        if iva is None:
            iva = iva_default
        if unit is None and subtotal is not None and cant:
            unit = subtotal / cant
        if unit is None and total is not None and cant:
            unit = (total / (1 + iva / 100)) / cant
        if unit is None:
            nums = [parse_number(v) for v in r.tolist()]
            nums = [x for x in nums if x is not None and x > 0]
            if nums:
                unit = nums[-1]
        if unit is None:
            continue
        subtotal_calc = cant * unit
        iva_monto = subtotal_calc * iva / 100
        total_calc = subtotal_calc + iva_monto
        if total is not None and total > 0:
            total_calc = total
            subtotal_calc = total_calc / (1 + iva / 100)
            iva_monto = total_calc - subtotal_calc
            unit = subtotal_calc / cant
        rows.append({
            "archivo": filename, "tabla": sheet, "proveedor": proveedor, "codigo": codigo,
            "descripcion": desc or codigo, "grupo_comparable": codigo, "moneda": mon,
            "cantidad": cant, "precio_unitario": unit, "subtotal_sin_iva": subtotal_calc,
            "iva_pct": iva, "iva_monto": iva_monto, "total_con_iva": total_calc,
            "estado": "ok" if codigo else "sin_grupo"
        })
    return rows

def compare(items):
    df = pd.DataFrame(items)
    if df.empty:
        return [], "No se detectaron ítems."
    df["grupo_comparable"] = df["grupo_comparable"].fillna("").astype(str).str.strip()
    df = df[df["grupo_comparable"] != ""]
    comps = []
    summaries = []
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
            d["recomendado_precio"] = bool(r["precio_unitario"] == best["precio_unitario"])
            offers.append(d)
        comps.append({"grupo_comparable": group, "mejor_proveedor": best["proveedor"], "ofertas": offers})
        summaries.append(f"{group}: mejor precio unitario {best['proveedor']} - {best['moneda']} {best['precio_unitario']:.2f}. Total c/IVA {best['moneda']} {best['total_con_iva']:.2f}.")
    if not comps:
        return [], "Se detectaron ítems, pero no hay códigos/grupos repetidos entre proveedores. No se compara para evitar mezclar materiales distintos."
    return comps, "\n".join(summaries)

@app.post("/api/analyze")
async def analyze(files: List[UploadFile] = File(...)):
    all_items, raw_tables = [], []
    for f in files:
        data = await f.read()
        name = f.filename or "archivo"
        if name.lower().endswith((".xlsx", ".xls")):
            tables = read_excel(data)
        elif name.lower().endswith(".pdf"):
            tables = read_pdf(data)
        elif name.lower().endswith((".csv", ".txt")):
            tables = read_csv(data)
        else:
            tables = []
        for sheet, df, text in tables:
            raw_tables.append({"archivo": name, "tabla": sheet, "columns": [str(c) for c in df.columns], "rows": df.fillna("").astype(str).head(200).values.tolist()})
            all_items.extend(standardize(df, name, sheet, text))
    comps, summary = compare(all_items)
    return {"items": all_items, "comparisons": comps, "summary": summary, "raw_tables": raw_tables}

@app.get("/api/health")
def health():
    return {"status": "ok"}
