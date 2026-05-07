
from fastapi import FastAPI, File, UploadFile
                        nums = re.findall(r"\d+[\.,]?\d*", row_text)

                        if len(nums) >= 2:
                            try:
                                qty = float(nums[-2].replace(",", "."))
                            except:
                                qty = 1

                            try:
                                unit_price = float(nums[-1].replace(",", "."))
                            except:
                                unit_price = 0

                        items.append({
                            "archivo": filename,
                            "descripcion": row_text,
                            "cantidad": qty,
                            "precio_unitario": unit_price,
                            "subtotal": qty * unit_price,
                        })

            else:
                # fallback texto simple
                text = page.extract_text() or ""
                lines = text.split("\n")

                for line in lines:
                    line = clean(line)

                    if not valid_material(line):
                        continue

                    nums = re.findall(r"\d+[\.,]?\d*", line)

                    qty = 1
                    unit_price = 0

                    if len(nums) >= 2:
                        try:
                            qty = float(nums[-2].replace(",", "."))
                        except:
                            pass

                        try:
                            unit_price = float(nums[-1].replace(",", "."))
                        except:
                            pass

                    items.append({
                        "archivo": filename,
                        "descripcion": line,
                        "cantidad": qty,
                        "precio_unitario": unit_price,
                        "subtotal": qty * unit_price,
                    })

    return items


@app.post("/api/analyze")
async def analyze(files: List[UploadFile] = File(...)):

    all_items = []

    for file in files:
        data = await file.read()

        if file.filename.lower().endswith(".pdf"):
            items = extract_items_from_pdf(data, file.filename)
            all_items.extend(items)

    return {
        "items": all_items
    }


@app.get("/")
def home():
    return {
        "status": "ok"
    }
