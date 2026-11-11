"""
Bill to Spreadsheet Converter - Minimal Backend
Run with: uvicorn main:app --reload
Requires: pip install fastapi uvicorn python-multipart pdfplumber openpyxl
"""

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import pdfplumber
import openpyxl
from openpyxl.styles import Font, Alignment
import tempfile
import os
import re
from datetime import datetime
import zipfile
from io import BytesIO

app = FastAPI(title="Bill to Spreadsheet Converter")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def extract_bill_data(pdf_path: str) -> list[dict]:
    rows = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue
                
                # Detect simple 2-column tables (e.g. Amount | Due Date)
                cleaned_table = []
                for row in table:
                    cells = [str(c).strip() if c else "" for c in row]
                    cells = [c for c in cells if c]
                    if len(cells) >= 2:
                        cleaned_table.append(cells)
                
                if not cleaned_table:
                    continue
                
                # If it looks like a simple amount/date or key/value table
                header = cleaned_table[0]
                is_simple_table = len(header) == 2 and any(
                    kw in " ".join(header).lower() for kw in ["amount", "date", "due", "payment"]
                )
                
                if is_simple_table and len(cleaned_table) > 1:
                    # Use proper headers
                    col1, col2 = header[0], header[1]
                    for data_row in cleaned_table[1:]:
                        if len(data_row) >= 2:
                            rows.append({
                                "field": col1,
                                "value": data_row[0],
                                "extra": data_row[1],
                                "raw": " | ".join(data_row)
                            })
                    continue
                
                # Generic table handling
                for row in cleaned_table:
                    if len(row) >= 3 and row[0].isdigit() and len(row[0]) <= 3:
                        indicator = row[1]
                        value = " ".join(row[2:])
                        if len(value) > 1:
                            rows.append({
                                "field": indicator[:100],
                                "value": value[:200],
                                "raw": " | ".join(row)[:250]
                            })
                    elif len(row) == 2:
                        key, value = row[0], row[1]
                        if key.isdigit() and len(key) <= 3:
                            continue
                        if len(value) > 1:
                            rows.append({
                                "field": key[:100],
                                "value": value[:200],
                                "raw": " | ".join(row)[:250]
                            })
                    elif len(row) > 2:
                        text = " | ".join(row)
                        if len(text) > 10:
                            rows.append({
                                "field": row[0][:100],
                                "value": " ".join(row[1:])[:200],
                                "raw": text[:250]
                            })
            
            # Text fallback
            if not rows:
                text = page.extract_text() or ""
                for line in text.split("\n"):
                    line = line.strip()
                    if len(line) < 8:
                        continue
                    if ":" in line:
                        parts = line.split(":", 1)
                        if len(parts) == 2 and len(parts[1].strip()) > 1:
                            rows.append({
                                "field": parts[0].strip()[:100],
                                "value": parts[1].strip()[:200],
                                "raw": line[:250]
                            })
    
    # Deduplicate
    seen = set()
    unique = []
    for r in rows:
        key = (r.get("field","")[:60], r.get("value","")[:60], r.get("extra",""))
        if key not in seen and (r.get("value") or r.get("extra") or len(r.get("field","")) > 5):
            seen.add(key)
            unique.append(r)
    
    return unique[:100]


def create_excel(rows: list[dict], output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extracted Data"
    
    # Dynamic headers
    has_extra = any(r.get("extra") for r in rows)
    headers = ["Field", "Value"]
    if has_extra:
        headers.append("Date")
    headers.append("Raw")
    
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
    
    for i, row in enumerate(rows, 2):
        ws.cell(row=i, column=1, value=row.get("field", ""))
        ws.cell(row=i, column=2, value=row.get("value", ""))
        col = 3
        if has_extra:
            ws.cell(row=i, column=col, value=row.get("extra", ""))
            col += 1
        ws.cell(row=i, column=col, value=row.get("raw", ""))
    
    for col in ws.columns:
        max_len = 0
        letter = col[0].column_letter
        for cell in col:
            try:
                max_len = max(max_len, len(str(cell.value or "")))
            except:
                pass
        ws.column_dimensions[letter].width = min(max_len + 2, 80)
    
    wb.save(output_path)


@app.post("/convert")
async def convert_bills(files: list[UploadFile] = File(...)):
    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded")
    
    if len(files) > 10:
        raise HTTPException(status_code=400, detail="Maximum 10 files at a time")
    
    results = []
    temp_files = []
    
    try:
        for file in files:
            if not file.filename.lower().endswith(".pdf"):
                raise HTTPException(status_code=400, detail=f"{file.filename} is not a PDF")
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                content = await file.read()
                tmp.write(content)
                tmp_path = tmp.name
                temp_files.append(tmp_path)
            
            rows = extract_bill_data(tmp_path)
            if not rows:
                rows = [{"field": "No data extracted", "value": "", "raw": "Could not parse this document"}]
            
            excel_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
            temp_files.append(excel_path)
            create_excel(rows, excel_path)
            
            base_name = os.path.splitext(file.filename)[0]
            results.append((f"{base_name}.xlsx", excel_path))
        
        # Single file: return Excel directly
        if len(results) == 1:
            filename, path = results[0]
            return FileResponse(
                path,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=filename
            )
        
        # Multiple files: return a ZIP
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for filename, path in results:
                zf.write(path, filename)
        zip_buffer.seek(0)
        
        return StreamingResponse(
            zip_buffer,
            media_type="application/zip",
            headers={"Content-Disposition": 'attachment; filename="billsheet_exports.zip"'}
        )
    
    finally:
        for p in temp_files:
            if os.path.exists(p):
                try:
                    os.unlink(p)
                except:
                    pass


@app.get("/")
async def root():
    return {"message": "Bill to Spreadsheet Converter API", "usage": "POST PDFs to /convert"}
