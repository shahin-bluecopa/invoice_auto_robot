import logging
import os
import re
import json
import pandas as pd
from docx import Document
from docx2pdf import convert

# ---------- USER CONFIG ----------
# Kept as per original script, though ideally should be config-driven
TEMPLATE_PATH = r"C:\Users\Admin\OneDrive\Desktop\invoice_automation\Input\invoice_template.docx"
EXCEL_PATH    = r"C:\Users\Admin\OneDrive\Desktop\invoice_automation\Input\invoice_fill_data.xlsx"
OUTPUT_DIR    = r"C:\Users\Admin\OneDrive\Desktop\invoice_automation\Output"

def is_blank(text: str) -> bool:
    """Detects whether a paragraph or cell is effectively blank."""
    return bool(re.fullmatch(r"[\s_\-.]*", str(text or "")))

def excel_to_json(excel_path: str, json_path: str, logger: logging.Logger) -> list[dict]:
    """Convert Excel sheet to JSON list."""
    df = pd.read_excel(excel_path, dtype=str).fillna("")
    records = df.to_dict(orient="records")
    os.makedirs(os.path.dirname(json_path), exist_ok=True)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    logger.info(f"JSON file created: {json_path}")
    return records

def fill_docx_no_labels(template_path: str, record: dict, out_docx: str):
    """Fill a template sequentially with Excel row data (no placeholders)."""
    doc = Document(template_path)
    values = list(record.values())
    idx = 0
    # Fill paragraphs
    for p in doc.paragraphs:
        if idx >= len(values):
            break
        if is_blank(p.text):
            p.text = str(values[idx])
            idx += 1
    # Fill table cells
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                if idx >= len(values):
                    break
                if is_blank(c.text):
                    c.text = str(values[idx])
                    idx += 1
            if idx >= len(values):
                break
        if idx >= len(values):
            break
    os.makedirs(os.path.dirname(out_docx), exist_ok=True)
    doc.save(out_docx)
    return out_docx

def docx_to_pdf(docx_path: str, pdf_path: str, logger: logging.Logger):
    """Convert Word to PDF using Microsoft Word (via docx2pdf)."""
    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
    try:
        convert(docx_path, pdf_path)
    except Exception as e:
        logger.warning(f"Could not convert to PDF: {e}")

def process_invoices(logger: logging.Logger):
    """Main function to process invoices - converts Excel data to PDF invoices."""
    # Ensure output directory exists
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    json_path = os.path.join(OUTPUT_DIR, "invoice_data.json")

    # Step 1: Excel -> JSON
    logger.info("Step 1: Converting Excel to JSON")
    records = excel_to_json(EXCEL_PATH, json_path, logger)

    # Step 2: For each row -> DOCX -> PDF
    logger.info("Generating filled invoices...")
    for i, record in enumerate(records, start=1):
        inv_no = record.get("Invoice_No", f"INV-{i:04d}")
        safe_name = "".join(ch for ch in inv_no if ch.isalnum() or ch in "-_")
        out_docx = os.path.join(OUTPUT_DIR, f"{safe_name}.docx")
        out_pdf  = os.path.join(OUTPUT_DIR, f"{safe_name}.pdf")
        
        fill_docx_no_labels(TEMPLATE_PATH, record, out_docx)
        docx_to_pdf(out_docx, out_pdf, logger)
        
        logger.info(f"Generated: {safe_name}.pdf")

    logger.info(f"All PDFs and JSON saved in: {OUTPUT_DIR}")

