#updated invoice robot
import json
import logging
import os
from docxtpl import DocxTemplate
from docx import Document
from docx2pdf import convert

# --- Utilities ---

def to_float(value):
    return float(str(value).replace(",", "")) if value else 0.0

def format_inr(value):
    try:
        s = str(int(round(float(value))))
        if len(s) <= 3: return s
        rest, last3 = s[:-3], s[-3:]
        rest = ",".join([rest[max(i - 2, 0):i] for i in range(len(rest), 0, -2)][::-1])
        return f"{rest},{last3}"
    except Exception:
        return "0"

def fill_services_table(doc: Document, services: list):
    for table in doc.tables:
        if not table.rows: continue
        header = " ".join(c.text.lower() for c in table.rows[0].cells)
        if "description" in header and "amount" in header:
            if len(table.rows) > 1: table._tbl.remove(table.rows[1]._tr)
            for idx, s in enumerate(services, start=1):
                row = table.add_row().cells
                row[0].text = str(idx)
                row[1].text = s["description"]
                row[2].text = s.get("sacCode", "")
                row[3].text = s["qty_disp"]
                row[4].text = s["rate_disp"]
                row[5].text = s["amount_disp"]
            return

def select_template(inv, template_base):
    inv_type = inv.get("invoice", {}).get("type", "").strip().upper()
    gstin = (inv.get("supplier", {}).get("gstin") or "").strip().lower()
    
    # Check explicitly for BOS or Tax Invoice
    if inv_type in ("BOS", "BILL_OF_SUPPLY"): return os.path.join(template_base, "BOS_AllPlaceholders.docx")
    if inv_type in ("TAX", "TAX_INVOICE"): return os.path.join(template_base, "Invoice_Template_AllPlaceholders.docx")
    
    # Fallback logic
    is_unregistered = gstin in ("", "unregistered", "na", "n/a")
    tax = inv.get("taxDetails", {})
    total_gst = sum(to_float(tax.get(k)) for k in ["cgstAmount", "sgstAmount", "igstAmount"])
    
    if is_unregistered or total_gst == 0:
        return os.path.join(template_base, "BOS_AllPlaceholders.docx")
    return os.path.join(template_base, "Invoice_Template_AllPlaceholders.docx")

# --- Logic Handlers ---

def process_tax_logic(inv):
    """Applies GST, Display Flags, and TDS logic in one go."""
    supplier, buyer, tax = inv["supplier"], inv["buyer"], inv["taxDetails"]
    gstin = (supplier.get("gstin") or "").strip().lower()
    is_unregistered = gstin in ("", "unregistered")
    same_state = supplier.get("stateCode") == buyer.get("stateCode")

    # 1. GST Calculations
    if is_unregistered:
        tax.update({"cgstRate": 0, "cgstAmount": 0, "sgstRate": 0, "sgstAmount": 0, "igstRate": 0, "igstAmount": 0, "totalTax": 0})
        tax["grandTotal"] = tax["taxableAmount"]
    else:
        if same_state:
            tax.update({"igstRate": 0, "igstAmount": 0, "totalTax": tax["cgstAmount"] + tax["sgstAmount"]})
        else:
            tax.update({"cgstRate": 0, "cgstAmount": 0, "sgstRate": 0, "sgstAmount": 0, "totalTax": tax["igstAmount"]})
        tax["grandTotal"] = tax["taxableAmount"] + tax["totalTax"]

    # 2. Display Flags
    total_gst = tax["cgstAmount"] + tax["sgstAmount"] + tax["igstAmount"]
    inv["display"] = {
        "show_gst_section": not is_unregistered and total_gst > 0,
        "show_cgst": not is_unregistered and same_state and tax["cgstAmount"] > 0,
        "show_sgst": not is_unregistered and same_state and tax["sgstAmount"] > 0,
        "show_igst": not is_unregistered and not same_state and tax["igstAmount"] > 0,
        "show_tds": inv.get("display", {}).get("show_tds", False),
    }

    # 3. TDS Logic
    if inv["display"]["show_tds"]:
        tds_amt = round(tax["taxableAmount"] * 0.10)
        inv["tdsDetails"] = {"tdsIncomeTax": {"rate": 10, "amount": tds_amt}, "totalTdsDeducted": tds_amt, "netPayable": tax["grandTotal"] - tds_amt}
    else:
        inv["tdsDetails"] = {"tdsIncomeTax": {"rate": 10, "amount": 0}, "totalTdsDeducted": 0, "netPayable": tax["grandTotal"]}

def normalize_and_format(inv):
    # Services
    for s in inv["services"]:
        qty, rate = to_float(s["qty"]), to_float(s["rate"])
        s.update({"qty_disp": format_inr(qty), "rate_disp": format_inr(rate), "amount_disp": format_inr(qty * rate)})
    
    # Tax
    tax = inv["taxDetails"]
    for k in ["taxableAmount", "cgstAmount", "sgstAmount", "igstAmount"]:
        tax[k] = to_float(tax.get(k, 0))
    
    process_tax_logic(inv)
    
    # Final Formatting
    for k in ["taxableAmount", "cgstAmount", "sgstAmount", "igstAmount", "totalTax", "grandTotal"]:
        tax[k] = format_inr(tax[k])
    
    tds = inv["tdsDetails"]
    tds["tdsIncomeTax"]["amount"] = format_inr(tds["tdsIncomeTax"]["amount"])
    tds["totalTdsDeducted"] = format_inr(tds["totalTdsDeducted"])
    tds["netPayable"] = format_inr(tds["netPayable"])

def render_docx(template_path, data, out_path):
    tpl = DocxTemplate(template_path)
    tpl.render(data)
    tpl.save(out_path)
    doc = Document(out_path)
    fill_services_table(doc, data["services"])
    doc.save(out_path)

# --- Main ---

def process_invoices(logger: logging.Logger, config):
    INPUT, OUTPUT, TEMPLATES = config["input_folder"], config["output_folder"], config["template_folder"]
    os.makedirs(OUTPUT, exist_ok=True)

    with open(os.path.join(INPUT, "invoice_data_all_scenario.json"), "r", encoding="utf-8") as f:
        data = json.load(f)

    for inv in data["invoiceSamples"]:
        normalize_and_format(inv)
        
        template = select_template(inv, TEMPLATES)
        out_docx = os.path.join(OUTPUT, f"{inv['invoice']['number'].replace('/', '-')}.docx")
        render_docx(template, inv, out_docx)

    try:
        logger.info(f"Starting batch conversion in {OUTPUT}...")
        convert(OUTPUT)
        logger.info("Batch conversion completed. Cleaning up DOCX files...")
        for f in [f for f in os.listdir(OUTPUT) if f.lower().endswith(".docx")]:
             try: os.remove(os.path.join(OUTPUT, f))
             except Exception as e: logger.warning(f"Failed to delete {f}: {e}")
    except Exception as e:
        logger.error(f"Batch conversion failed: {e}")

    logger.info("All invoices processed successfully")
