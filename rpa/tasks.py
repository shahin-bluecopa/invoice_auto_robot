import json
import logging
import os

from docxtpl import DocxTemplate
from docx import Document
from docx2pdf import convert


# --------------------------------------------------
# Utility: Indian number formatting
# --------------------------------------------------
def format_inr(value):
    try:
        value = int(round(float(value)))
        s = str(value)
        if len(s) <= 3:
            return s
        last3 = s[-3:]
        rest = s[:-3]
        rest = ",".join(
            [rest[max(i - 2, 0):i] for i in range(len(rest), 0, -2)][::-1]
        )
        return f"{rest},{last3}"
    except Exception:
        return ""


# --------------------------------------------------
# Fill services table using python-docx (NO placeholders)
# --------------------------------------------------
def fill_services_table(doc: Document, services: list):
    for table in doc.tables:
        header = " ".join(c.text.lower() for c in table.rows[0].cells)
        if "description of service" in header and "amount" in header:

            # Remove placeholder row (row index 1)
            if len(table.rows) > 1:
                table._tbl.remove(table.rows[1]._tr)

            for idx, s in enumerate(services, start=1):
                row = table.add_row().cells
                row[0].text = str(idx)
                row[1].text = s.get("description", "")
                row[2].text = s.get("sacCode", "")
                row[3].text = s.get("qty", "")
                row[4].text = s.get("rate", "")
                row[5].text = s.get("amount", "")
            return


# --------------------------------------------------
# Validation: Invoice vs BOS logic
# --------------------------------------------------
def validate_invoice_vs_bos(inv, selected_template, logger):
    supplier = inv.get("supplier", {})
    buyer = inv.get("buyer", {})

    gstin = (supplier.get("gstin") or "").strip()
    supplier_state = supplier.get("stateCode")
    buyer_state = buyer.get("stateCode")

    if not supplier_state or not buyer_state:
        logger.warning("StateCode missing for supplier or buyer")

    if "Invoice_Template" in selected_template:
        if not gstin or gstin.lower() == "unregistered":
            logger.warning(
                f"Invoice {inv['invoice']['number']} generated without GSTIN"
            )

    if "BOS_" in selected_template:
        if gstin and gstin.lower() != "unregistered":
            logger.warning(
                f"BOS generated for registered supplier "
                f"(Invoice {inv['invoice']['number']})"
            )


# --------------------------------------------------
# Template selection
# --------------------------------------------------
def select_template(inv, base_path):
    supplier = inv.get("supplier", {})
    buyer = inv.get("buyer", {})

    gstin = (supplier.get("gstin") or "").strip()
    supplier_state = (supplier.get("stateCode") or "").strip()
    buyer_state = (buyer.get("stateCode") or "").strip()

    invoice_tpl = os.path.join(base_path, "Invoice_Template_AllPlaceholders.docx")
    bos_intra_tpl = os.path.join(base_path, "BOS_Intra_AllPlaceholders.docx")
    bos_inter_tpl = os.path.join(base_path, "BOS_Inter_AllPlaceholders.docx")

    if gstin and gstin.lower() != "unregistered":
        return invoice_tpl

    return bos_intra_tpl if supplier_state == buyer_state else bos_inter_tpl


# --------------------------------------------------
# Render DOCX (2-phase)
# --------------------------------------------------
def render_docx(template_path, data, out_path):
    # Phase 1: render normal placeholders
    tpl = DocxTemplate(template_path)
    tpl.render(data)
    tpl.save(out_path)

    # Phase 2: insert services safely
    doc = Document(out_path)
    fill_services_table(doc, data["services"])
    doc.save(out_path)


# --------------------------------------------------
# Main processing
# --------------------------------------------------
def process_invoices(logger: logging.Logger, config):
    INPUT_BASE = config["input_folder"]
    OUTPUT_DIR = config["output_folder"]
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    json_path = os.path.join(INPUT_BASE, "invoice_data.json")
    logger.info(f"Reading JSON from {json_path}")

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    invoices = data["invoiceSamples"]

    for inv in invoices:
        logger.info(
            f"Invoice {inv['invoice']['number']} → "
            f"{len(inv['services'])} services"
        )

        # ---- Normalize services
        for s in inv["services"]:
            qty = float(s.get("qty", 0))
            rate = float(s.get("rate", 0))
            s["qty"] = format_inr(qty)
            s["rate"] = format_inr(rate)
            s["amount"] = format_inr(qty * rate)

        # ---- Normalize tax
        tax = inv["taxDetails"]
        tax["taxableAmount"] = format_inr(tax.get("taxableAmount", 0))
        tax["cgstAmount"] = format_inr(tax.get("cgstAmount", 0))
        tax["sgstAmount"] = format_inr(tax.get("sgstAmount", 0))
        tax["totalTax"] = format_inr(tax.get("totalTax", 0))
        tax["grandTotal"] = format_inr(tax.get("grandTotal", 0))

        # ---- TDS safety
        if inv.get("tdsDetails") is None:
            inv["tdsDetails"] = {
                "tdsIncomeTax": {"amount": ""},
                "tdsGst": {"amount": ""},
                "totalTdsDeducted": "",
                "netPayable": tax["grandTotal"],
            }

        template = select_template(inv, INPUT_BASE)
        validate_invoice_vs_bos(inv, template, logger)

        invoice_no = inv["invoice"]["number"].replace("/", "-")
        out_docx = os.path.join(OUTPUT_DIR, f"{invoice_no}.docx")
        out_pdf = os.path.join(OUTPUT_DIR, f"{invoice_no}.pdf")

        render_docx(template, inv, out_docx)

        try:
            convert(out_docx, out_pdf)
        except Exception as e:
            logger.warning(f"PDF conversion skipped: {e}")

    logger.info("All invoices processed successfully")
