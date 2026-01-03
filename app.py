from flask import Flask, render_template, request, send_file
import pdfplumber
import pandas as pd
import re
import os
import uuid

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
COST_PRICE = 150  # change if needed

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def extract(pattern, text):
    match = re.search(pattern, text, re.S | re.I)
    return match.group(1).strip() if match else ""

@app.route("/", methods=["GET", "POST"])
def upload_pdf():
    if request.method == "POST":
        pdf_file = request.files["pdf"]
        if not pdf_file:
            return "No file uploaded"

        filename = str(uuid.uuid4()) + ".pdf"
        pdf_path = os.path.join(UPLOAD_FOLDER, filename)
        pdf_file.save(pdf_path)

        records = []
        sr_no = 1

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                customer = extract(r"Customer Address\n([A-Za-z ].+)", text)
                state = extract(r",\s*([A-Za-z &]+),\s*\d{6}", text)
                order_no = extract(r"Order No\.\n([\w_]+)", text)
                invoice_no = extract(r"Invoice No\.\n([\w\d]+)", text)
                sku = extract(r"Product Details\nSKU.*?\n([\w\-]+)", text)
                color = extract(r"Qty\s+Color\s+Order No\.\n.*?\s(\w+)\s[\w_]+", text)

                taxable = extract(r"Taxable Value\s*\nRs\.([\d\.]+)", text)
                taxable = float(taxable) if taxable else 0

                payment = "COD" if "COD:" in text else "Prepaid"

                if "Valmo Pickup" in text:
                    pickup = "Valmo"
                elif "Delhivery" in text:
                    pickup = "Delhivery"
                elif "Shadowfax" in text:
                    pickup = "Shadowfax"
                else:
                    pickup = ""

                pickup_code = extract(r"(VL\d+|SF\d+)", text)
                manifest_date = extract(r"Pickup\s(\d{2}/\d{2})", text)

                profit = taxable - COST_PRICE
                margin = round((profit / taxable) * 100, 2) if taxable else 0

                records.append([
                    sr_no, manifest_date, order_no, customer, state,
                    pickup, pickup_code, sku, color, invoice_no,
                    taxable, profit, margin, payment
                ])
                sr_no += 1

        columns = [
            "Sr. No.", "Manifest Date", "Order No.", "Customer Name", "State",
            "Pickup", "Pickup Code", "SKU", "Color", "Invoice No.",
            "Taxable Value", "Profit", "Margin (%)", "Payment Option"
        ]

        df = pd.DataFrame(records, columns=columns)

        excel_name = str(uuid.uuid4()) + ".xlsx"
        excel_path = os.path.join(OUTPUT_FOLDER, excel_name)
        df.to_excel(excel_path, index=False)

        return send_file(excel_path, as_attachment=True)

    return render_template("upload.html")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
