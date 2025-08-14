from flask import Flask, render_template, request, send_file
import pdfplumber
import openpyxl
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["pdf_file"]
        if file.filename == "":
            return "No file selected!"
        
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(pdf_path)

        excel_filename = filename.rsplit(".", 1)[0] + ".xlsx"
        excel_path = os.path.join(UPLOAD_FOLDER, excel_filename)

        convert_pdf_to_excel(pdf_path, excel_path)
        return send_file(excel_path, as_attachment=True)

    return render_template("index.html")


def convert_pdf_to_excel(pdf_path, excel_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PDF Tables"

    with pdfplumber.open(pdf_path) as pdf:
        row_num = 1
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                for row in table:
                    ws.append(row)
            else:
                text = page.extract_text()
                if text:
                    for line in text.split("\n"):
                        ws.cell(row=row_num, column=1).value = line
                        row_num += 1

    wb.save(excel_path)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
