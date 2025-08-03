from flask import Flask, render_template, request, send_file
from docx import Document
import io
import os
from datetime import datetime

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("form.html")

@app.route("/generate", methods=["POST"])
def generate():
    doc = Document("templates/form_template.docx")
    form = request.form

    # --- 1. 先處理一般單值欄位（包括空載數據）---
    for para in doc.paragraphs:
        for key in form:
            if f"{{{{{key}}}}}" in para.text:
                para.text = para.text.replace(f"{{{{{key}}}}}", form[key])

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key in form:
                        if f"{{{{{key}}}}}" in paragraph.text:
                            paragraph.text = paragraph.text.replace(f"{{{{{key}}}}}", form[key])

    # --- 2. 特別處理加載的5組數據 ---
    # 加載數據的欄位前綴
    load_fields = ['time', 'rpm', 'hz', 'kw', 'voltage_rs', 'voltage_st', 
                  'voltage_tr', 'current_r', 'current_s', 'current_t']
    
    # 處理每組加載數據 (1~5組)
    for group_num in range(1, 6):
        for field in load_fields:
            form_key = f"load_{field}_{group_num}"
            placeholder = f"{{{{load_{field}_{group_num}}}}}"
            replacement = form.get(form_key, "")
            
            # 在段落中替換
            for para in doc.paragraphs:
                if placeholder in para.text:
                    para.text = para.text.replace(placeholder, replacement)
            
            # 在表格中替換
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            if placeholder in paragraph.text:
                                paragraph.text = paragraph.text.replace(placeholder, replacement)

    # --- 3. 生成最終文件 ---
    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)
    filename = f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    return send_file(output_stream, as_attachment=True, download_name=filename)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
