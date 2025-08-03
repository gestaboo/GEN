
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

    # 替換段落中的佔位符
    for para in doc.paragraphs:
        for key in form:
            if f"{{{{{key}}}}}" in para.text:
                para.text = para.text.replace(f"{{{{{key}}}}}", form[key])

    # 新增：替換表格中的佔位符
    for table in doc.tables:  # 遍歷所有表格
        for row in table.rows:  # 遍歷所有行
            for cell in row.cells:  # 遍歷所有儲存格
                for key in form:
                    if f"{{{{{key}}}}}" in cell.text:
                        cell.text = cell.text.replace(f"{{{{{key}}}}}", form[key])

    # 以下保持不變
    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)
    filename = f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    return send_file(output_stream, as_attachment=True, download_name=filename)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
