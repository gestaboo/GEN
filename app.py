
from flask import Flask, render_template, request, send_file
from docx import Document
import io
from datetime import datetime

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("form.html")

@app.route("/generate", methods=["POST"])
def generate():
    doc = Document("templates/form_template.docx")
    room = request.form.get("room", "")
    kw = request.form.get("kw", "")
    reason = request.form.get("reason", "")

    for para in doc.paragraphs:
        if "機房" in para.text:
            para.text = para.text.replace("機房", f"機房：{room}")
        if "kW" in para.text:
            para.text = para.text.replace("kW", f"kW：{kw}")
        if "開機原因：" in para.text:
            para.text = f"開機原因：{reason}"

    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)

    filename = f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    return send_file(output_stream, as_attachment=True, download_name=filename)

if __name__ == "__main__":
    app.run(debug=True)
