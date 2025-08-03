from flask import Flask, render_template, request, send_file
from docx import Document
import io
import os
from datetime import datetime

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("form.html")

def safe_replace_in_document(doc, placeholder, replacement):
    """安全替換函數，避免無限遞迴"""
    # 處理段落
    for para in doc.paragraphs:
        if placeholder in para.text:
            para.text = para.text.replace(placeholder, replacement)
    
    # 處理表格 - 使用更安全的方式
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                # 直接處理cell.text而不是遞迴處理paragraphs
                if placeholder in cell.text:
                    cell.text = cell.text.replace(placeholder, replacement)
                
                # 如果需要保留格式，可以這樣處理
                for para in cell.paragraphs:
                    if placeholder in para.text:
                        para.text = para.text.replace(placeholder, replacement)

@app.route("/generate", methods=["POST"])
def generate():
    try:
        doc = Document("templates/form_template.docx")
        form = request.form

        # 1. 處理空載數據
        no_load_fields = [
            'time', 'rpm', 'hz', 'kw',
            'voltage_rs', 'voltage_st', 'voltage_tr',
            'current_r', 'current_s', 'current_t'
        ]
        
        for field in no_load_fields:
            placeholder = f"{{{{no_load_{field}}}}}"
            replacement = form.get(f"no_load_{field}", "")
            safe_replace_in_document(doc, placeholder, replacement)

        # 2. 處理加載數據 (5組)
        load_fields = no_load_fields  # 相同欄位結構
        
        for group_num in range(1, 6):
            for field in load_fields:
                placeholder = f"{{{{load_{field}_{group_num}}}}}"
                replacement = form.get(f"load_{field}_{group_num}", "")
                safe_replace_in_document(doc, placeholder, replacement)

        # 生成文件
        output_stream = io.BytesIO()
        doc.save(output_stream)
        output_stream.seek(0)
        
        filename = f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        return send_file(
            output_stream,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        
    except Exception as e:
        app.logger.error(f"生成文件時發生錯誤: {str(e)}")
        return "生成文件時發生錯誤，請檢查日誌", 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
