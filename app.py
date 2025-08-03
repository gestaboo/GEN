from flask import Flask, render_template, request, send_file
from docx import Document
import io
import os
import re
from datetime import datetime

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("form.html")

def clean_placeholder(text):
    """清理佔位符中的特殊字符"""
    return re.sub(r'[\\_{}]', '', text)

def process_document(doc, form_data):
    """安全處理文檔替換"""
    # 先處理所有段落
    for para in doc.paragraphs:
        for key, value in form_data.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in para.text:
                para.text = para.text.replace(placeholder, value)
    
    # 處理表格（簡化版，避免深度解析）
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                # 直接處理整個單元格文本
                cell_text = cell.text
                for key, value in form_data.items():
                    placeholder = f"{{{{{key}}}}}"
                    if placeholder in cell_text:
                        cell.text = cell.text.replace(placeholder, value)
                
                # 如果需要保留格式，可以處理段落
                for para in cell.paragraphs:
                    for key, value in form_data.items():
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder in para.text:
                            para.text = para.text.replace(placeholder, value)

@app.route("/generate", methods=["POST"])
def generate():
    try:
        doc = Document("templates/form_template.docx")
        form = request.form

        # 準備替換數據（清理鍵名並處理空值）
        replacements = {}
        for key in form:
            clean_key = clean_placeholder(key)
            replacements[clean_key] = form.get(key, "")

        # 特別處理加載數據
        for i in range(1, 6):
            for field in ['time', 'rpm', 'hz', 'kw', 
                        'voltage_rs', 'voltage_st', 'voltage_tr',
                        'current_r', 'current_s', 'current_t']:
                form_key = f"load_{field}_{i}"
                clean_key = clean_placeholder(form_key)
                replacements[clean_key] = form.get(form_key, "")

        # 執行文檔處理
        process_document(doc, replacements)

        # 生成輸出文件
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
        return "生成文件時發生錯誤，請檢查模板格式", 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
