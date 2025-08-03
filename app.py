from flask import Flask, render_template, request, send_file
from docx import Document
import io
import os
from datetime import datetime

app = Flask(__name__)

# 增加超時時間設定
app.config['GENERATE_TIMEOUT'] = 300  # 5分鐘

@app.route("/")
def index():
    return render_template("form.html")

def replace_in_document(doc, placeholder, replacement):
    """安全的替換函數，避免無限循環"""
    max_iterations = 1000
    count = 0
    
    # 在段落中替換
    for para in doc.paragraphs:
        if placeholder in para.text:
            para.text = para.text.replace(placeholder, replacement)
            count += 1
            if count >= max_iterations:
                break
    
    # 在表格中替換
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if placeholder in para.text:
                        para.text = para.text.replace(placeholder, replacement)
                        count += 1
                        if count >= max_iterations:
                            break

@app.route("/generate", methods=["POST"])
def generate():
    try:
        doc = Document("templates/form_template.docx")
        form = request.form

        # 1. 先處理空載數據
        no_load_fields = [
            'time', 'rpm', 'hz', 'kw', 
            'voltage_rs', 'voltage_st', 'voltage_tr',
            'current_r', 'current_s', 'current_t'
        ]
        
        for field in no_load_fields:
            placeholder = f"{{{{no_load_{field}}}}}"
            replacement = form.get(f"no_load_{field}", "")
            replace_in_document(doc, placeholder, replacement)

        # 2. 處理加載數據 (5組)
        load_fields = no_load_fields  # 相同欄位結構
        
        for group_num in range(1, 6):
            for field in load_fields:
                placeholder = f"{{{{load_{field}_{group_num}}}}}"
                replacement = form.get(f"load_{field}_{group_num}", "")
                replace_in_document(doc, placeholder, replacement)

        # 3. 生成文件
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
        return f"生成文件時發生錯誤: {str(e)}", 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
