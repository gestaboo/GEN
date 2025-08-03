from flask import Flask, render_template, request, send_file
from docx import Document
import io
import os
import re
from datetime import datetime

app = Flask(__name__)

@app.route("/", methods=['GET'])
def home():
    return render_template("form.html")

@app.route("/generate", methods=['POST'])
def generate_doc():
    try:
        # 檢查模板是否存在
        if not os.path.exists("templates/form_template.docx"):
            return "Word模板文件不存在", 404

        doc = Document("templates/form_template.docx")
        form_data = request.form

        # 除錯：印出接收到的表單數據
        app.logger.debug(f"收到的表單數據: {form_data}")

        # 處理所有可能的佔位符變體
        replacements = {}
        for key, value in form_data.items():
            clean_key = re.sub(r'[\\_{}\s]', '', key)  # 徹底清理鍵名
            replacements[f"{{{{{clean_key}}}}}"] = value
            replacements[f"{{{{{key}}}}}"] = value  # 保留原始格式

        # 特別處理加載數據
        for i in range(1, 6):
            for field in ['time', 'rpm', 'hz', 'kw', 'voltage_rs', 'voltage_st', 'voltage_tr', 'current_r', 'current_s', 'current_t']:
                placeholder = f"{{{{load_{field}_{i}}}}}"
                replacements[placeholder] = form_data.get(f"load_{field}_{i}", "")

        # 執行文檔替換
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    # 先處理整個單元格
                    cell_text = cell.text
                    for ph, val in replacements.items():
                        cell_text = cell_text.replace(ph, val)
                    cell.text = cell_text
                    
                    # 再處理段落保留格式
                    for para in cell.paragraphs:
                        para_text = para.text
                        for ph, val in replacements.items():
                            para_text = para_text.replace(ph, val)
                        para.text = para_text

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
        app.logger.error(f"文件生成失敗: {str(e)}", exc_info=True)
        return f"文件生成失敗: {str(e)}", 500

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000, debug=True)
