from flask import Flask, render_template, request, send_file
from docx import Document
import io
import os
import re
from datetime import datetime

app = Flask(__name__)

def clean_placeholder(key):
    """統一清理佔位符格式"""
    return key.replace('\\', '').replace(' ', '').strip('{}')

@app.route("/generate", methods=["POST"])
def generate():
    try:
        doc = Document("templates/form_template.docx")
        form = request.form

        # 準備替換字典（處理所有可能的佔位符變體）
        replacements = {}
        for key in form:
            # 處理常規欄位
            clean_key = clean_placeholder(key)
            replacements[f"{{{clean_key}}}"] = form[key]
            replacements[f"{{ {clean_key} }}"] = form[key]  # 處理帶空格的變體
            
            # 特別處理加載數據 (1-5組)
            if key.startswith('load_'):
                for i in range(1, 6):
                    numbered_key = key.replace('_1', f'_{i}')
                    clean_num_key = clean_placeholder(numbered_key)
                    replacements[f"{{{clean_num_key}}}"] = form.get(key, "")
        
        # 雙重替換確保完整處理
        for _ in range(2):  # 兩次處理確保嵌套替換
            for para in doc.paragraphs:
                for ph, val in replacements.items():
                    if ph in para.text:
                        para.text = para.text.replace(ph, val)
            
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        # 直接處理整個單元格
                        cell_text = cell.text
                        for ph, val in replacements.items():
                            if ph in cell_text:
                                cell.text = cell.text.replace(ph, val)
                        
                        # 處理段落保留格式
                        for para in cell.paragraphs:
                            for ph, val in replacements.items():
                                if ph in para.text:
                                    para.text = para.text.replace(ph, val)

        # 處理未填寫的欄位（清空剩餘的{{...}}）
        for para in doc.paragraphs:
            para.text = re.sub(r'\{\{.*?\}\}', '', para.text)
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    cell.text = re.sub(r'\{\{.*?\}\}', '', cell.text)
                    for para in cell.paragraphs:
                        para.text = re.sub(r'\{\{.*?\}\}', '', para.text)

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
        return f"文件生成錯誤: {str(e)}", 500
