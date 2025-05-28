from flask import Flask, request, render_template, send_file
import os
import pandas as pd
from docx import Document

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/generate_oa', methods=['POST'])
def generate_oa():
    file = request.files['excel_file']
    if not file:
        return '엑셀 파일을 선택해주세요.'

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    # 엑셀에서 'YourRef' 아래 값 찾기
    df = pd.read_excel(filepath, header=None)
    ref_value = None
    for i in range(len(df)):
        for j in range(len(df.columns)):
            if str(df.iat[i, j]).strip() == "YourRef":
                ref_value = str(df.iat[i+1, j]).strip()
                break
        if ref_value:
            break

    if not ref_value:
        return '엑셀에서 YourRef 값을 찾을 수 없습니다.'

    # 템플릿 로드 및 치환
    template_path = 'templates/OA검토보고_이메일.docx'
    doc = Document(template_path)

    for p in doc.paragraphs:
        for run in p.runs:
            if 'your_ref' in run.text:
                run.text = run.text.replace('your_ref', ref_value)

    output_path = os.path.join(OUTPUT_FOLDER, 'OA검토보고_이메일_new.docx')
    doc.save(output_path)

    return send_file(output_path, as_attachment=True)


# 👇 이것이 꼭 필요합니다!
if __name__ == '__main__':
    app.run(debug=True)
