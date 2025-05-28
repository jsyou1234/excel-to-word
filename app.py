import os
import pandas as pd
from docx import Document
from flask import Flask, request, render_template, send_file

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['excel_file']
        if file:
            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)

            df = pd.read_excel(filepath)

            generated_files = []
            for index, row in df.iterrows():
                doc = Document('template.docx')
                for p in doc.paragraphs:
                    for key in row.index:
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder in p.text:
                            p.text = p.text.replace(placeholder, str(row[key]))
                output_path = os.path.join(OUTPUT_FOLDER, f'output_{index+1}.docx')
                doc.save(output_path)
                generated_files.append(output_path)

            return send_file(generated_files[0], as_attachment=True)

    return render_template('index.html')

# 👇 이것이 꼭 필요합니다!
if __name__ == '__main__':
    app.run(debug=True)
