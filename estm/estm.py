from flask import Flask, render_template, request, send_file
import pandas as pd
from io import BytesIO

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/preview_excel', methods=['POST'])
def preview_excel():
    # 폼에서 행과 열 수 가져오기
    rows = int(request.form['rows'])
    cols = int(request.form['cols'])

    return render_template('index.html', preview=True, rows=rows, cols=cols)

@app.route('/create_excel', methods=['POST'])
def create_excel():
    # 폼에서 행과 열 수 가져오기
    rows = int(request.form['rows'])
    cols = int(request.form['cols'])

    # 입력된 데이터를 DataFrame으로 변환
    data = []
    for row in range(rows):
        row_data = []
        for col in range(cols):
            cell_name = f'cell_{row}_{col}'
            row_data.append(request.form.get(cell_name, ''))
        data.append(row_data)
    
    df = pd.DataFrame(data, columns=[f'Column {i+1}' for i in range(cols)])

    # 엑셀 파일을 메모리에 생성
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)

    return send_file(output, download_name="example.xlsx", as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
