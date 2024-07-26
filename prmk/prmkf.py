from flask import Flask, render_template, request, send_file, abort
from pptx_creator.presentation import create_presentation
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/create_presentation', methods=['POST'])
def create_presentation_route():
    try:
        slide_count = int(request.form['slide_count'])
        head_title = request.form['head_title']
        subtitle = request.form['subtitle']
        section0 = request.form['section0']
        section1 = request.form['section1']

        # 하위 항목을 리스트로 수집
        subsections1 = [request.form[f'subsection1_{i}'] for i in range(1, 11) if f'subsection1_{i}' in request.form]
        subsections2 = [request.form[f'subsection2_{i}'] for i in range(1, 11) if f'subsection2_{i}' in request.form]
        subsections3 = [request.form[f'subsection3_{i}'] for i in range(1, 11) if f'subsection3_{i}' in request.form]
        subsections4 = [request.form[f'subsection4_{i}'] for i in range(1, 11) if f'subsection4_{i}' in request.form]
        
        section2 = request.form['section2']
        section3 = request.form['section3']
        section4 = request.form['section4']
        add_additional_page = int(request.form['add_additional_page'])
        save_path = 'presentations'  # 저장 경로 설정
        os.makedirs(save_path, exist_ok=True)
        output_file = os.path.join(save_path, 'proposal.pptx')
        
        # 하위 항목 리스트를 create_presentation 함수에 전달
        create_presentation(slide_count, save_path, head_title, subtitle, section0, section1, subsections1, section2, subsections2, section3, subsections3, section4, subsections4, add_additional_page)
        
        # 파일이 실제로 존재하는지 확인
        if not os.path.exists(output_file):
            abort(404, description="File not found")
        
        return send_file(output_file, as_attachment=True)
    except ValueError:
        return "유효한 숫자를 입력하세요.", 400

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)