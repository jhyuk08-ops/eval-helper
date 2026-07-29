import os
import json
import io
from flask import Flask, request, jsonify, render_template, send_file
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)

# 데이터 폴더 경로
DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/get_subjects', methods=['POST'])
def get_subjects():
    try:
        data = request.get_json() or {}
        school_type = data.get('schoolType', 'high')
        
        file_path = os.path.join(DATA_DIR, f'{school_type}_subjects.json')
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding='utf-8') as f:
                subject_data = json.load(f)
            return jsonify({"success": True, **subject_data})
        
        # 데이터 파일이 없을 경우 제공하는 기본값
        if school_type == 'middle':
            return jsonify({
                "success": True,
                "subjects": ["국어", "수학", "영어", "사회", "과학", "도덕", "기술·가정", "체육", "음악", "미술"]
            })
        else:
            return jsonify({
                "success": True,
                "subjects": ["국어과", "수학과", "영어과", "사회과", "과학과"],
                "hierarchy": {
                    "국어과": ["공통국어1", "공통국어2", "화법과 언어", "독서와 작문", "문학"],
                    "수학과": ["공통수학1", "공통수학2", "대수", "미적분I", "확률과 통계"],
                    "영어과": ["공통영어1", "공통영어2", "영어 회화", "영어 독해와 작문"],
                    "사회과": ["통합사회1", "통합사회2", "세계시민과 지리", "한국사1", "한국사2"],
                    "과학과": ["통합과학1", "통합과학2", "과학탐구실험1", "과학탐구실험2", "물리학", "화학", "생명과학", "지구과학"]
                }
            })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/get_standards', methods=['POST'])
def get_standards():
    try:
        data = request.get_json() or {}
        school_type = data.get('schoolType', 'high')
        main_subject = data.get('교과', '')
        detail_subject = data.get('세부교과', '')
        
        target_subject = detail_subject if detail_subject else main_subject
        
        # 중학교 / 고등학교 폴더 구조에 맞춘 경로 탐색
        possible_paths = [
            # 중학교: data/evaluation standard/middle/과학.json
            os.path.join(DATA_DIR, 'evaluation standard', 'middle', f'{main_subject}.json'),
            os.path.join(DATA_DIR, 'evaluation standard', 'middle', f'{target_subject}.json'),
            # 고등학교: data/evaluation standard/high/사회/경제.json
            os.path.join(DATA_DIR, 'evaluation standard', 'high', main_subject, f'{detail_subject}.json'),
            os.path.join(DATA_DIR, 'evaluation standard', 'high', main_subject, f'{target_subject}.json'),
            os.path.join(DATA_DIR, 'evaluation standard', 'high', f'{target_subject}.json'),
            # 백업용 경로
            os.path.join(DATA_DIR, f'{target_subject}.json')
        ]
        
        # 파일이 실제로 존재하는 경로를 순서대로 확인하여 읽기
        for file_path in possible_paths:
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as f:
                    standards_data = json.load(f)
                return jsonify({"success": True, "standards": standards_data})
        
        # 데이터 파일이 없을 경우 기본 제공 샘플 성취기준
        sample_standards = [
            f"[{target_subject}-01] 핵심 개념과 원리를 정확히 이해하고 관련 현상을 설명할 수 있다.",
            f"[{target_subject}-02] 다양한 탐구 방법과 자료를 활용하여 문제를 분석하고 해결할 수 있다.",
            f"[{target_subject}-03] 실생활 맥락에 개념을 적용하여 창의적인 대안을 도출할 수 있다.",
            f"[{target_subject}-04] 자신의 생각을 논리적으로 표현하고 타인과 협력적으로 소통할 수 있다.",
            f"[{target_subject}-05] 비판적 사고를 바탕으로 종합적인 평가와 가치 판단을 내릴 수 있다."
        ]
        return jsonify({"success": True, "standards": sample_standards})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/generate', methods=['POST'])
def generate():
    try:
        user_api_key = request.form.get('apiKey')
        if not user_api_key:
            return jsonify({"success": False, "error": "API 키가 전달되지 않았습니다."}), 400

        # 사용자가 입력한 API 키 적용
        genai.configure(api_key=user_api_key)

        school_type = request.form.get('schoolType', '')
        grade = request.form.get('grade', '')
        semester = request.form.get('semester', '')
        main_subject = request.form.get('교과', '')
        detail_subject = request.form.get('세부교과', '')
        
        exam_ratios = json.loads(request.form.get('examRatios', '[30, 30]'))
        exam_standards = json.loads(request.form.get('examStandards', '{}'))
        perf_list = json.loads(request.form.get('perfList', '[]'))

        # Gemini 2.5 Flash 모델 사용
        model = genai.GenerativeModel('gemini-2.5-flash')

        prompt = f"""
당신은 2022 개정 교육과정 평가 전문가입니다. 아래 평가계획 정보를 바탕으로 각 수행평가 영역별 상세 평가 기준 및 채점 루브릭(상/중/하)을 구체적인 교사 어조로 작성해주세요.

[기본 정보]
- 학교급: {'중학교' if school_type == 'middle' else '고등학교'}
- 학년/학기: {grade}학년 {semester}
- 교과: {main_subject} ({detail_subject})

[지필평가 비율 및 성취기준]
- 1차 정기시험 ({exam_ratios[0]}%): {', '.join(exam_standards.get('1차 정기시험', []))}
- 2차 정기시험 ({exam_ratios[1]}%): {', '.join(exam_standards.get('2차 정기시험', []))}

[수행평가 영역별 정보]
{json.dumps(perf_list, ensure_ascii=False, indent=2)}

각 수행평가 영역마다 아래 항목이 포함되도록 명확히 구분하여 작성해 주세요:
1. 평가 목적 및 방침
2. 세부 평가요소 및 배점표
3. 성취수준별(상/중/하) 구체적 채점기준 (루브릭)
"""

        response = model.generate_content(prompt)
        ai_text = response.text if response else "AI 내용 생성에 실패했습니다."

        # Word 문서 (.docx) 생성
        doc = Document()
        
        # 제목
        title_p = doc.add_paragraph()
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title_p.add_run(f"2022 개정 교육과정 [{detail_subject if detail_subject else main_subject}] 평가계획서")
        run.font.name = '맑은 고딕'
        run.font.size = Pt(18)
        run.font.bold = True

        # 1. 개요
        doc.add_heading("1. 기본 정보", level=1)
        p = doc.add_paragraph()
        p.add_run(f"• 학년/학기: {grade}학년 {semester}\n")
        p.add_run(f"• 대상 교과: {main_subject} - {detail_subject}\n")
        p.add_run(f"• 1차 정기시험 반영비율: {exam_ratios[0]}%\n")
        p.add_run(f"• 2차 정기시험 반영비율: {exam_ratios[1]}%\n")

        # 2. 성취기준
        doc.add_heading("2. 평가 대상 성취기준", level=1)
        p_exam = doc.add_paragraph()
        p_exam.add_run("[1차 정기시험 평가 성취기준]\n").bold = True
        for std in exam_standards.get('1차 정기시험', []):
            p_exam.add_run(f"- {std}\n")
        
        p_exam.add_run("\n[2차 정기시험 평가 성취기준]\n").bold = True
        for std in exam_standards.get('2차 정기시험', []):
            p_exam.add_run(f"- {std}\n")

        # 3. 수행평가 요약
        doc.add_heading("3. 수행평가 영역 요약", level=1)
        for idx, perf in enumerate(perf_list, 1):
            doc.add_heading(f"3.{idx} {perf.get('name')} (반영비율: {perf.get('ratio')}%)", level=2)
            p_perf = doc.add_paragraph()
            p_perf.add_run("• 선택 성취기준:\n").bold = True
            for std in perf.get('standards', []):
                p_perf.add_run(f"  - {std}\n")
            
            p_perf.add_run("• 세부 평가요소:\n").bold = True
            for elem in perf.get('evalElements', []):
                p_perf.add_run(f"  - {elem.get('name')}: 최고 {elem.get('maxScore')}점 / 최저 {elem.get('minScore')}점\n")

        # 4. AI 생성 상세 세부 기준
        doc.add_heading("4. 수행평가 세부 평가기준 (Gemini 2.5 Flash 생성)", level=1)
        doc.add_paragraph(ai_text)

        # 메모리 스트림으로 저장 후 반환
        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)

        filename = f"평가계획서_{detail_subject if detail_subject else main_subject}.docx"
        return send_file(
            file_stream,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        print("생성 중 에러 발생:", str(e))
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)