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
        
        possible_paths = [
            os.path.join(DATA_DIR, 'evaluation standard', 'middle', f'{main_subject}.json'),
            os.path.join(DATA_DIR, 'evaluation standard', 'middle', f'{target_subject}.json'),
            os.path.join(DATA_DIR, 'evaluation standard', 'high', main_subject, f'{detail_subject}.json'),
            os.path.join(DATA_DIR, 'evaluation standard', 'high', main_subject, f'{target_subject}.json'),
            os.path.join(DATA_DIR, 'evaluation standard', 'high', f'{target_subject}.json'),
            os.path.join(DATA_DIR, f'{target_subject}.json')
        ]
        
        for file_path in possible_paths:
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as f:
                    standards_data = json.load(f)
                return jsonify({"success": True, "standards": standards_data})
        
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

        genai.configure(api_key=user_api_key)

        school_type = request.form.get('schoolType', '')
        grade = request.form.get('grade', '')
        semester = request.form.get('semester', '')
        main_subject = request.form.get('교과', '')
        detail_subject = request.form.get('세부교과', '')
        
        exam_ratios = json.loads(request.form.get('examRatios', '[30, 30]'))
        exam_standards = json.loads(request.form.get('examStandards', '{}'))
        perf_list = json.loads(request.form.get('perfList', '[]'))

        subject_title = detail_subject if detail_subject else main_subject

        # Gemini 3.6 Flash 모델 활용
        model = genai.GenerativeModel('gemini-3.6-flash')

        prompt = f"""
당신은 2022 개정 교육과정 평가 전문가입니다. 아래 제출된 정보를 바탕으로 평가계획서 세부 항목을 구체적인 교사 어조로 작성해주세요.

[기본 정보]
- 학교급: {'중학교' if school_type == 'middle' else '고등학교'}
- 학년/학기: {grade}학년 {semester}
- 교과: {main_subject} ({subject_title})

[수행평가 정보]
{json.dumps(perf_list, ensure_ascii=False, indent=2)}

반드시 **오직 순수한 JSON 형식만** 출력해주세요 (마크다운 ```json 및 설명 문구 절대 금지):
{{
  "purposes": [
    "평가 목적 가항목 (수학적 개념/원리 이해 및 문제해결 능력...)",
    "평가 목적 나항목 (실생활 연결 및 유용성 인식...)",
    "평가 목적 다항목 (흥미, 자신감 및 학습 역량...)"
  ],
  "directions": [
    "인지적 발달 수준 고려 및 교과 역량 균형 평가",
    "과정 중심 수행평가 강화 및 맞춤형 피드백 제공",
    "서술형, 포트폴리오, 공학도구 활용 등 다양한 평가방법 적용"
  ],
  "policies": [
    "성적반영 비율 설정 내용 (정기시험 비율 및 수행평가 비율 언급)",
    "정기시험 횟수 언급",
    "서·논술형 평가 수업 중 실시 및 평가 방식",
    "객관적인 채점기준 사전 마련",
    "교과협의회를 통한 신뢰성 있는 평가",
    "평가 결과 피드백 및 수업 개선 활용"
  ],
  "evaluations": [
    {{
      "name": "수행평가 영역명",
      "level_high": "[상] 해당 영역의 최고 수준 성취 기준...",
      "level_mid": "[중] 해당 영역의 보통 수준 성취 기준...",
      "level_low": "[하] 해당 영역의 미흡 수준 성취 기준...",
      "rubrics": [
        {{
          "element": "세부 평가 요소명",
          "criteria": [
            "최상 수준 채점 기준 상세 설명",
            "우수 수준 채점 기준 상세 설명",
            "보통 수준 채점 기준 상세 설명",
            "미흡 수준 채점 기준 상세 설명"
          ],
          "scores": ["10점", "8점", "5점", "2점"]
        }}
      ]
    }}
  ]
}}
"""

        response = model.generate_content(prompt)
        ai_data = None
        if response and response.text:
            raw_text = response.text.strip()
            if raw_text.startswith("```json"):
                raw_text = raw_text.split("```json", 1)[1]
            if raw_text.startswith("```"):
                raw_text = raw_text.split("```", 1)[1]
            if raw_text.endswith("```"):
                raw_text = raw_text.rsplit("```", 1)[0]
            try:
                ai_data = json.loads(raw_text.strip())
            except Exception as pe:
                print("JSON 파싱 에러:", pe)

        # Word 문서 생성 시작
        doc = Document()

        # 문서 제목
        title_p = doc.add_paragraph()
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title_p.add_run(f"2022 개정 교육과정 {subject_title} 평가계획서")
        run.font.name = '맑은 고딕'
        run.font.size = Pt(18)
        run.font.bold = True

        # 1. 기본 정보 (표 형태)
        doc.add_heading("1. 기본 정보", level=1)
        t1 = doc.add_table(rows=3, cols=2)
        t1.style = 'Table Grid'
        
        t1.cell(0, 0).paragraphs[0].add_run("학교급 및 학년").bold = True
        t1.cell(0, 1).paragraphs[0].add_run(f"{'중학교' if school_type == 'middle' else '고등학교'} {grade}학년 {semester}")
        
        t1.cell(1, 0).paragraphs[0].add_run("정기시험 비율").bold = True
        t1.cell(1, 1).paragraphs[0].add_run(f"1차 정기시험: {exam_ratios[0]}%, 2차 정기시험: {exam_ratios[1]}%")
        
        t1.cell(2, 0).paragraphs[0].add_run("정기시험 성취기준").bold = True
        p_std1 = t1.cell(2, 1).paragraphs[0]
        p_std1.add_run("■ 1차 정기시험:\n").bold = True
        for std in exam_standards.get('1차 정기시험', []):
            p_std1.add_run(f"  - {std}\n")
        p_std1.add_run("\n■ 2차 정기시험:\n").bold = True
        for std in exam_standards.get('2차 정기시험', []):
            p_std1.add_run(f"  - {std}\n")

        # 2. 평가의 목적
        doc.add_heading("2. 평가의 목적", level=1)
        purposes = ai_data.get('purposes', []) if ai_data else []
        labels = ['가', '나', '다', '라', '마']
        if purposes:
            for idx, p_text in enumerate(purposes):
                lbl = labels[idx] if idx < len(labels) else f"({idx+1})"
                clean_p = p_text.replace(f"{lbl}.", "").strip()
                p = doc.add_paragraph()
                p.add_run(f"  {lbl}. {clean_p}")
        else:
            doc.add_paragraph("  가. 교과 핵심 개념과 원리를 이해하고 문제해결 능력을 함양한다.")

        # 3. 평가의 기본 방향과 방침
        doc.add_heading("3. 평가의 기본 방향과 방침", level=1)
        p_dir = doc.add_paragraph()
        p_dir.add_run("  가. 기본 방향\n").bold = True
        directions = ai_data.get('directions', []) if ai_data else []
        for idx, d_text in enumerate(directions, 1):
            clean_d = d_text.replace(f"{idx})", "").strip()
            p_dir.add_run(f"    {idx}) {clean_d}\n")

        p_pol = doc.add_paragraph()
        p_pol.add_run("  나. 방침\n").bold = True
        policies = ai_data.get('policies', []) if ai_data else []
        for idx, pol_text in enumerate(policies, 1):
            clean_pol = pol_text.replace(f"{idx})", "").strip()
            p_pol.add_run(f"    {idx}) {clean_pol}\n")

        # 4. 성취기준 및 평가기준(상, 중, 하) (표 형태)
        doc.add_heading("4. 성취기준 및 평가기준(상, 중, 하)", level=1)
        t4 = doc.add_table(rows=1, cols=3)
        t4.style = 'Table Grid'
        h4 = t4.rows[0].cells
        h4[0].paragraphs[0].add_run("수행평가 영역명 (비율)").bold = True
        h4[1].paragraphs[0].add_run("성취기준").bold = True
        h4[2].paragraphs[0].add_run("성취수준 (상 / 중 / 하)").bold = True

        eval_list = ai_data.get('evaluations', []) if ai_data else []
        eval_dict = {e.get('name'): e for e in eval_list}

        for perf in perf_list:
            p_name = perf.get('name', '')
            p_ratio = perf.get('ratio', 0)
            p_stds = perf.get('standards', [])
            ai_eval = eval_dict.get(p_name, {})

            row_cells = t4.add_row().cells
            row_cells[0].paragraphs[0].add_run(f"{p_name}({p_ratio}%)")
            
            # 성취기준 목록
            std_p = row_cells[1].paragraphs[0]
            for std in p_stds:
                std_p.add_run(f"{std}\n")
                
            # 성취수준 (상/중/하)
            lvl_p = row_cells[2].paragraphs[0]
            lh = ai_eval.get('level_high', '성취기준을 완벽히 이해하고 적절히 적용할 수 있다.')
            lm = ai_eval.get('level_mid', '성취기준을 대체로 이해하나 일부 서술에 미흡함이 있다.')
            ll = ai_eval.get('level_low', '성취기준 이해가 부족하여 핵심 내용을 완성하지 못한다.')
            lvl_p.add_run(f"{lh}\n{lm}\n{ll}")

        # 5. 수행평가 세부기준 (영역별 세부 채점기준표 표 형태)
        doc.add_heading("5. 수행평가 세부기준", level=1)
        sub_labels = ['가', '나', '다', '라', '마', '바']

        for idx, perf in enumerate(perf_list):
            p_name = perf.get('name', '')
            p_ratio = perf.get('ratio', 0)
            sub_lbl = sub_labels[idx] if idx < len(sub_labels) else f"({idx+1})"
            
            doc.add_paragraph().add_run(f"{sub_lbl}. [{p_name}] 채점 기준표 (반영비율: {p_ratio}%)").bold = True
            
            t5 = doc.add_table(rows=1, cols=3)
            t5.style = 'Table Grid'
            h5 = t5.rows[0].cells
            h5[0].paragraphs[0].add_run("평가 요소").bold = True
            h5[1].paragraphs[0].add_run("세부 채점 기준").bold = True
            h5[2].paragraphs[0].add_run("배점").bold = True

            ai_eval = eval_dict.get(p_name, {})
            rubrics = ai_eval.get('rubrics', [])
            
            if rubrics:
                for r in rubrics:
                    r_cells = t5.add_row().cells
                    r_cells[0].paragraphs[0].add_run(r.get('element', ''))
                    
                    crit_p = r_cells[1].paragraphs[0]
                    for c in r.get('criteria', []):
                        bullet = c if c.startswith('•') else f"• {c}"
                        crit_p.add_run(f"{bullet}\n")
                        
                    score_p = r_cells[2].paragraphs[0]
                    for s in r.get('scores', []):
                        score_p.add_run(f"{s}\n")
            else:
                for elem in perf.get('evalElements', []):
                    r_cells = t5.add_row().cells
                    r_cells[0].paragraphs[0].add_run(elem.get('name', ''))
                    r_cells[1].paragraphs[0].add_run(f"• {elem.get('name')} 수행 능력이 우수함\n• 보통 수준\n• 노력 요함")
                    r_cells[2].paragraphs[0].add_run(f"최고 {elem.get('maxScore')}점\n최저 {elem.get('minScore')}점")

        # 메모리 스트림으로 파일 전달
        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)

        filename = f"평가계획서_{subject_title}.docx"
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
