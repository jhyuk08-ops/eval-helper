import os
import json
import io
import re
from flask import Flask, request, jsonify, render_template, send_file
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)

# 데이터 폴더 경로
DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')

def is_generic_name(name):
    """영역명이 기본값(수행평가1, 수행평가2 등)이거나 비어있는지 검사"""
    if not name:
        return True
    clean = re.sub(r'\s+', '', str(name))
    generic_list = ['수행평가1', '수행평가2', '수행평가3', '수행평가4', '수행평가', '영역1', '영역2', '영역3']
    if clean in generic_list or (clean.startswith('수행평가') and len(clean) <= 6):
        return True
    return False

def build_score_array(max_s, min_s, step):
    """최고점, 최저점, 급간을 바탕으로 정확한 점수 배열 생성"""
    try:
        max_s = int(max_s)
        min_s = int(min_s)
        step = int(step) if int(step) > 0 else 1
    except Exception:
        max_s, min_s, step = 5, 1, 2

    if min_s > max_s:
        max_s, min_s = min_s, max_s

    scores = []
    curr = max_s
    while curr >= min_s:
        scores.append(curr)
        curr -= step

    if not scores or scores[-1] != min_s:
        if not scores or scores[-1] > min_s:
            # 최소점 보장
            if min_s not in scores:
                scores.append(min_s)

    return scores

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

        # 사용자 입력 데이터 전처리 및 급간 점수 계산
        processed_perf_list = []
        for perf in perf_list:
            base_s = int(perf.get('baseScore', 0) or 0)
            user_elems = perf.get('evalElements', [])
            
            elem_info_list = []
            element_max_sum = 0
            element_min_sum = 0
            
            if user_elems:
                for elem in user_elems:
                    e_name = elem.get('name', '').strip()
                    m_s = int(elem.get('maxScore', 5) or 5)
                    l_s = int(elem.get('minScore', 1) or 1)
                    st = int(elem.get('step', 2) or 2)
                    
                    score_arr = build_score_array(m_s, l_s, st)
                    element_max_sum += m_s
                    element_min_sum += l_s
                    
                    elem_info_list.append({
                        "name": e_name if e_name else "자동생성 필요",
                        "target_scores": score_arr,
                        "maxScore": m_s,
                        "minScore": l_s,
                        "step": st
                    })
            else:
                arr1 = build_score_array(5, 1, 2)
                element_max_sum = 5
                element_min_sum = 1
                elem_info_list = [
                    {"name": "탐구 수행 및 결과 보고", "target_scores": arr1, "maxScore": 5, "minScore": 1, "step": 2}
                ]

            total_max = element_max_sum + base_s
            total_min = element_min_sum + base_s
            
            processed_perf_list.append({
                "name": perf.get('name', ''),
                "ratio": perf.get('ratio', 0),
                "baseScore": base_s,
                "totalMaxScore": total_max,
                "totalMinScore": total_min,
                "standards": perf.get('standards', []),
                "elements": elem_info_list
            })

        # Gemini 3.6 Flash 모델 활용
        model = genai.GenerativeModel('gemini-3.6-flash')

        prompt = f"""
당신은 2022 개정 교육과정 평가 전문가입니다. 아래 입력된 정보를 바탕으로 학교 평가계획서 세부 항목을 작성해주세요.

[기본 정보]
- 학교급: {'중학교' if school_type == 'middle' else '고등학교'}
- 학년/학기: {grade}학년 {semester}
- 교과: {main_subject} ({subject_title})

[수행평가 입력 데이터]
{json.dumps(processed_perf_list, ensure_ascii=False, indent=2)}

[작성 규칙]
1. 절대 '~합니다', '~입니다' 등의 경어체를 쓰지 마시고, 반드시 **'~한다', '~함', '~이다', '~수 있다'** (개조식/평서문) 문체만 사용하세요.
2. 영역명(`name`)이 '수행평가1' 등 기본값이거나 비어있으면 성취기준에 알맞은 전문적 영역명으로 직접 작명하세요.
3. 각 평가요소의 `target_scores` 배열(예: [5, 3, 1])의 개수와 **정확히 1:1 대응하는 `criteria` (채점기준 설명문구)**를 순서대로 작성하십시오.
   - `target_scores`가 [5, 4, 3, 2, 1]이면 `criteria` 설명 문장도 5개가 작성되어야 합니다.
   - 요소명(`name`)이 '자동생성 필요'인 경우 성취기준에 걸맞은 적절한 평가요소명을 만들어 넣어주세요.

반드시 **오직 순수한 JSON 데이터만** 출력하십시오:
{{
  "purposes": [
    "교과 핵심 개념과 원리를 정확히 이해하고 문제해결 능력을 함양한다.",
    "실생활 문제 해결 과정에서 교과의 유용성과 가치를 파악한다.",
    "자기주도적 학습 태도와 올바른 가치관을 형성한다."
  ],
  "directions": [
    "인지적 발달 수준에 맞춰 교과 역량을 균형 있게 평가한다.",
    "과정 중심 수행평가를 강화하여 구체적이고 적시적인 피드백을 제공한다."
  ],
  "policies": [
    "성적반영 비율은 정기시험 {exam_ratios[0]+exam_ratios[1]}%, 수행평가 {100-(exam_ratios[0]+exam_ratios[1])}%로 한다.",
    "정기시험은 학기당 2회 실시한다.",
    "서·논술형 평가는 수업 중에 실시하며 사전 공지된 명확한 기준에 따른다."
  ],
  "evaluations": [
    {{
      "index": 0,
      "name": "성취기준에 맞는 구체적 영역명",
      "level_high": "[상] 성취기준의 핵심 개념을 정확히 이해하고 논리적으로 적용할 수 있다.",
      "level_mid": "[중] 개념을 대체로 이해하고 있으나 과정에 일부분 오류가 존재한다.",
      "level_low": "[하] 기초적인 내용 파악이 부족하여 과제를 제대로 해결하지 못한다.",
      "rubrics": [
        {{
          "element": "세부 평가 요소명",
          "criteria": [
            "최고점 채점기준 문장 (~함 또는 ~한다 체)",
            "다음 점수 채점기준 문장",
            "최저점 채점기준 문장"
          ]
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

        # Word 문서 생성
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

        # AI 결과 매핑 준비
        eval_list = ai_data.get('evaluations', []) if ai_data else []

        # 4. 성취기준 및 평가기준(상, 중, 하) (표 형태)
        doc.add_heading("4. 성취기준 및 평가기준(상, 중, 하)", level=1)
        t4 = doc.add_table(rows=1, cols=3)
        t4.style = 'Table Grid'
        h4 = t4.rows[0].cells
        h4[0].paragraphs[0].add_run("수행평가 영역명 (비율/만점)").bold = True
        h4[1].paragraphs[0].add_run("성취기준").bold = True
        h4[2].paragraphs[0].add_run("성취수준 (상 / 중 / 하)").bold = True

        resolved_names = []

        for idx, proc_perf in enumerate(processed_perf_list):
            raw_name = proc_perf.get('name', '').strip()
            p_ratio = proc_perf.get('ratio', 0)
            p_stds = proc_perf.get('standards', [])
            total_max = proc_perf.get('totalMaxScore', 20)
            
            ai_eval = eval_list[idx] if idx < len(eval_list) else {}
            ai_name = ai_eval.get('name', '').strip()

            if is_generic_name(raw_name) and ai_name:
                display_name = ai_name
            else:
                display_name = raw_name if raw_name else (ai_name if ai_name else f"수행평가 영역 {idx+1}")

            resolved_names.append(display_name)

            row_cells = t4.add_row().cells
            row_cells[0].paragraphs[0].add_run(f"{display_name}\n({p_ratio}%, 만점 {total_max}점)")
            
            std_p = row_cells[1].paragraphs[0]
            for std in p_stds:
                std_p.add_run(f"{std}\n")
                
            lvl_p = row_cells[2].paragraphs[0]
            lh = ai_eval.get('level_high', '[상] 성취기준의 핵심 개념을 완벽히 이해하고 적용할 수 있다.')
            lm = ai_eval.get('level_mid', '[중] 핵심 개념을 이해하나 작성 과정에 일부분 미흡함이 있다.')
            ll = ai_eval.get('level_low', '[하] 기초 내용 이해가 미흡하여 핵심 과제를 완성하지 못한다.')
            lvl_p.add_run(f"{lh}\n{lm}\n{ll}")

        # 5. 수행평가 세부기준 (웹페이지에서 입력받은 점수 완벽 출력)
        doc.add_heading("5. 수행평가 세부기준", level=1)
        sub_labels = ['가', '나', '다', '라', '마', '바']

        for idx, proc_perf in enumerate(processed_perf_list):
            display_name = resolved_names[idx]
            p_ratio = proc_perf.get('ratio', 0)
            base_s = proc_perf.get('baseScore', 0)
            total_max = proc_perf.get('totalMaxScore', 20)
            total_min = proc_perf.get('totalMinScore', 8)
            sub_lbl = sub_labels[idx] if idx < len(sub_labels) else f"({idx+1})"
            
            doc.add_paragraph().add_run(
                f"{sub_lbl}. [{display_name}] 채점 기준표 (반영비율: {p_ratio}%, 영역 만점: {total_max}점 / 기본점수: {base_s}점)"
            ).bold = True
            
            t5 = doc.add_table(rows=1, cols=3)
            t5.style = 'Table Grid'
            h5 = t5.rows[0].cells
            h5[0].paragraphs[0].add_run("평가 요소").bold = True
            h5[1].paragraphs[0].add_run("세부 채점 기준").bold = True
            h5[2].paragraphs[0].add_run("배점").bold = True

            ai_eval = eval_list[idx] if idx < len(eval_list) else {}
            rubrics = ai_eval.get('rubrics', [])
            p_elements = proc_perf.get('elements', [])

            for r_idx, e_info in enumerate(p_elements):
                r_cells = t5.add_row().cells
                
                # 1) 평가요소명
                elem_name = e_info.get('name', '')
                if not elem_name or elem_name == "자동생성 필요":
                    if r_idx < len(rubrics) and rubrics[r_idx].get('element'):
                        elem_name = rubrics[r_idx].get('element')
                    else:
                        elem_name = f"평가 요소 {r_idx + 1}"
                r_cells[0].paragraphs[0].add_run(elem_name)

                # 2) 사용자가 설정한 정확한 급간 점수 배열
                target_scores = e_info.get('target_scores', [5, 3, 1])
                
                # AI가 작성한 채점 문구 매핑
                ai_criteria = rubrics[r_idx].get('criteria', []) if r_idx < len(rubrics) else []
                
                crit_p = r_cells[1].paragraphs[0]
                score_p = r_cells[2].paragraphs[0]
                
                for s_idx, sc in enumerate(target_scores):
                    c_text = ai_criteria[s_idx] if s_idx < len(ai_criteria) else "해당 기준에 맞게 수행함"
                    c_clean = re.sub(r'^[•\-\*\d점\s:]+', '', str(c_text)).strip()
                    crit_p.add_run(f"• [{sc}점] {c_clean}\n")
                    score_p.add_run(f"{sc}점\n")

            # 기본점수 안내 문구
            if base_s > 0:
                note_p = doc.add_paragraph()
                note_p.add_run(
                    f"  ※ 기본점수 {base_s}점이 부여되며, 세부 평가요소 득점 합계와 합산하여 최종 점수를 산출함. (최고점: {total_max}점, 최저점: {total_min}점)"
                )

        # 파일 스트림 전달
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
