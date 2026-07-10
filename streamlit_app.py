# -*- coding: utf-8 -*-
"""
LIFO 진단 결과 → PPT 자동 생성 (Streamlit)
"""
import os
import re
import io
from copy import deepcopy
from datetime import datetime
import streamlit as st
import openpyxl
from pypdf import PdfReader
from pptx import Presentation
from pptx.util import Pt, Emu
from lxml import etree

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.pptx')

# LIFO 유형 순서 (S/G, C/T, C/H, A/D)
LIFO_KEYS = ['SG', 'CT', 'CH', 'AD']

# 직급 키워드 (이름에서 제거 대상)
TITLE_KEYWORDS = [
    '책임', '선임', '수석', '팀장', '부장', '차장', '과장', '대리', '사원',
    '매니저', '파트장', '실장', '본부장', '상무', '전무', '이사',
]


# ── 엑셀 파싱 ───────────────────────────────────────────────
def parse_lifo_excel(file_bytes, original_filename=''):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    sheets = wb.sheetnames
    if len(sheets) < 2:
        raise ValueError(f"시트가 2개 이상이어야 합니다. (현재: {len(sheets)}개)")

    # 시트1에서 응답 규칙 검증
    ws1 = wb[sheets[0]]
    response_errors = validate_responses(ws1)

    # 시트2에서 점수 추출
    ws = wb[sheets[1]]
    s_plus = ws['D12'].value or 0
    c_plus = ws['F12'].value or 0
    ch_plus = ws['H12'].value or 0
    a_plus = ws['J12'].value or 0
    g_minus = ws['D13'].value or 0
    t_minus = ws['F13'].value or 0
    h_minus = ws['H13'].value or 0
    d_minus = ws['J13'].value or 0

    name, dept = extract_name_from_filename(original_filename)

    top_type = ws['I5'].value or ''
    if not top_type or str(top_type).startswith('#') or str(top_type) == 'None':
        type_map = {'SG': s_plus, 'CT': c_plus, 'CH': ch_plus, 'AD': a_plus}
        top_type = max(type_map, key=type_map.get)

    wb.close()
    return {
        'name': name, 'dept': dept,
        'top_type': str(top_type).strip(),
        'plus': {'SG': int(s_plus), 'CT': int(c_plus), 'CH': int(ch_plus), 'AD': int(a_plus)},
        'minus': {'SG': int(g_minus), 'CT': int(t_minus), 'CH': int(h_minus), 'AD': int(d_minus)},
        'response_errors': response_errors,
    }


def parse_lifo_file(file_bytes, original_filename=''):
    """확장자에 따라 엑셀/PDF 파서로 분기합니다."""
    if original_filename.lower().endswith('.pdf'):
        return parse_lifo_pdf(file_bytes, original_filename)
    return parse_lifo_excel(file_bytes, original_filename)


# ── PDF 파싱 ────────────────────────────────────────────────
def parse_lifo_pdf(file_bytes, original_filename=''):
    """PDF 진단 결과지에서 LIFO 점수를 추출합니다.

    엑셀과 동일한 데이터를 담은 PDF 리포트를 지원합니다.
    - "TOTAL +" / "TOTAL -" 행에서 4개 유형 점수를 읽습니다.
    - "문항N V" 응답으로 규칙(4,3,2,1 배정)을 검증합니다.
    """
    reader = PdfReader(io.BytesIO(file_bytes))
    text = "\n".join((page.extract_text() or "") for page in reader.pages)

    plus = _parse_total_line(text, '+')
    minus = _parse_total_line(text, '-')
    if plus is None:
        plus = _parse_component_scores(text, upper=True)
    if minus is None:
        minus = _parse_component_scores(text, upper=False)
    if plus is None or minus is None:
        raise ValueError('PDF에서 LIFO 점수를 찾을 수 없습니다. (TOTAL +/- 행 확인)')

    name, dept = extract_name_from_filename(original_filename)

    m = re.search(r'(SG|CT|CH|AD)\s*나의\s*유형|나의\s*유형\s*(SG|CT|CH|AD)', text)
    top_type = (m.group(1) or m.group(2)) if m else max(plus, key=plus.get)

    return {
        'name': name, 'dept': dept,
        'top_type': str(top_type).strip(),
        'plus': {k: int(plus[k]) for k in LIFO_KEYS},
        'minus': {k: int(minus[k]) for k in LIFO_KEYS},
        'response_errors': validate_pdf_responses(text),
    }


def _parse_total_line(text, sign):
    """'TOTAL +' 또는 'TOTAL -' 행에서 4개 유형 점수를 추출합니다."""
    marker = 'TOTAL +' if sign == '+' else 'TOTAL -'
    for line in text.splitlines():
        if marker in line:
            nums = [int(n) for _, n in re.findall(r'([A-Za-z])\s+(\d+)', line)]
            if len(nums) >= 4:
                return dict(zip(LIFO_KEYS, nums[:4]))
    return None


def _parse_component_scores(text, upper=True):
    """세부 항목 점수(A~L 또는 a~l)를 합산하여 4개 유형 점수를 계산합니다."""
    letters = 'ABCDEFGHIJKL' if upper else 'abcdefghijkl'
    scores = {}
    for line in text.splitlines():
        for letter, val in re.findall(r'(?:^|\s)([A-La-l])\s+(\d+)(?:\s|$)', line):
            if letter in letters and letter not in scores:
                scores[letter] = int(val)
    if len(scores) < 12:
        return None
    groups = {
        'SG': letters[0::4][:3], 'CT': letters[1::4][:3],
        'CH': letters[2::4][:3], 'AD': letters[3::4][:3],
    }
    return {k: sum(scores[c] for c in cols) for k, cols in groups.items()}


def validate_pdf_responses(text):
    """PDF의 '문항N V' 응답이 규칙(각 4문항에 4,3,2,1 중복 없이 배정)을 지켰는지 검증합니다."""
    resp = {int(q): int(v) for q, v in re.findall(r'문항(\d+)\s+(\d)', text)}
    errors = []
    qs = sorted(resp)
    for i in range(0, len(qs), 4):
        grp = qs[i:i + 4]
        if len(grp) == 4:
            scores = [resp[q] for q in grp]
            if sorted(scores) != [1, 2, 3, 4]:
                errors.append({'questions': f'{grp[0]}-{grp[-1]}', 'scores': scores})
    return errors


def validate_responses(ws):
    """시트1의 응답이 규칙(각 4문항 그룹에서 4,3,2,1 중복 없이 배정)을 지켰는지 검증합니다."""
    # E열에서 C열 문항번호를 기준으로 4개씩 그룹화
    groups = []
    current_group = []
    for r in range(5, 128):
        c_val = ws[f'C{r}'].value
        e_val = ws[f'E{r}'].value
        if c_val is not None and e_val is not None:
            try:
                current_group.append((int(c_val), int(e_val)))
            except (ValueError, TypeError):
                continue
            if len(current_group) == 4:
                groups.append(current_group)
                current_group = []

    errors = []
    for g in groups:
        scores = [item[1] for item in g]
        if sorted(scores) != [1, 2, 3, 4]:
            q_start, q_end = g[0][0], g[3][0]
            errors.append({'questions': f'{q_start}-{q_end}', 'scores': scores})

    return errors


def strip_title(name):
    """이름에서 직급 키워드를 제거합니다. '정청산 책임' → '정청산'"""
    result = name.strip()
    for title in TITLE_KEYWORDS:
        result = re.sub(rf'\s*{title}\s*$', '', result)
    return result.strip()


# 파일명에서 제거할 라벨 토큰 (LIFO 리포트 관련 단어 / 날짜)
FILENAME_LABEL_RE = re.compile(
    r'^(LIFO|진단지|진단|결과지|결과|리포트|레포트|보고서|report|\d{4,})',
    re.IGNORECASE,
)


def extract_name_from_filename(filename):
    """파일명에서 이름과 부서를 추출합니다.

    지원 패턴:
      - ...진단지_부서_이름              (엑셀 진단지)
      - 이름_LIFO진단, 부서_이름_LIFO진단  (PDF 리포트)
    """
    base = os.path.splitext(filename)[0]
    match = re.search(r'진단지[_\s]+(.+?)_([^_]+?)\s*$', base)
    if match:
        return strip_title(match.group(2).strip()), match.group(1).strip()
    match2 = re.search(r'진단지[_\s]+(.+?)\s+(\S+)\s*$', base)
    if match2:
        return strip_title(match2.group(2).strip()), match2.group(1).strip()

    # 'LIFO진단', '결과', 날짜 등 라벨 필드를 제거하고 남는 것을 부서/이름으로 사용
    # (필드 구분은 '_'만 사용해 '이 웅'처럼 공백이 들어간 이름을 보존)
    tokens = [t.strip() for t in base.split('_') if t.strip()]
    kept = [t for t in tokens if not FILENAME_LABEL_RE.search(t)]
    if len(kept) >= 2:
        return strip_title(kept[-1].strip()), kept[-2].strip()
    if len(kept) == 1:
        return strip_title(kept[0].strip()), ''

    parts = re.split(r'[_]', base)
    if len(parts) >= 2:
        return strip_title(parts[-1].strip()), parts[-2].strip() if len(parts) >= 3 else ''
    return base, ''


# ── PPT 생성 ────────────────────────────────────────────────
def clone_slide(prs, template_slide):
    slide_layout = template_slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)
    spTree = new_slide.shapes._spTree
    for sp in list(spTree):
        tag = sp.tag.split('}')[-1] if '}' in sp.tag else sp.tag
        if tag in ('sp', 'cxnSp', 'graphicFrame', 'pic', 'grpSp'):
            spTree.remove(sp)
    for child in template_slide.shapes._spTree:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag in ('sp', 'cxnSp', 'graphicFrame', 'pic', 'grpSp'):
            spTree.append(deepcopy(child))
    return new_slide


def find_tables(slide):
    tables = [s for s in slide.shapes if s.has_table]
    tables.sort(key=lambda s: s.top)
    return tables


def update_cell_text_preserve_style(cell, new_text):
    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    runs = cell._tc.findall('.//a:r', ns)
    if runs:
        t_elem = runs[0].find('a:t', ns)
        if t_elem is not None:
            t_elem.text = str(new_text)
        for extra_run in runs[1:]:
            extra_run.getparent().remove(extra_run)
    else:
        cell.text = str(new_text)


def update_score_color(cell, is_red):
    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    rPr = cell._tc.find('.//a:rPr', ns)
    if rPr is not None:
        old_fill = rPr.find('a:solidFill', ns)
        if old_fill is not None:
            rPr.remove(old_fill)
        color_val = 'FF0000' if is_red else '000000'
        solidFill = etree.Element(f'{{{A}}}solidFill')
        srgbClr = etree.SubElement(solidFill, f'{{{A}}}srgbClr')
        srgbClr.set('val', color_val)
        rPr.insert(0, solidFill)


def check_red_condition(plus_score, minus_score):
    diff = abs(plus_score - minus_score)
    plus_red = plus_score >= 28 or (diff > 4 and plus_score > minus_score)
    minus_red = minus_score >= 28 or (diff > 4 and minus_score > plus_score)
    return plus_red, minus_red


def fill_table_data(table_shape, person_data):
    table = table_shape.table
    update_cell_text_preserve_style(table.cell(0, 0), person_data['name'])
    plus = person_data['plus']
    minus = person_data['minus']
    for i, key in enumerate(['SG', 'CT', 'CH', 'AD']):
        p_red, m_red = check_red_condition(plus[key], minus[key])
        cell_plus = table.cell(1, i + 1)
        update_cell_text_preserve_style(cell_plus, str(plus[key]))
        update_score_color(cell_plus, p_red)
        cell_minus = table.cell(2, i + 1)
        update_cell_text_preserve_style(cell_minus, str(minus[key]))
        update_score_color(cell_minus, m_red)


def clear_table(table_shape):
    table = table_shape.table
    update_cell_text_preserve_style(table.cell(0, 0), '')
    for ri in range(1, 3):
        for ci in range(1, 5):
            update_cell_text_preserve_style(table.cell(ri, ci), '')
            update_score_color(table.cell(ri, ci), False)


def delete_slide(prs, slide_index):
    rId = prs.slides._sldIdLst[slide_index].get(
        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    )
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[slide_index]


def generate_pptx(participants):
    prs = Presentation(TEMPLATE_PATH)
    template_slide = prs.slides[1]
    num_slides = (len(participants) + 1) // 2

    new_slides = [clone_slide(prs, template_slide) for _ in range(num_slides)]

    for slide_idx, slide in enumerate(new_slides):
        tables = find_tables(slide)
        if len(tables) < 2:
            continue
        p1_idx = slide_idx * 2
        if p1_idx < len(participants):
            fill_table_data(tables[0], participants[p1_idx])
        p2_idx = slide_idx * 2 + 1
        if p2_idx < len(participants):
            fill_table_data(tables[1], participants[p2_idx])
        else:
            clear_table(tables[1])

    delete_slide(prs, 1)
    delete_slide(prs, 0)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ── Streamlit UI ─────────────────────────────────────────────
st.set_page_config(page_title='LIFO 진단 → PPT', page_icon='📊', layout='centered')

st.markdown("""
<style>
    .block-container { max-width: 960px; padding-top: 2rem; }
    header[data-testid="stHeader"] { background: transparent; }
    .hero { text-align: center; padding: 1rem 0 0.5rem; }
    .hero h1 { font-size: 2rem; margin-bottom: 0.25rem; }
    .hero p { color: #888; font-size: 0.95rem; }
    .step-label {
        display: inline-block; background: #FF4B4B; color: white;
        border-radius: 50%; width: 26px; height: 26px; line-height: 26px;
        text-align: center; font-size: 13px; font-weight: 700;
        margin-right: 6px; vertical-align: middle;
    }
    .how-box {
        background: #f8f9fb; border-radius: 12px; padding: 1.2rem 1.5rem;
        border: 1px solid #e8eaed; margin: 0.5rem 0 1rem;
    }
    .how-box ol { margin: 0.5rem 0 0 1.2rem; padding: 0; }
    .how-box li { margin-bottom: 0.3rem; color: #444; font-size: 0.9rem; }
</style>
""", unsafe_allow_html=True)

# 헤더
st.markdown("""
<div class="hero">
    <h1>📊 LIFO 진단 → PPT 자동 생성</h1>
    <p>엑셀 또는 PDF 진단 파일을 올리면, 서식이 적용된 PPT 리포트를 자동으로 만들어 드립니다.</p>
</div>
""", unsafe_allow_html=True)

st.divider()

# ── 파일 업로드 (sidebar 또는 하단에 배치하기 위해 먼저 처리)
# 결과가 없을 때는 메인에, 결과가 있을 때는 하단에 표시
has_results = 'results' in st.session_state and st.session_state['results']

if not has_results:
    # 첫 화면: 사용법 + 업로드가 메인
    with st.expander('💡 사용 방법 보기', expanded=False):
        st.markdown("""
<div class="how-box">
<ol>
    <li><b>파일 업로드</b> — LIFO 행동강약점 진단 엑셀(.xlsx) 또는 PDF를 여러 개 한꺼번에 올려주세요.</li>
    <li><b>결과 확인</b> — 자동으로 이름, 부서, 점수가 추출됩니다. 이름이 잘못 나왔으면 수정할 수 있습니다.</li>
    <li><b>PPT 다운로드</b> — 버튼 하나로 서식이 적용된 PPT가 생성됩니다.</li>
</ol>
</div>
""", unsafe_allow_html=True)

    st.markdown('<span class="step-label">1</span> <b>엑셀 · PDF 파일 업로드</b>', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    '파일 선택',
    type=['xlsx', 'pdf'],
    accept_multiple_files=True,
    label_visibility='collapsed' if not has_results else 'visible',
    help='LIFO 행동강약점(행동유형) 진단지 엑셀(.xlsx) 또는 PDF 파일을 선택하세요. 여러 개를 한 번에 올릴 수 있습니다.',
    key='file_uploader',
)

if uploaded_files:
    # 파일 변경 감지
    file_key = '|'.join(sorted(f.name for f in uploaded_files))
    if st.session_state.get('_file_key') != file_key:
        results = []
        errors = []
        progress_placeholder = st.empty()
        progress = progress_placeholder.progress(0, text='분석 중...')
        for idx, f in enumerate(uploaded_files):
            try:
                data = parse_lifo_file(f.read(), f.name)
                f.seek(0)
                results.append(data)
            except Exception as e:
                errors.append(f'{f.name}: {e}')
            progress.progress((idx + 1) / len(uploaded_files), text=f'{idx+1}/{len(uploaded_files)} 분석 완료')
        progress_placeholder.empty()
        st.session_state['results'] = results
        st.session_state['errors'] = errors
        st.session_state['_file_key'] = file_key
        st.session_state['ppt_buf'] = None  # 새 파일이면 기존 PPT 초기화
        st.rerun()

    results = st.session_state.get('results', [])
    errors = st.session_state.get('errors', [])

    if errors:
        for err in errors:
            st.error(err)

    if results:
        # ── 결과 헤더 + PPT 다운로드 버튼을 같은 줄에 ──
        col_title, col_btn = st.columns([3, 2])
        with col_title:
            st.markdown(
                f'<span class="step-label">✓</span> <b>결과 확인</b> '
                f'<span style="color:#888; font-size:0.85rem;">({len(results)}명 / {(len(results)+1)//2}슬라이드)</span>',
                unsafe_allow_html=True,
            )
        with col_btn:
            # PPT 생성이 안 됐으면 생성 버튼, 됐으면 다운로드 버튼
            if st.session_state.get('ppt_buf') is None:
                if st.button('📥 PPT 생성', type='primary', use_container_width=True):
                    with st.spinner('PPT 생성 중...'):
                        buf = generate_pptx(results)
                        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                        st.session_state['ppt_buf'] = buf
                        st.session_state['ppt_filename'] = f'LIFO_Report_{timestamp}.pptx'
                    st.rerun()
            else:
                st.download_button(
                    label=f'💾 PPT 다운로드',
                    data=st.session_state['ppt_buf'],
                    file_name=st.session_state['ppt_filename'],
                    mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                    use_container_width=True,
                    type='primary',
                )

        # 정렬
        sort_col = st.radio(
            '정렬 기준', ['업로드 순', '이름순', '부서순', '유형순'],
            horizontal=True, label_visibility='collapsed',
        )
        sorted_results = list(results)
        if sort_col == '이름순':
            sorted_results = sorted(results, key=lambda x: x['name'])
        elif sort_col == '부서순':
            sorted_results = sorted(results, key=lambda x: x['dept'])
        elif sort_col == '유형순':
            sorted_results = sorted(results, key=lambda x: x['top_type'])

        # 마크다운 테이블
        error_count = sum(1 for r in sorted_results if r.get('response_errors'))
        if error_count:
            st.warning(f'응답 규칙 오류가 {error_count}명에게서 발견되었습니다. (각 문항 그룹에서 4,3,2,1을 중복 없이 배정해야 합니다)')

        header = '| # | 이름 | 부서 | 유형 | +SG | +CT | +CH | +AD | -SG | -CT | -CH | -AD | 상태 |\n'
        header += '|--:|:--|:--|:--:|--:|--:|--:|--:|--:|--:|--:|--:|:--:|\n'
        rows_md = ''
        for i, r in enumerate(sorted_results):
            p, m = r['plus'], r['minus']
            row = [f"{i+1}", r['name'], r['dept'], f"`{r['top_type']}`"]
            for key in ['SG', 'CT', 'CH', 'AD']:
                p_red, _ = check_red_condition(p[key], m[key])
                row.append(f"**:red[{p[key]}]**" if p_red else str(p[key]))
            for key in ['SG', 'CT', 'CH', 'AD']:
                _, m_red = check_red_condition(p[key], m[key])
                row.append(f"**:red[{m[key]}]**" if m_red else str(m[key]))
            # 응답 오류 표시
            resp_errs = r.get('response_errors', [])
            if resp_errs:
                row.append(f'**:orange[오류 {len(resp_errs)}건]**')
            else:
                row.append('')
            rows_md += '| ' + ' | '.join(row) + ' |\n'

        st.markdown(header + rows_md)

        # 오류 상세 내역
        if error_count:
            error_persons = [r for r in sorted_results if r.get('response_errors')]
            with st.expander(f'⚠️ 응답 오류 ({error_count}명)'):
                # 명단 / 상세 탭
                tab_list, tab_detail = st.tabs(['📋 오류 명단', '🔍 오류 상세'])
                with tab_list:
                    for r in error_persons:
                        errs = r['response_errors']
                        st.markdown(f"- **{r['name']}** ({r['dept']}) — {len(errs)}건")
                with tab_detail:
                    for r in error_persons:
                        errs = r['response_errors']
                        with st.expander(f"**{r['name']}** ({r['dept']}) — {len(errs)}건"):
                            for e in errs:
                                st.caption(f"문항 {e['questions']}: 응답 {e['scores']}")

        # 이름/부서 수정
        with st.expander('✏️ 이름이 잘못 나왔나요? 클릭해서 수정하세요'):
            cols = st.columns(4)
            for i, r in enumerate(results):
                with cols[i % 4]:
                    new_name = st.text_input(
                        '이름', value=r['name'], key=f'name_{i}',
                        label_visibility='collapsed',
                        placeholder=f'{r["name"]}',
                    )
                    if new_name != r['name']:
                        r['name'] = new_name
                        st.session_state['ppt_buf'] = None  # 이름 바꾸면 PPT 재생성 필요

        st.divider()

        # ── 파일 추가/변경 (하단)
        st.markdown(
            '<span style="color:#888; font-size:0.85rem;">📂 파일을 추가하거나 변경하려면 위의 업로더를 사용하세요.</span>',
            unsafe_allow_html=True,
        )

elif not has_results:
    st.markdown("""
    <div style="text-align:center; padding: 3rem 1rem; color: #aaa;">
        <p style="font-size: 3rem; margin-bottom: 0.5rem;">📂</p>
        <p>위의 <b>Browse files</b> 버튼을 눌러<br>LIFO 진단 엑셀 또는 PDF 파일을 업로드하세요.</p>
    </div>
    """, unsafe_allow_html=True)
