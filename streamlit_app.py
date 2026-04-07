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
from pptx import Presentation
from pptx.util import Pt, Emu
from lxml import etree

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.pptx')

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
    }


def strip_title(name):
    """이름에서 직급 키워드를 제거합니다. '정청산 책임' → '정청산'"""
    result = name.strip()
    for title in TITLE_KEYWORDS:
        result = re.sub(rf'\s*{title}\s*$', '', result)
    return result.strip()


def extract_name_from_filename(filename):
    base = os.path.splitext(filename)[0]
    match = re.search(r'진단지[_\s]+(.+?)_([^_]+?)\s*$', base)
    if match:
        return strip_title(match.group(2).strip()), match.group(1).strip()
    match2 = re.search(r'진단지[_\s]+(.+?)\s+(\S+)\s*$', base)
    if match2:
        return strip_title(match2.group(2).strip()), match2.group(1).strip()
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
    <p>엑셀 진단 파일을 올리면, 서식이 적용된 PPT 리포트를 자동으로 만들어 드립니다.</p>
</div>
""", unsafe_allow_html=True)

st.divider()

# 사용법 안내
with st.expander('💡 사용 방법 보기', expanded=False):
    st.markdown("""
<div class="how-box">
<ol>
    <li><b>엑셀 파일 업로드</b> — LIFO 행동강약점 진단 엑셀(.xlsx)을 여러 개 한꺼번에 올려주세요.</li>
    <li><b>결과 확인</b> — 자동으로 이름, 부서, 점수가 추출됩니다. 이름이 잘못 나왔으면 수정할 수 있습니다.</li>
    <li><b>PPT 다운로드</b> — 버튼 하나로 서식이 적용된 PPT가 생성됩니다.</li>
</ol>
</div>
""", unsafe_allow_html=True)

# ── STEP 1: 파일 업로드
st.markdown('<span class="step-label">1</span> <b>엑셀 파일 업로드</b>', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    '파일 선택',
    type=['xlsx'],
    accept_multiple_files=True,
    label_visibility='collapsed',
    help='LIFO 행동강약점(행동유형) 진단지 엑셀 파일을 선택하세요. 여러 개를 한 번에 올릴 수 있습니다.',
)

if uploaded_files:
    # 파일 해시로 변경 감지 (파일 개수 + 파일명 조합)
    file_key = '|'.join(sorted(f.name for f in uploaded_files))
    if st.session_state.get('_file_key') != file_key:
        results = []
        errors = []
        progress = st.progress(0, text='분석 중...')
        for idx, f in enumerate(uploaded_files):
            try:
                data = parse_lifo_excel(f.read(), f.name)
                f.seek(0)
                results.append(data)
            except Exception as e:
                errors.append(f'{f.name}: {e}')
            progress.progress((idx + 1) / len(uploaded_files), text=f'{idx+1}/{len(uploaded_files)} 분석 완료')
        progress.empty()
        st.session_state['results'] = results
        st.session_state['errors'] = errors
        st.session_state['_file_key'] = file_key

    results = st.session_state['results']
    errors = st.session_state['errors']

    if errors:
        for err in errors:
            st.error(err)

    if results:
        st.success(f'{len(results)}명의 진단 데이터를 추출했습니다.')

        st.divider()

        # ── STEP 2: 결과 확인
        st.markdown('<span class="step-label">2</span> <b>결과 확인</b>', unsafe_allow_html=True)

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
        header = '| # | 이름 | 부서 | 유형 | +SG | +CT | +CH | +AD | -SG | -CT | -CH | -AD |\n'
        header += '|--:|:--|:--|:--:|--:|--:|--:|--:|--:|--:|--:|--:|\n'
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
            rows_md += '| ' + ' | '.join(row) + ' |\n'

        st.markdown(header + rows_md)

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

        st.divider()

        # ── STEP 3: PPT 생성
        st.markdown('<span class="step-label">3</span> <b>PPT 생성</b>', unsafe_allow_html=True)

        slide_count = (len(results) + 1) // 2
        st.caption(f'{len(results)}명 → {slide_count}장 슬라이드 (슬라이드당 2명)')

        if st.button('📥 PPT 생성 및 다운로드', type='primary', use_container_width=True):
            with st.spinner('PPT를 만들고 있습니다...'):
                buf = generate_pptx(results)
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f'LIFO_Report_{timestamp}.pptx'

            st.balloons()
            st.download_button(
                label=f'💾 {filename} 저장',
                data=buf,
                file_name=filename,
                mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                use_container_width=True,
            )
else:
    st.markdown("""
    <div style="text-align:center; padding: 3rem 1rem; color: #aaa;">
        <p style="font-size: 3rem; margin-bottom: 0.5rem;">📂</p>
        <p>위의 <b>Browse files</b> 버튼을 눌러<br>LIFO 진단 엑셀 파일을 업로드하세요.</p>
    </div>
    """, unsafe_allow_html=True)
