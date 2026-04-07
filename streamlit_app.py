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


def extract_name_from_filename(filename):
    base = os.path.splitext(filename)[0]
    match = re.search(r'진단지[_\s]+(.+?)_([^_]+?)\s*$', base)
    if match:
        return match.group(2).strip(), match.group(1).strip()
    match2 = re.search(r'진단지[_\s]+(.+?)\s+(\S+)\s*$', base)
    if match2:
        return match2.group(2).strip(), match2.group(1).strip()
    parts = re.split(r'[_]', base)
    if len(parts) >= 2:
        return parts[-1].strip(), parts[-2].strip() if len(parts) >= 3 else ''
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
st.set_page_config(page_title='LIFO → PPT', page_icon='📊', layout='wide')

st.markdown("""
<style>
    .block-container { max-width: 1000px; }
    .stDataFrame { font-size: 13px; }
    div[data-testid="stFileUploader"] { margin-bottom: 0; }
    .red-score { color: #FF0000; font-weight: bold; }
    .black-score { color: #000000; }
</style>
""", unsafe_allow_html=True)

st.title('📊 LIFO 진단 → PPT 자동 생성')
st.caption('CLiCK IAM MICRO _ LIFO Analysis Report Generator')
st.divider()

# ── STEP 1: 파일 업로드
st.subheader('① 엑셀 파일 업로드')
uploaded_files = st.file_uploader(
    'LIFO 진단 엑셀 파일을 선택하세요 (.xlsx)',
    type=['xlsx'],
    accept_multiple_files=True,
)

if uploaded_files:
    # 파싱
    if 'results' not in st.session_state or st.session_state.get('_file_count') != len(uploaded_files):
        results = []
        errors = []
        for f in uploaded_files:
            try:
                data = parse_lifo_excel(f.read(), f.name)
                f.seek(0)
                results.append(data)
            except Exception as e:
                errors.append(f'{f.name}: {e}')
        st.session_state['results'] = results
        st.session_state['errors'] = errors
        st.session_state['_file_count'] = len(uploaded_files)

    results = st.session_state['results']
    errors = st.session_state['errors']

    if errors:
        for err in errors:
            st.error(err)

    if results:
        st.divider()
        st.subheader(f'② 진단 결과 확인 ({len(results)}명)')

        # 정렬
        sort_col = st.radio('정렬', ['업로드 순', '이름순', '부서순', '유형순'], horizontal=True)
        if sort_col == '이름순':
            results = sorted(results, key=lambda x: x['name'])
        elif sort_col == '부서순':
            results = sorted(results, key=lambda x: x['dept'])
        elif sort_col == '유형순':
            results = sorted(results, key=lambda x: x['top_type'])

        # 테이블 표시
        def fmt_score(plus_val, minus_val):
            p_red, m_red = check_red_condition(plus_val, minus_val)
            p_style = 'red-score' if p_red else 'black-score'
            m_style = 'red-score' if m_red else 'black-score'
            return (
                f'<span class="{p_style}">{plus_val}</span>',
                f'<span class="{m_style}">{minus_val}</span>',
            )

        header = '| No. | 이름 | 부서 | 1순위 | +SG | +CT | +CH | +AD | -SG | -CT | -CH | -AD |\n'
        header += '|--:|:--|:--|:--:|--:|--:|--:|--:|--:|--:|--:|--:|\n'
        rows_md = ''
        for i, r in enumerate(results):
            p = r['plus']
            m = r['minus']
            row_parts = [f"{i+1}", r['name'], r['dept'], f"`{r['top_type']}`"]
            for key in ['SG', 'CT', 'CH', 'AD']:
                p_red, _ = check_red_condition(p[key], m[key])
                row_parts.append(f"**:red[{p[key]}]**" if p_red else str(p[key]))
            for key in ['SG', 'CT', 'CH', 'AD']:
                _, m_red = check_red_condition(p[key], m[key])
                row_parts.append(f"**:red[{m[key]}]**" if m_red else str(m[key]))
            rows_md += '| ' + ' | '.join(row_parts) + ' |\n'

        st.markdown(header + rows_md, unsafe_allow_html=True)

        # 이름 수정
        with st.expander('이름/부서 수정'):
            edited = False
            cols = st.columns(3)
            for i, r in enumerate(results):
                col = cols[i % 3]
                with col:
                    new_name = st.text_input(f'{r["name"]}', value=r['name'], key=f'name_{i}')
                    new_dept = st.text_input(f'{r["name"]} 부서', value=r['dept'], key=f'dept_{i}')
                    if new_name != r['name'] or new_dept != r['dept']:
                        r['name'] = new_name
                        r['dept'] = new_dept
                        edited = True

        st.divider()
        st.subheader('③ PPT 생성')
        slide_count = (len(results) + 1) // 2
        st.info(f'{len(results)}명 → {slide_count}장 슬라이드 (슬라이드당 2명)')

        if st.button('🎯 PPT 생성 및 다운로드', type='primary', use_container_width=True):
            with st.spinner('PPT 생성 중...'):
                buf = generate_pptx(results)
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f'LIFO_Report_{timestamp}.pptx'

            st.success(f'{len(results)}명의 리포트가 생성되었습니다!')
            st.download_button(
                label='📥 PPT 다운로드',
                data=buf,
                file_name=filename,
                mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                use_container_width=True,
            )
else:
    st.info('엑셀 파일을 업로드하면 자동으로 분석됩니다.')
