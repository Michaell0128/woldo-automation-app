import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook

st.set_page_config(page_title="월도자동화시스템", layout="wide")

with st.container():
    st.markdown("""
    <h1 style='text-align:center; color:#4A90E2;'>📦 <span style='font-weight:500'>월도자동화시스템</span></h1>
    <p style='text-align:center; font-size:16px; color:gray;'>월도 발주서 및 네이버 송장 엑셀 자동 생성 솔루션</p>
    """, unsafe_allow_html=True)

st.markdown("---")

def extract_keywords(text):
    return re.sub(r'[^\w\s]', '', str(text)).lower().split()

def match_product_candidates(a_row, b_df):
    option_info = str(a_row.get("옵션정보", ""))
    if ':' in option_info:
        product_part, option_part = map(str.strip, option_info.split(':', 1))
    else:
        product_part, option_part = option_info.strip(), ""

    product_keywords = extract_keywords(product_part)
    option_keywords = extract_keywords(option_part)

    candidates = []
    for i, b_row in b_df.iterrows():
        b_name_keywords = extract_keywords(b_row.get("상품명", ""))
        b_option_keywords = extract_keywords(b_row.get("옵션명", ""))

        product_matches = sum(1 for kw in product_keywords if kw in b_name_keywords)
        option_matches = sum(1 for kw in option_keywords if kw in b_option_keywords)
        total_score = product_matches + option_matches

        if product_matches > 0 and option_matches > 0:
            candidates.append((total_score, i, b_row))

    candidates.sort(reverse=True)
    return candidates

# ▒▒ TAB UI 구성 ▒▒
tabs = st.tabs(["🛒 월도 발주서 생성", "📦 네이버 송장 엑셀 생성"])

# ▒▒ TAB 1: 발주 생성 ▒▒
with tabs[0]:
    st.markdown("""
    <h3 style='color:#4A90E2;'>🛒 네이버 주문서 + 월도 상품목록 → <strong>C 발주서</strong></h3>
    """, unsafe_allow_html=True)
    with st.expander("📁 파일 업로드 및 정보 입력", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            a_file = st.file_uploader("네이버 주문서", type=["xlsx"])
        with col2:
            b_file = st.file_uploader("월도 상품목록", type=["xlsx"])

        sender_name = st.text_input("송하인 이름", value="전국농가자랑")
        sender_phone = st.text_input("송하인 연락처", value="010-2890-0086")

        submitted = st.button("🚀 월도 발주서 생성하기")

    if submitted and a_file and b_file:
        a_df = pd.read_excel(a_file)
        b_df = pd.read_excel(b_file)

        c_rows = []
        match_count = 0

        for _, a_row in a_df.iterrows():
            candidates = match_product_candidates(a_row, b_df)
            if len(candidates) == 1:
                match = candidates[0][2]
            elif len(candidates) > 1:
                st.warning(f"⚠️ 복수 후보가 발견되었습니다: {a_row['옵션정보']}")
                option_map = {
                    f"{c[2]['상품명']} / {c[2]['옵션명']} (점수:{c[0]})": c[2] for c in candidates
                }
                selected = st.selectbox("⛳ 아래 중 어떤 상품과 매칭할까요?", list(option_map.keys()), key=a_row['옵션정보'])
                match = option_map[selected]
            else:
                match = None
            if match is not None:
                c_rows.append({
                    '순서': match['순서'],
                    '상품번호': match['상품번호'],
                    '상품명': match['상품명'],
                    '옵션번호': match['옵션번호'],
                    '옵션명': match['옵션명'],
                    '배송비조건': match['배송비조건'],
                    '판매가격': match['판매가격'],
                    '수량': a_row.get('수량', 1),
                    '주문자 성명': sender_name,
                    '주문자 전화번호': sender_phone,
                    '수취인 성명': a_row.get('수취인명', ''),
                    '수취인 전화번호': a_row.get('수취인연락처1', ''),
                    '수취인 주소': a_row.get('통합배송지', ''),
                    '배송메시지': a_row.get('배송메세지', ''),
                    '판매사 주문번호': '',
                    '판매사 옵션번호': ''
                })
                match_count += 1

        st.success(f"🎉 총 {match_count}건의 상품이 매칭되었습니다.")
        st.markdown("#### 📊 생성된 발주서 미리보기")
        st.dataframe(pd.DataFrame(c_rows).head(), use_container_width=True)

        c_buffer = BytesIO()
        with pd.ExcelWriter(c_buffer, engine='openpyxl') as writer:
            pd.DataFrame(c_rows).to_excel(writer, index=False)
        c_buffer.seek(0)

        st.download_button(
            label="📥 월도 발주서 다운로드",
            data=c_buffer,
            file_name="월도발주서.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ▒▒ TAB 2: 송장 생성 ▒▒
with tabs[1]:
    st.markdown("""
    <h3 style='color:#50AF61;'>📦 네이버 주문서 + 월도 송장서 → <strong>E 네이버 송장 엑셀</strong></h3>
    """, unsafe_allow_html=True)
    with st.expander("📁 파일 업로드", expanded=True):
        col3, col4 = st.columns(2)
        with col3:
            a_file2 = st.file_uploader("네이버 주문서", type=["xlsx"], key="a2")
        with col4:
            d_file = st.file_uploader("월도 송장서", type=["xlsx"], key="d")

        invoice_submitted = st.button("🚀 네이버 송장 엑셀")

    if invoice_submitted and a_file2 and d_file:
        a_df2 = pd.read_excel(a_file2)
        d_df = pd.read_excel(d_file)

        e_rows = []
        for _, d_row in d_df.iterrows():
            d_product = str(d_row['상품명']) + ' ' + str(d_row['옵션명'])
            d_keywords = set(extract_keywords(d_product))

            best_match = None
            max_score = 0

            for _, a_row in a_df2.iterrows():
                a_option_info = str(a_row.get('옵션정보', ''))
                a_keywords = set(extract_keywords(a_option_info))
                match_score = len(d_keywords & a_keywords)
                if match_score > max_score:
                    max_score = match_score
                    best_match = a_row

            if best_match is not None:
                e_rows.append({
                    '상품주문번호': str(best_match['상품주문번호']),
                    '배송방법': '택배,등기,소포',
                    '택배사': str(d_row.get('판매사 주문번호', '')),
                    '송장번호': str(d_row.get('판매사 옵션번호', ''))
                })

        st.success(f"📦 총 {len(e_rows)}건의 송장번호가 정상 매칭되었습니다.")
        st.markdown("#### 📊 생성된 네이버 송장 엑셀 미리보기")
        st.dataframe(pd.DataFrame(e_rows).head(), use_container_width=True)

        e_buffer = BytesIO()
        with pd.ExcelWriter(e_buffer, engine='openpyxl') as writer:
            pd.DataFrame(e_rows).to_excel(writer, index=False)
        e_buffer.seek(0)

        st.download_button(
            label="📥 E문서 다운로드 (네이버 송장엑셀 - .xls 형식)",
            data=e_buffer,
            file_name="네이버송장엑셀.xls",
            mime="application/vnd.ms-excel"
        )
