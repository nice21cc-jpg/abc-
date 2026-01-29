import streamlit as st
import pandas as pd
import plotly.express as px
import os
import re
import glob

# --------------------------------------------------------------------------
# 1. 페이지 설정 및 디자인
# --------------------------------------------------------------------------
st.set_page_config(layout="wide", page_title="진로-진학 나침반", page_icon="🧭")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700;900&display=swap');
    html, body, [class*="css"] { font-family: 'Noto Sans KR', sans-serif; color: #333; }

    /* 메인 타이틀 */
    .main-title { font-size: 38px; font-weight: 900; color: #1e3a8a; margin-bottom: 5px; border-bottom: 4px solid #1e3a8a; padding-bottom: 10px; }
    
    /* 자료 출처 스타일 */
    .source-text { font-size: 12px; color: #64748b; text-align: right; margin-bottom: 30px; letter-spacing: -0.5px; }

    /* 학과명 타이틀 */
    .dept-title { 
        font-size: 34px; font-weight: 800; color: #111827; 
        background-color: #f8fafc; padding: 15px; border-radius: 12px; 
        text-align: center; border: 2px solid #e2e8f0; margin-bottom: 25px;
    }
    
    /* 섹션 헤더 */
    .section-title { 
        font-size: 24px; font-weight: 700; color: #374151; 
        margin-top: 45px; margin-bottom: 15px; 
        border-left: 6px solid #2563eb; padding-left: 15px; 
    }

    /* 설명 박스 */
    .desc-box { 
        background-color: #f0f9ff; border: 2px solid #bae6fd; border-radius: 12px; padding: 25px; 
        font-size: 1.15em; line-height: 1.8; color: #0c4a6e; 
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }

    /* 과목 카드 */
    .subj-box { 
        border-radius: 16px; padding: 20px; text-align: center; height: 100%; 
        box-shadow: 0 4px 6px rgba(0,0,0,0.05); transition: transform 0.2s; 
        border-width: 2px; border-style: solid; display: flex; flex-direction: column; justify-content: start;
    }
    .subj-box:hover { transform: translateY(-5px); box-shadow: 0 10px 15px rgba(0,0,0,0.1); }
    
    .theme-blue   { background-color: #eff6ff; border-color: #bfdbfe; color: #1e40af; }
    .theme-orange { background-color: #fff7ed; border-color: #fed7aa; color: #9a3412; }
    .theme-purple { background-color: #faf5ff; border-color: #e9d5ff; color: #6b21a8; }

    .subj-header { display: block; font-weight: 800; font-size: 1.3em; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 1px dashed rgba(0,0,0,0.1); }
    .subj-content { font-size: 1.05em; font-weight: 500; word-break: keep-all; line-height: 1.6; }

    /* 탐구 주제 박스 */
    .inq-box { 
        background-color: #f0fdf4; border: 1px solid #86efac; border-left: 5px solid #16a34a;
        border-radius: 8px; padding: 15px 20px; margin-bottom: 10px; font-size: 1.05em; color: #14532d; 
    }
    .subject-tag {
        font-size: 0.85em; color: #15803d; border: 1px solid #86efac; padding: 3px 10px;
        border-radius: 15px; margin-right: 8px; background-color: #dcfce7; font-weight: 800;
    }
    .type-tag {
        font-size: 0.85em; color: #b45309; border: 1px solid #fdba74; padding: 3px 10px;
        border-radius: 15px; margin-right: 15px; background-color: #ffedd5; font-weight: 800;
    }

    /* KPI 지표 */
    .kpi-container { background-color: white; border: 1px solid #e2e8f0; border-radius: 12px; padding: 20px; text-align: center; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
    .kpi-label { font-size: 1em; color: #64748b; font-weight: 700; margin-bottom: 8px; }
    .kpi-value { font-size: 2em; font-weight: 900; color: #2563eb; }
</style>
""", unsafe_allow_html=True)

# --------------------------------------------------------------------------
# 2. 유틸리티 함수
# --------------------------------------------------------------------------
def pick_file(folder, keywords):
    """폴더에서 키워드가 포함된 파일 찾기"""
    if not os.path.isdir(folder): return None
    candidates = []
    # 대소문자 구분 없이 검색
    for file in os.listdir(folder):
        if file.endswith(".xlsx") or file.endswith(".xlsm"):
            if any(k in file for k in keywords):
                candidates.append(os.path.join(folder, file))
    
    candidates = sorted(list(set(candidates)), key=lambda x: len(os.path.basename(x)))
    return candidates[0] if candidates else None

def normalize(text):
    if pd.isna(text): return ""
    return str(text).replace(" ", "").replace("학과", "").replace("전공", "").replace("학부", "").replace("계열", "").strip()

def is_match(target, source):
    if pd.isna(source): return False
    t, s = normalize(target), normalize(source)
    if not t or not s: return False
    return (t in s) or (s in t)

def get_col_val(row, idx):
    return row.iloc[idx] if len(row) > idx else "-"

# --------------------------------------------------------------------------
# 3. 데이터 로드 (경로 자동 감지)
# --------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_all_data():
    # 현재 파일이 있는 폴더 위치 자동 감지
    data_dir = os.path.dirname(os.path.abspath(__file__))
    
    df_info, df_book, df_inq, df_susi = pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # 1. 학과 정보 파일 찾기
    f_info = pick_file(data_dir, ["계열 학과", "학과 계열"])
    if not f_info: f_info = pick_file(data_dir, ["학과카드"])
    
    if f_info:
        try: df_info = pd.read_excel(f_info, sheet_name=0, header=0)
        except: pass

    # 2. 추천 도서 파일 찾기
    f_book = pick_file(data_dir, ["추천도서"])
    if not f_book: f_book = pick_file(data_dir, ["학과카드"])

    if f_book:
        try:
            if "학과카드" in f_book:
                df_book = pd.read_excel(f_book, sheet_name=1, header=0).fillna('')
            else:
                df_book = pd.read_excel(f_book, header=0).fillna('')
            
            # 연도 전처리
            year_col = next((c for c in df_book.columns if "연도" in str(c)), None)
            if year_col:
                df_book[year_col] = pd.to_numeric(df_book[year_col], errors='coerce').fillna(0).astype(int).astype(str)
                df_book[year_col] = df_book[year_col].replace('0', '')
        except: pass

    # 3. 탐구 주제 파일 찾기
    f_inq_file = pick_file(data_dir, ["탐구주제", "탐구"])
    if f_inq_file:
        try: df_inq = pd.read_excel(f_inq_file).fillna('')
        except: pass

    # 4. 수시 데이터 파일 찾기
    f_susi_file = pick_file(data_dir, ["susi", "수시"])
    if f_susi_file:
        try:
            df_susi = pd.read_excel(f_susi_file, sheet_name="대학자료")
            cols_int = ['연도', '모집인원', '충원인원', '추합', '예비']
            for c in cols_int:
                if c in df_susi.columns: df_susi[c] = pd.to_numeric(df_susi[c], errors='coerce').fillna(0).astype(int)
            cols_float = ['경쟁률', '실경쟁률', '실질경쟁률', '등급50', '등급70']
            for c in cols_float:
                if c in df_susi.columns: df_susi[c] = pd.to_numeric(df_susi[c], errors='coerce')
            for c in ['지역', '대학', '전형', '학과']:
                if c in df_susi.columns: df_susi[c] = df_susi[c].astype(str)
        except: pass

    return df_info, df_book, df_inq, df_susi

# --------------------------------------------------------------------------
# 4. 실행 및 사이드바
# --------------------------------------------------------------------------
df_info, df_book, df_inq, df_susi = load_all_data()

st.sidebar.title("🔍 검색 옵션")

# [파트 1] 학과 특색
st.sidebar.header("1. 학과 특색 검색")
sel_cat = "전체"
sel_dept = "선택안함"

if not df_info.empty and len(df_info.columns) >= 2:
    cat_list = ["전체"] + sorted(df_info.iloc[:, 0].dropna().astype(str).unique().tolist())
    sel_cat = st.sidebar.selectbox("📂 계열 선택", cat_list)

    if sel_cat != "전체":
        filtered = df_info[df_info.iloc[:, 0].astype(str) == sel_cat]
        dept_list = sorted(filtered.iloc[:, 1].dropna().astype(str).unique().tolist())
    else:
        dept_list = sorted(df_info.iloc[:, 1].dropna().astype(str).unique().tolist())
    
    sel_dept = st.sidebar.selectbox("🎓 학과 선택", ["선택안함"] + dept_list)
else:
    st.sidebar.error("🚨 데이터 파일을 읽지 못했습니다. 엑셀 파일이 같은 폴더에 있는지 확인해주세요.")

st.sidebar.markdown("---")

# [파트 2] 수시 입결
st.sidebar.header("2. 수시 입결 검색")
s_region, s_univ, s_type, s_susi_dept = "전체", "전체", "전체", "전체"

if not df_susi.empty:
    regs = ["전체"] + sorted(df_susi['지역'].dropna().unique().tolist()) if '지역' in df_susi.columns else ["전체"]
    s_region = st.sidebar.selectbox("지역", regs)
    
    tmp = df_susi[df_susi['지역']==s_region] if s_region != "전체" else df_susi
    unvs = ["전체"] + sorted(tmp['대학'].dropna().unique().tolist()) if '대학' in tmp.columns else ["전체"]
    s_univ = st.sidebar.selectbox("대학", unvs)

    tmp = tmp[tmp['대학']==s_univ] if s_univ != "전체" else tmp
    typs = ["전체"] + sorted(tmp['전형'].dropna().unique().tolist()) if '전형' in tmp.columns else ["전체"]
    s_type = st.sidebar.selectbox("전형", typs)

    tmp = tmp[tmp['전형']==s_type] if s_type != "전체" else tmp
    dps = ["전체"] + sorted(tmp['학과'].dropna().unique().tolist()) if '학과' in tmp.columns else ["전체"]
    s_susi_dept = st.sidebar.selectbox("학과 (입결용)", dps)

# --------------------------------------------------------------------------
# 5. 메인 화면 출력
# --------------------------------------------------------------------------
st.markdown('<div class="main-title">🧭 진로-진학 나침반</div>', unsafe_allow_html=True)
st.markdown("""
<div class="source-text">
    자료 출처: 2025학년도 2022 개정 교육과정 선택 과목 안내서(낱장).pdf, 
    대학어디가(2022~2025)수시입결자료4(목포제일여고 김현석) ver0905, 
    정확한 출처를 찾을 수 없는(인터넷 검색 자료) 전공별 추천도서.slxs&lt;작성자께 감사&gt;
</div>
""", unsafe_allow_html=True)

# ==========================================================================
# [SECTION 1] 학과 특색
# ==========================================================================
if sel_dept != "선택안함" and not df_info.empty:
    st.markdown(f'<div class="dept-title">📘 {sel_dept} <span style="font-size:0.6em; color:#666;">({sel_cat if sel_cat!="전체" else "전체"})</span></div>', unsafe_allow_html=True)

    matches = df_info[df_info.iloc[:, 1].astype(str) == sel_dept]
    if not matches.empty:
        info_row = matches.iloc[0]
        
        # 학과 설명
        desc = get_col_val(info_row, 2)
        st.markdown(f'<div class="desc-box"><b>💡 학과 소개</b><br>{desc}</div>', unsafe_allow_html=True)

        # 권장 선택 과목
        st.markdown('<div class="section-title">📚 권장 선택 과목</div>', unsafe_allow_html=True)
        subj_gen = get_col_val(info_row, 3)
        subj_car = get_col_val(info_row, 4)
        subj_con = get_col_val(info_row, 5)

        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="subj-box theme-blue"><span class="subj-header">📘 일반 선택</span><div class="subj-content">{subj_gen}</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="subj-box theme-orange"><span class="subj-header">📙 진로 선택</span><div class="subj-content">{subj_car}</div></div>', unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="subj-box theme-purple"><span class="subj-header">🔮 융합 선택</span><div class="subj-content">{subj_con}</div></div>', unsafe_allow_html=True)
        
        st.markdown("<br><hr><br>", unsafe_allow_html=True)

        # 추천 도서
        st.markdown('<div class="section-title">📖 전공 추천 도서</div>', unsafe_allow_html=True)
        if not df_book.empty:
            books = df_book
            bk_cat_col = next((c for c in df_book.columns if "계열" in c), df_book.columns[1])
            bk_dept_col = next((c for c in df_book.columns if "전공" in c or "학과" in c), df_book.columns[2])
            bk_name_col = next((c for c in df_book.columns if "도서" in c or "책" in c), df_book.columns[3])

            if sel_cat != "전체":
                books_in_field = books[books[bk_cat_col].astype(str) == sel_cat]
            else:
                books_in_field = books

            matched_books = books_in_field[books_in_field[bk_dept_col].astype(str).apply(lambda x: is_match(sel_dept, x))]

            if matched_books.empty:
                st.info(f"💡 '{sel_dept}' 관련 도서가 없어, '{sel_cat}' 관련 추천 도서를 표시합니다.")
                matched_books = books_in_field

            if not matched_books.empty:
                if bk_name_col:
                    matched_books = matched_books.sort_values(by=bk_name_col)
                st.dataframe(matched_books.iloc[:, :9], hide_index=True, use_container_width=True)
            else:
                st.info("검색된 추천 도서가 없습니다.")
        else:
            st.warning("도서 데이터가 없습니다.")

        # 탐구 주제
        st.markdown('<div class="section-title">🔬 추천 탐구 주제</div>', unsafe_allow_html=True)
        if not df_inq.empty:
            inq_dept_col = next((c for c in df_inq.columns if "학과" in c or "전공" in c), None)
            inq_top_col = next((c for c in df_inq.columns if "주제" in c or "탐구" in c), None)
            inq_sub_col = next((c for c in df_inq.columns if "교과" in c or "과목" in c), None)
            inq_type_col = next((c for c in df_inq.columns if "유형" in c), None)

            if inq_dept_col:
                inq_matches = df_inq[df_inq[inq_dept_col].astype(str).apply(lambda x: is_match(sel_dept, x))]
                
                if not inq_matches.empty:
                    if inq_sub_col:
                        inq_matches = inq_matches.sort_values(by=inq_sub_col)

                    for _, q in inq_matches.iterrows():
                        subj_txt = q[inq_sub_col] if inq_sub_col else "전공"
                        type_txt = q[inq_type_col] if inq_type_col else "탐구"
                        top_txt = q[inq_top_col] if inq_top_col else q.iloc[1]
                        
                        st.markdown(f'''
                        <div class="inq-box">
                            <span class="type-tag">{type_txt}</span>
                            <span class="subject-tag">{subj_txt}</span> 
                            {top_txt}
                        </div>
                        ''', unsafe_allow_html=True)
                else: st.info(f"'{sel_dept}' 관련 탐구 주제가 없습니다.")
        else: st.info("탐구주제 데이터가 없습니다.")

    st.divider()

# ==========================================================================
# [SECTION 2] 수시 입결 분석
# ==========================================================================
if s_susi_dept != "전체":
    st.markdown(f"## 📊 {s_susi_dept} 입시 결과 분석", unsafe_allow_html=True)
    
    if not df_susi.empty:
        cond = (df_susi['학과'] == s_susi_dept)
        if s_region != "전체": cond &= (df_susi['지역'] == s_region)
        if s_univ != "전체": cond &= (df_susi['대학'] == s_univ)
        if s_type != "전체": cond &= (df_susi['전형'] == s_type)
        
        res = df_susi[cond].copy()
        
        if not res.empty:
            res = res.sort_values('연도')
            last_row = res.iloc[-1]
            
            # KPI
            st.markdown("##### 📌 최신 입시 결과 요약")
            k1, k2, k3, k4 = st.columns(4)
            with k1:
                yr = int(last_row['연도']) if pd.notna(last_row['연도']) else "-"
                st.markdown(f"<div class='kpi-container'><div class='kpi-label'>기준 연도</div><div class='kpi-value'>{yr}</div></div>", unsafe_allow_html=True)
            with k2:
                gr = last_row['등급70'] if '등급70' in res.columns else "-"
                st.markdown(f"<div class='kpi-container'><div class='kpi-label'>70% 컷</div><div class='kpi-value'>{gr}</div></div>", unsafe_allow_html=True)
            with k3:
                cp = last_row['경쟁률'] if '경쟁률' in res.columns else "-"
                st.markdown(f"<div class='kpi-container'><div class='kpi-label'>경쟁률</div><div class='kpi-value'>{cp} : 1</div></div>", unsafe_allow_html=True)
            with k4:
                ex = last_row['충원인원'] if '충원인원' in res.columns and last_row['충원인원'] > 0 else (last_row['추합'] if '추합' in res.columns else 0)
                st.markdown(f"<div class='kpi-container'><div class='kpi-label'>충원인원</div><div class='kpi-value'>{int(ex)}명</div></div>", unsafe_allow_html=True)

            st.write("")

            tab1, tab2 = st.tabs(["📈 입결 시각화", "📋 상세 데이터"])
            
            with tab1:
                # 1. 꺾은선 (성적 값 표시 강화)
                if '등급70' in res.columns:
                    fig_grade = px.line(res, x='연도', y='등급70', color='전형', markers=True, 
                                        text='등급70', # 값 표시
                                        title=f"📉 {s_susi_dept} 등급(70%컷) 추이")
                    
                    # 텍스트 포맷팅 및 스타일 강화
                    fig_grade.update_traces(
                        mode="lines+markers+text", 
                        texttemplate='%{text:.2f}',
                        textposition="top center", 
                        textfont=dict(size=16, color='#000000', family='Arial Black')
                    )
                    fig_grade.update_yaxes(autorange="reversed")
                    fig_grade.update_xaxes(type='category')
                    st.plotly_chart(fig_grade, use_container_width=True)

                # 2. 막대 (편안한 색상 & 값 글자 확대)
                bar_cols = []
                if '모집인원' in res.columns: bar_cols.append('모집인원')
                if '경쟁률' in res.columns: bar_cols.append('경쟁률')
                if '실질경쟁률' in res.columns: bar_cols.append('실질경쟁률')
                elif '실경쟁률' in res.columns: bar_cols.append('실경쟁률')
                if '충원인원' in res.columns: bar_cols.append('충원인원')
                elif '추합' in res.columns: bar_cols.append('추합')

                if bar_cols:
                    df_melt = res.melt(id_vars=['연도', '전형'], value_vars=bar_cols, var_name='구분', value_name='값')
                    
                    def format_val(row):
                        if '경쟁률' in str(row['구분']): return f"{row['값']:.2f}"
                        return f"{row['값']:.0f}"
                    
                    df_melt['Label'] = df_melt.apply(format_val, axis=1)
                    
                    # 편안한 색상 (Teal, Salmon, LightBlue, Peach)
                    custom_colors = ['#4DB6AC', '#FF8A65', '#90CAF9', '#FFB74D'] 

                    fig_bar = px.bar(df_melt, x='연도', y='값', color='구분', barmode='group',
                                     text='Label', title="📊 모집·경쟁·충원 현황 비교",
                                     color_discrete_sequence=custom_colors)
                    # 막대 위 숫자 14pt로 확대
                    fig_bar.update_traces(textposition='outside', textfont=dict(size=14))
                    fig_bar.update_xaxes(type='category')
                    st.plotly_chart(fig_bar, use_container_width=True)

            with tab2:
                # 표 표시
                view_cols = ['연도', '지역', '대학', '전형', '학과', '모집인원', '경쟁률', '실질경쟁률', '실경쟁률', '등급50', '등급70', '충원인원', '추합']
                real_cols = [c for c in view_cols if c in res.columns]
                
                format_dict = {
                    '연도': '{:.0f}', '모집인원': '{:.0f}', '충원인원': '{:.0f}', '추합': '{:.0f}',
                    '경쟁률': '{:.2f}', '실경쟁률': '{:.2f}', '실질경쟁률': '{:.2f}', '등급50': '{:.2f}', '등급70': '{:.2f}'
                }
                final_format = {k: v for k, v in format_dict.items() if k in res.columns}

                # 표 내용 중앙 정렬 적용
                st.dataframe(
                    res[real_cols].sort_values('연도', ascending=False)
                    .style
                    .format(final_format)
                    .set_properties(**{'text-align': 'center'})
                    .set_table_styles([dict(selector='th', props=[('text-align', 'center')])]),
                    hide_index=True, 
                    use_container_width=True
                )
        else: st.warning("조건에 맞는 입시 결과가 없습니다.")