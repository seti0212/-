import streamlit as st
import pandas as pd
import os
import datetime
import re
import io
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from crewai import Agent, Task, Crew, LLM

# ============================================================================
# 1. 보안 및 환경 설정 (API 키 숨기기)
# ============================================================================
st.set_page_config(page_title="구매지원팀 통합 분석 시스템", layout="wide")

# [보안] API 키를 불러오는 3단계 전략
if "OPENAI_API_KEY" in st.secrets:
    # 1순위: Streamlit Cloud의 Secrets에 저장된 키 사용
    os.environ["OPENAI_API_KEY"] = st.secrets["OPENAI_API_KEY"]
else:
    # 2순위: 키가 설정되지 않은 경우 사이드바에서 사용자에게 직접 입력받음
    with st.sidebar:
        st.header("🔑 보안 설정")
        user_key = st.text_input("OpenAI API Key를 입력하세요", type="password")
        if user_key:
            os.environ["OPENAI_API_KEY"] = user_key
            st.success("API 키가 임시 설정되었습니다.")
        else:
            st.warning("AI 기능을 사용하려면 API 키가 필요합니다.")

# 구글 스프레드시트 URL
url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vST3eDNhF1GLc231d4RdAnSCb8DnSznnZ4lJfPxxmtIHIcuEXbvFmrBI9LRdbURog-ik09vSOHTOAMp/pub?output=csv"

@st.cache_data(ttl=600)
def load_data():
    try:
        data = pd.read_csv(url)
        data['날짜'] = pd.to_datetime(data['날짜'])
        return data.sort_values(['품목', '날짜'])
    except Exception as e:
        st.error(f"데이터 로드 중 오류 발생: {e}")
        return None

# Word 파일 생성 유틸리티
def markdown_to_docx_stream(markdown_text):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
    lines = markdown_text.split('\n')
    for line in lines:
        line = line.strip()
        if not line: continue
        if line.startswith('# '): doc.add_heading(line[2:], level=0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif line.startswith('## '): doc.add_heading(line[3:], level=1)
        elif line.startswith('### '): doc.add_heading(line[4:], level=2)
        else: doc.add_paragraph(line)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

df_raw = load_data()

# ============================================================================
# 2. 통계 계산 함수
# ============================================================================
def calculate_all_stats(df):
    df = df.copy()
    df['연주'] = df['날짜'].dt.to_period('W').astype(str)
    df['연월'] = df['날짜'].dt.to_period('M').astype(str)
    df['연도'] = df['날짜'].dt.year

    def get_stats(df_group, col_name):
        grouped = df_group.groupby(['품목', '단위', col_name])['y'].mean().reset_index()
        grouped.columns = ['품목', '단위', '기간', '평균시세']
        grouped['이전시세'] = grouped.groupby('품목')['평균시세'].shift(1)
        grouped['증감률'] = ((grouped['평균시세'] - grouped['이전시세']) / grouped['이전시세'] * 100).round(2)
        return grouped.fillna(0)

    return get_stats(df, '연주'), get_stats(df, '연월'), get_stats(df, '연도')

# ============================================================================
# 3. 메인 대시보드 (9개 테이블 3x3 레이아웃)
# ============================================================================
if df_raw is not None:
    weekly_df, monthly_df, yearly_df = calculate_all_stats(df_raw)

    st.title("📊 원자재 시세 실시간 분석 및 전문 AI 보고서")
    st.info(f"데이터 업데이트 시각: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # 공통 데이터 포맷
    display_cols = ['품목', '평균시세', '단위', '증감률']
    def format_df(df):
        return df[display_cols].style.format({'평균시세': '{:,.2f}', '증감률': '{:+.2f}%'})

    # --- Section 1: 기간별 전체 현황 ---
    st.header("🔍 기간별 전체 시세 현황")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.subheader("🗓️ 주간 평균")
        st.dataframe(format_df(weekly_df[weekly_df['기간'] == weekly_df['기간'].max()]), use_container_width=True, hide_index=True)
    with c2:
        st.subheader("📅 월간 평균")
        st.dataframe(format_df(monthly_df[monthly_df['기간'] == monthly_df['기간'].max()]), use_container_width=True, hide_index=True)
    with c3:
        st.subheader("📂 연간 평균")
        st.dataframe(format_df(yearly_df[yearly_df['기간'] == yearly_df['기간'].max()]), use_container_width=True, hide_index=True)

    st.divider()

    # --- Section 2: 가격 상승 TOP 5 ---
    st.header("📈 가격 상승 TOP 5")
    u1, u2, u3 = st.columns(3)
    with u1:
        st.subheader("🗓️ 주간 상승")
        st.dataframe(format_df(weekly_df[weekly_df['기간'] == weekly_df['기간'].max()].nlargest(5, '증감률')), use_container_width=True, hide_index=True)
    with u2:
        st.subheader("📅 월간 상승")
        st.dataframe(format_df(monthly_df[monthly_df['기간'] == monthly_df['기간'].max()].nlargest(5, '증감률')), use_container_width=True, hide_index=True)
    with u3:
        st.subheader("📂 연간 상승")
        st.dataframe(format_df(yearly_df[yearly_df['기간'] == yearly_df['기간'].max()].nlargest(5, '증감률')), use_container_width=True, hide_index=True)

    # --- Section 3: 가격 하락 TOP 5 ---
    st.header("📉 가격 하락 TOP 5")
    d1, d2, d3 = st.columns(3)
    with d1:
        st.subheader("🗓️ 주간 하락")
        st.dataframe(format_df(weekly_df[weekly_df['기간'] == weekly_df['기간'].max()].nsmallest(5, '증감률')), use_container_width=True, hide_index=True)
    with d2:
        st.subheader("📅 월간 하락")
        st.dataframe(format_df(monthly_df[monthly_df['기간'] == monthly_df['기간'].max()].nsmallest(5, '증감률')), use_container_width=True, hide_index=True)
    with d3:
        st.subheader("📂 연간 하락")
        st.dataframe(format_df(yearly_df[yearly_df['기간'] == yearly_df['기간'].max()].nsmallest(5, '증감률')), use_container_width=True, hide_index=True)

# ============================================================================
# 4. 전문 AI 에이전트 분석 섹션 (이슈 품목 자동 선정 로직 강화)
# ============================================================================
st.divider()
st.header("📝 구매부서 전용 심층 마켓 보고서 (AI)")

# --- [로직 개선] 분석 대상 선정: 주/월/연 상승 및 하락 품목 중 핵심 추출 ---
def get_critical_items(w_df, m_df, y_df):
    # 각 카테고리별로 상위/하위 품목 추출 (중복 허용하여 리스트업)
    w_top = w_df[w_df['기간'] == w_df['기간'].max()].nlargest(3, '증감률')['품목'].tolist()
    w_bot = w_df[w_df['기간'] == w_df['기간'].max()].nsmallest(3, '증감률')['품목'].tolist()
    
    m_top = m_df[m_df['기간'] == m_df['기간'].max()].nlargest(2, '증감률')['품목'].tolist()
    m_bot = m_df[m_df['기간'] == m_df['기간'].max()].nsmallest(2, '증감률')['품목'].tolist()
    
    y_top = y_df[y_df['기간'] == y_df['기간'].max()].nlargest(2, '증감률')['품목'].tolist()
    
    # 전체를 합친 후 중복 제거 (Set 활용)
    combined_items = list(set(w_top + w_bot + m_top + m_bot + y_top))
    return combined_items

critical_items = get_critical_items(weekly_df, monthly_df, yearly_df)

st.write(f"🔔 **AI 분석 대상 후보:** {', '.join(critical_items)}")
st.caption("※ 주간/월간/연간 변동폭이 큰 품목들을 중심으로 AI가 종합 분석을 수행합니다.")

if st.button("🔥 전문 AI 팀 분석 시작"):
    if not os.environ.get("OPENAI_API_KEY"):
        st.error("사이드바에 OpenAI API Key를 입력하거나 Secrets를 설정해주세요.")
    else:
        with st.status("전문 분석팀이 기간별 변동 추이를 추적 중입니다...", expanded=True) as status:
            # 1. 에이전트 설정 (페르소나 강화)
            analyst = Agent(
                role="농축수산물 시장 및 거시경제 분석가",
                goal="주간/월간/연간 시세 변동 데이터를 기반으로 가격 급등락의 근본 원인을 분석",
                backstory="당신은 단기적인 수급 불균형뿐만 아니라 장기적인 트렌드(연간 추세)까지 분석하는 20년 경력의 수석 분석가입니다.",
                llm=LLM(model="gpt-4o"),
                verbose=True
            )
            
            procurement = Agent(
                role="전략적 구매 의사결정 전문가",
                goal="시장 분석 보고서를 바탕으로 구매 시점 결정 및 리스크 관리 전략 수립",
                backstory="당신은 가격이 하락할 때는 매수 적기를, 급등할 때는 대체재 확보 시점을 판단하는 구매 전략가입니다.",
                llm=LLM(model="gpt-4o"),
                verbose=True
            )

            all_reports = []
            
            # 선정된 핵심 이슈 품목들을 순회하며 분석
            for item in critical_items:
                st.write(f"🔍 **{item}** (복합 기간 변동성 추적 중...)")
                
                # 태스크 1: 원인 분석 (기간별 맥락 포함)
                t1 = Task(
                    description=f"""
                    품목: {item}
                    대상 데이터: 주간/월간/연간 변동 TOP 리스트에 포함된 이슈 품목입니다.
                    작업: 이 품목의 최근 시세 흐름(급등 혹은 급락)이 발생하는 근본 원인을 5대 지표(공급, 수요, 기후/환경, 유통구조, 연관시장) 관점에서 분석하세요. 
                    특히 일시적인 변동인지, 장기적인 구조적 변화인지 구분하여 설명하세요.
                    """,
                    expected_output=f"{item}의 기간별 변동 원인 분석 보고서",
                    agent=analyst
                )
                
                # 태스크 2: 구매 전략 (So What?)
                t2 = Task(
                    description=f"""
                    시장 분석 결과에 따라 {item}에 대한 구매 대응 가이드를 작성하세요.
                    - 가격 상승세인 경우: 추가 상승 가능성 및 선매수 여부 판단
                    - 가격 하락세인 경우: 바닥 시점 예측 및 매수 타이밍 조언
                    - 공통: 구매 담당자가 매일 체크해야 할 '핵심 모니터링 지표' 제시
                    """,
                    expected_output=f"{item} 구매 전략 제안서",
                    agent=procurement
                )
                
                crew = Crew(agents=[analyst, procurement], tasks=[t1, t2])
                report_out = crew.kickoff()
                all_reports.append(report_out.raw)

            # 최종 보고서 통합
            final_report_md = f"# 📑 구매부서 종합 마켓 이슈 보고서 ({datetime.date.today()})\n\n"
            final_report_md += "## 🎯 이번 주 핵심 분석 품목 요약\n"
            final_report_md += f"본 보고서는 주간/월간/연간 변동성이 가장 컸던 **{', '.join(critical_items)}** 품목을 집중 분석했습니다.\n\n"
            final_report_md += "\n\n---\n\n".join(all_reports)
            
            status.update(label="✅ 분석 및 전략 수립 완료!", state="complete", expanded=False)

        # 결과 출력 및 다운로드 버튼
        st.markdown(final_report_md)
        
        docx_file = markdown_to_docx_stream(final_report_md)
        st.download_button(
            label="📄 전문 분석 보고서 다운로드 (Word)", 
            data=docx_file, 
            file_name=f"Market_Analysis_Report_{datetime.date.today()}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
