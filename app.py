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
from crewai_tools import SerperDevTool

# ============================================================================
# 1. 보안 및 환경 설정
# ============================================================================
st.set_page_config(page_title="구매지원팀 통합 분석 시스템", layout="wide")

# API 키 보안 로드
if "OPENAI_API_KEY" in st.secrets:
    os.environ["OPENAI_API_KEY"] = st.secrets["OPENAI_API_KEY"]
    os.environ["SERPER_API_KEY"] = st.secrets.get("SERPER_API_KEY", "")
else:
    with st.sidebar:
        st.header("🔑 보안 설정")
        user_key = st.text_input("OpenAI API Key", type="password")
        serper_key = st.text_input("Serper API Key", type="password")
        if user_key: os.environ["OPENAI_API_KEY"] = user_key
        if serper_key: os.environ["SERPER_API_KEY"] = serper_key

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
# 2. 통계 및 핵심 품목 추출 (안정성을 위해 최대 3개로 제한)
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

def get_critical_items(w_df, m_df, y_df):
    # 중복 제거 후 가장 변동이 큰 상위 품목만 추출
    w_top = w_df[w_df['기간'] == w_df['기간'].max()].nlargest(2, '증감률')['품목'].tolist()
    w_bot = w_df[w_df['기간'] == w_df['기간'].max()].nsmallest(1, '증감률')['품목'].tolist()
    m_top = m_df[m_df['기간'] == m_df['기간'].max()].nlargest(1, '증감률')['품목'].tolist()
    
    combined = list(set(w_top + w_bot + m_top))
    return combined[:3] # **에러 방지: 분석 대상을 최대 3개로 엄격히 제한**

# ============================================================================
# 3. 메인 대시보드 (9개 테이블 레이아웃 유지)
# ============================================================================
if df_raw is not None:
    weekly_df, monthly_df, yearly_df = calculate_all_stats(df_raw)

    st.title("📊 원자재 시세 실시간 분석 및 전문 AI 보고서")
    st.info(f"데이터 업데이트 시각: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    display_cols = ['품목', '평균시세', '단위', '증감률']
    def format_df(df):
        return df[display_cols].style.format({'평균시세': '{:,.2f}', '증감률': '{:+.2f}%'})

    # 3x3 테이블 레이아웃 (틀 유지)
    for section_title, func_name in [("🔍 기간별 전체 시세 현황", None), ("📈 가격 상승 TOP 5", "nlargest"), ("📉 가격 하락 TOP 5", "nsmallest")]:
        st.header(section_title)
        c1, c2, c3 = st.columns(3)
        for col, data, period_name in zip([c1, c2, c3], [weekly_df, monthly_df, yearly_df], ["🗓️ 주간", "📅 월간", "📂 연간"]):
            with col:
                st.subheader(period_name)
                curr = data[data['기간'] == data['기간'].max()]
                disp = getattr(curr, func_name)(5, '증감률') if func_name else curr
                st.dataframe(format_df(disp), use_container_width=True, hide_index=True)
        st.divider()

    # ============================================================================
    # 4. 전문 AI 분석 섹션 (RateLimit 방지 최적화)
    # ============================================================================
    st.header("📝 이슈 구매 품목 보고서 (AI 단가 예측)")
    
    critical_items = get_critical_items(weekly_df, monthly_df, yearly_df)
    st.write(f"🔔 **AI 집중 분석 대상:** {', '.join(critical_items)}")
    st.caption("※ API 한도 보호를 위해 가장 이슈가 되는 3개 품목을 정밀 분석합니다.")

    if st.button("🔥 전문 AI 팀 분석 시작"):
        if not os.environ.get("OPENAI_API_KEY") or not os.environ.get("SERPER_API_KEY"):
            st.error("보안 설정을 완료해주세요 (API Key 미입력).")
        else:
            search_tool = SerperDevTool()
            
            with st.status("Rate Limit을 준수하며 신중하게 분석 중입니다...", expanded=True) as status:
                # **모델 변경: gpt-4o-mini (한도가 훨씬 넉넉함)**
                llm_model = LLM(model="gpt-4o-mini")

                analyst = Agent(
                    role="시장 수급 예측 전문가",
                    goal="최신 뉴스를 근거로 향후 단가 예측",
                    backstory="뉴스 데이터를 수집해 가격 변동의 인과관계를 밝힙니다.",
                    llm=llm_model,
                    tools=[search_tool],
                    verbose=True
                )
                procurement = Agent(
                    role="구매 전략 전문가",
                    goal="예측 결과에 따른 구매 실행 가이드 작성",
                    backstory="리스크를 최소화하는 구매 시점을 결정합니다.",
                    llm=llm_model,
                    verbose=True
                )

                all_reports = []
                for item in critical_items:
                    st.write(f"🔍 **{item}** 분석 중 (대기 시간 포함)...")
                    t1 = Task(description=f"{item}의 최근 뉴스 근거 및 향후 3개월 단가 예측", expected_output="원인 및 예측", agent=analyst)
                    t2 = Task(description=f"{item} 구매 대응 전략 수립", expected_output="구매 가이드", agent=procurement)
                    
                    # **핵심 수정: max_rpm=2 설정을 통해 분당 요청 속도 제한**
                    crew = Crew(
                        agents=[analyst, procurement], 
                        tasks=[t1, t2], 
                        max_rpm=2, # 분당 2회 요청으로 제한하여 에러 방지
                        verbose=True
                    )
                    
                    report_out = crew.kickoff()
                    all_reports.append(report_out.raw)

                final_report_md = f"# 📑 구매부서 종합 보고서 ({datetime.date.today()})\n\n" + "\n\n---\n\n".join(all_reports)
                status.update(label="✅ 안정적으로 분석 완료!", state="complete", expanded=False)

            st.markdown(final_report_md)
            docx_file = markdown_to_docx_stream(final_report_md)
            st.download_button(label="📄 Word 다운로드", data=docx_file, file_name=f"Report_{datetime.date.today()}.docx")
