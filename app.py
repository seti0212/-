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
from crewai_tools import SerperDevTool # 뉴스 검색을 위한 도구

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
        serper_key = st.text_input("Serper API Key (뉴스 검색용)", type="password")
        if user_key: os.environ["OPENAI_API_KEY"] = user_key
        if serper_key: os.environ["SERPER_API_KEY"] = serper_key

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
# 2. 통계 계산 및 핵심 품목 추출 로직
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
    """주/월/연 변동 폭이 큰 핵심 이슈 품목들을 중복 없이 추출"""
    w_top = w_df[w_df['기간'] == w_df['기간'].max()].nlargest(3, '증감률')['품목'].tolist()
    w_bot = w_df[w_df['기간'] == w_df['기간'].max()].nsmallest(3, '증감률')['품목'].tolist()
    m_top = m_df[m_df['기간'] == m_df['기간'].max()].nlargest(2, '증감률')['품목'].tolist()
    m_bot = m_df[m_df['기간'] == m_df['기간'].max()].nsmallest(2, '증감률')['품목'].tolist()
    y_top = y_df[y_df['기간'] == y_df['기간'].max()].nlargest(2, '증감률')['품목'].tolist()
    
    combined_items = list(set(w_top + w_bot + m_top + m_bot + y_top))
    return combined_items

# ============================================================================
# 3. 메인 대시보드 (9개 테이블)
# ============================================================================
if df_raw is not None:
    weekly_df, monthly_df, yearly_df = calculate_all_stats(df_raw)

    st.title("📊 원자재 시세 실시간 분석 및 전문 AI 보고서")
    st.info(f"데이터 업데이트 시각: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    display_cols = ['품목', '평균시세', '단위', '증감률']
    def format_df(df):
        return df[display_cols].style.format({'평균시세': '{:,.2f}', '증감률': '{:+.2f}%'})

    # 시세표 레이아웃 (주/월/연 전체, 상승 TOP5, 하락 TOP5)
    for section_title, func_name in [("🔍 기간별 전체 시세 현황", None), ("📈 가격 상승 TOP 5", "nlargest"), ("📉 가격 하락 TOP 5", "nsmallest")]:
        st.header(section_title)
        c1, c2, c3 = st.columns(3)
        for col, data, period_name in zip([c1, c2, c3], [weekly_df, monthly_df, yearly_df], ["🗓️ 주간", "📅 월간", "📂 연간"]):
            with col:
                st.subheader(f"{period_name}")
                current_data = data[data['기간'] == data['기간'].max()]
                if func_name:
                    display_data = getattr(current_data, func_name)(5, '증감률')
                else:
                    display_data = current_data
                st.dataframe(format_df(display_data), use_container_width=True, hide_index=True)
        st.divider()

    # ============================================================================
    # 4. 전문 AI 에이전트 분석 섹션 (예측 강화)
    # ============================================================================
    st.header("📝 이슈 구매 품목 보고서 (AI 단가 예측)")
    
    critical_items = get_critical_items(weekly_df, monthly_df, yearly_df)
    st.write(f"🔔 **AI 분석 대상 후보:** {', '.join(critical_items)}")
    st.caption("※ 최신 뉴스와 수급 데이터를 바탕으로 향후 1~3개월 단가를 예측합니다.")

    search_tool = SerperDevTool()

    if st.button("🔥 전문 AI 팀 분석 시작"):
        if not os.environ.get("OPENAI_API_KEY") or not os.environ.get("SERPER_API_KEY"):
            st.error("사이드바에 OpenAI 및 Serper API Key를 입력해주세요.")
        else:
            with st.status("실시간 뉴스 검색 및 단가 예측 중...", expanded=True) as status:
                # 에이전트 설정
                analyst = Agent(
                    role="농축수산물 시장 수급 예측 전문가",
                    goal=f"{datetime.date.today()} 기준 최신 뉴스를 바탕으로 품목별 향후 시세 방향성 예측",
                    backstory="뉴스 보도와 정부 데이터를 종합하여 단기/중기 시세를 정확히 예측하는 분석가입니다.",
                    llm=LLM(model="gpt-4o"),
                    tools=[search_tool],
                    verbose=True
                )
                procurement = Agent(
                    role="구매 전략 및 리스크 관리 전문가",
                    goal="예측된 단가 변화에 따른 최적의 구매 시점 제시",
                    backstory="분석가의 예측을 바탕으로 선매수 혹은 대기 전략을 수립하는 구매 전략가입니다.",
                    llm=LLM(model="gpt-4o"),
                    verbose=True
                )

                all_reports = []
                for item in critical_items:
                    st.write(f"🔍 **{item}** 분석 중...")
                    t1 = Task(
                        description=f"{item}의 최근 급등락 원인을 뉴스에서 찾아 분석하고, 향후 1~3개월 단가(상승/하락/보합)를 예측하세요. 반드시 뉴스 기사 근거를 포함하세요.",
                        expected_output=f"{item} 원인 분석 및 단가 예측 보고서",
                        agent=analyst
                    )
                    t2 = Task(
                        description=f"{item}의 예측 결과에 따른 구체적인 구매 대응 전략(구매 시점, 재고 확보 등)을 제안하세요.",
                        expected_output=f"{item} 구매 가이드",
                        agent=procurement
                    )
                    crew = Crew(agents=[analyst, procurement], tasks=[t1, t2])
                    report_out = crew.kickoff()
                    all_reports.append(report_out.raw)

                final_report_md = f"# 📑 구매부서 종합 마켓 이슈 보고서 ({datetime.date.today()})\n\n" + "\n\n---\n\n".join(all_reports)
                status.update(label="✅ 분석 및 예측 완료!", state="complete", expanded=False)

            st.markdown(final_report_md)
            docx_file = markdown_to_docx_stream(final_report_md)
            st.download_button(label="📄 Word 보고서 다운로드", data=docx_file, file_name=f"Market_Report_{datetime.date.today()}.docx")
