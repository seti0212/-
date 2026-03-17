import streamlit as st
import pandas as pd
import os
import datetime
import re
import io
import time
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from crewai import Agent, Task, Crew, LLM
from crewai_tools import SerperDevTool

# ============================================================================
# 1. 보안 및 환경 설정 (Gemini 기반)
# ============================================================================
st.set_page_config(page_title="구매지원팀 통합 분석 시스템 (Gemini)", layout="wide")

# [보안] Gemini 및 Serper API 키 로드 로직
# 코드에 직접 적지 않고 Streamlit Secrets나 사이드바 입력을 사용합니다.
if "GEMINI_API_KEY" in st.secrets:
    os.environ["GEMINI_API_KEY"] = st.secrets["GEMINI_API_KEY"]
    os.environ["SERPER_API_KEY"] = st.secrets.get("SERPER_API_KEY", "")
else:
    with st.sidebar:
        st.header("🔑 보안 설정 (Gemini)")
        user_key = st.text_input("Gemini API Key를 입력하세요", type="password")
        serper_key = st.text_input("Serper API Key (뉴스 검색용)", type="password")
        if user_key: os.environ["GEMINI_API_KEY"] = user_key
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
    """주/월/연 전체에서 상승 및 하락 이슈 품목을 모두 수집"""
    w_top = w_df[w_df['기간'] == w_df['기간'].max()].nlargest(5, '증감률')['품목'].tolist()
    w_bot = w_df[w_df['기간'] == w_df['기간'].max()].nsmallest(5, '증감률')['품목'].tolist()
    m_top = m_df[m_df['기간'] == m_df['기간'].max()].nlargest(5, '증감률')['품목'].tolist()
    m_bot = m_df[m_df['기간'] == m_df['기간'].max()].nsmallest(5, '증감률')['품목'].tolist()
    y_top = y_df[y_df['기간'] == y_df['기간'].max()].nlargest(5, '증감률')['품목'].tolist()
    y_bot = y_df[y_df['기간'] == y_df['기간'].max()].nsmallest(5, '증감률')['품목'].tolist()
    
    combined = list(set(w_top + w_bot + m_top + m_bot + y_top + y_bot))
    return combined

# ============================================================================
# 3. 메인 대시보드 (9개 테이블 3x3 레이아웃 유지)
# ============================================================================
if df_raw is not None:
    weekly_df, monthly_df, yearly_df = calculate_all_stats(df_raw)

    st.title("📊 원자재 시세 실시간 분석 및 전문 AI 보고서")
    st.info(f"데이터 업데이트 시각: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    display_cols = ['품목', '평균시세', '단위', '증감률']
    def format_df(df):
        return df[display_cols].style.format({'평균시세': '{:,.2f}', '증감률': '{:+.2f}%'})

    # 3x3 레이아웃 구성
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
    # 4. 전문 AI 분석 섹션 (Gemini 기반 & 이슈 품목 전체 분석)
    # ============================================================================
    st.header("📝 이슈 구매 품목 종합 보고서 (Gemini 단가 예측)")
    
    critical_items = get_critical_items(weekly_df, monthly_df, yearly_df)
    st.write(f"🔔 **AI 분석 대상 ({len(critical_items)}개 품목):** {', '.join(critical_items)}")
    st.caption("※ 모든 이슈 품목의 최신 뉴스를 검색하여 향후 1~3개월 단가를 예측합니다.")

    if st.button("🔥 전체 품목 전문 Gemini 분석 시작"):
        # Gemini 및 Serper 키 확인
        if not os.environ.get("GEMINI_API_KEY") or not os.environ.get("SERPER_API_KEY"):
            st.error("보안 설정(Gemini & Serper Key)을 완료해주세요.")
        else:
            search_tool = SerperDevTool()
            # Gemini 모델 설정 (gemini-1.5-flash 모델 사용)
            gemini_llm = LLM(model="gemini/gemini-1.5-flash")

            with st.status("Gemini 분석팀이 뉴스 검색 및 단가 예측 시나리오를 작성 중입니다...", expanded=True) as status:
                analyst = Agent(
                    role="시장 수급 및 단가 예측 전문가",
                    goal="최신 뉴스와 데이터를 기반으로 품목별 향후 1~3개월 단가 방향성 예측",
                    backstory="뉴스 보도와 수급 데이터를 종합하여 가격 변동의 근거를 명확히 밝히는 분석가입니다.",
                    llm=gemini_llm,
                    tools=[search_tool],
                    verbose=True
                )
                procurement = Agent(
                    role="전략적 구매 관리 전문가",
                    goal="예측된 단가 흐름에 따라 최적의 구매 시점 및 대응 전략 수립",
                    backstory="원가 절감과 공급 안정성을 최우선으로 하는 구매 전략가입니다.",
                    llm=gemini_llm,
                    verbose=True
                )

                all_reports = []
                progress_bar = st.progress(0)
                
                for idx, item in enumerate(critical_items):
                    st.write(f"⏳ ({idx+1}/{len(critical_items)}) **{item}** 분석 및 예측 진행 중...")
                    
                    t1 = Task(
                        description=f"{item}의 최근 급등락 원인을 뉴스에서 찾아 분석하고, 향후 1~3개월 단가(상승/하락/보합)를 뉴스 근거와 함께 예측하세요.",
                        expected_output=f"{item} 뉴스 기반 분석 및 단가 예측 보고서",
                        agent=analyst
                    )
                    t2 = Task(
                        description=f"{item}의 예측 결과에 따른 구매 실행 전략(매수 타이밍, 재고 확보 등)을 제안하세요.",
                        expected_output=f"{item} 구매 가이드",
                        agent=procurement
                    )
                    
                    # Gemini의 속도를 고려하여 max_rpm 조정
                    crew = Crew(agents=[analyst, procurement], tasks=[t1, t2], max_rpm=10, verbose=True)
                    
                    try:
                        report_out = crew.kickoff()
                        all_reports.append(report_out.raw)
                    except Exception as e:
                        st.error(f"❌ {item} 분석 중 오류: {e}")
                        all_reports.append(f"### {item}\n분석 중 오류가 발생했습니다.")
                    
                    # 안정적인 처리를 위한 짧은 휴식
                    time.sleep(1)
                    progress_bar.progress((idx + 1) / len(critical_items))

                final_report_md = f"# 📑 구매부서 종합 이슈 보고서 ({datetime.date.today()})\n\n"
                final_report_md += f"본 보고서는 Gemini 엔진을 통해 총 **{len(critical_items)}개** 핵심 품목을 분석했습니다.\n\n"
                final_report_md += "\n\n---\n\n".join(all_reports)
                
                status.update(label="✅ 모든 품목 Gemini 분석 완료!", state="complete", expanded=False)

            st.markdown(final_report_md)
            docx_file = markdown_to_docx_stream(final_report_md)
            st.download_button(label="📄 Gemini 전문 보고서 다운로드 (Word)", data=docx_file, file_name=f"Full_Gemini_Report_{datetime.date.today()}.docx")
