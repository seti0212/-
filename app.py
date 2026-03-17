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
# 1. 환경 및 보안 설정 (2026년 최신 모델 규격 적용)
# ============================================================================
st.set_page_config(page_title="구매지원팀 통합 분석 시스템 (Gemini)", layout="wide")

# [보안] Streamlit Secrets에서 API 키 로드
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
    os.environ["GEMINI_API_KEY"] = api_key
    # 일부 라이브러리 호환성을 위해 GOOGLE_API_KEY도 동일하게 설정
    os.environ["GOOGLE_API_KEY"] = api_key 
    os.environ["SERPER_API_KEY"] = st.secrets.get("SERPER_API_KEY", "")
else:
    st.error("⚠️ API 키가 설정되지 않았습니다. Streamlit Cloud의 Secrets 설정을 완료해 주세요.")

# 데이터 소스
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
# 2. 통계 계산 및 핵심 이슈 품목 추출
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
    w_top = w_df[w_df['기간'] == w_df['기간'].max()].nlargest(5, '증감률')['품목'].tolist()
    w_bot = w_df[w_df['기간'] == w_df['기간'].max()].nsmallest(5, '증감률')['품목'].tolist()
    m_top = m_df[m_df['기간'] == m_df['기간'].max()].nlargest(5, '증감률')['품목'].tolist()
    m_bot = m_df[m_df['기간'] == m_df['기간'].max()].nsmallest(5, '증감률')['품목'].tolist()
    y_top = y_df[y_df['기간'] == y_df['기간'].max()].nlargest(5, '증감률')['품목'].tolist()
    y_bot = y_df[y_df['기간'] == y_df['기간'].max()].nsmallest(5, '증감률')['품목'].tolist()
    combined = list(set(w_top + w_bot + m_top + m_bot + y_top + y_bot))
    return combined

# ============================================================================
# 3. 메인 대시보드 (기존 3x3 레이아웃)
# ============================================================================
if df_raw is not None:
    weekly_df, monthly_df, yearly_df = calculate_all_stats(df_raw)

    st.title("📊 원자재 시세 실시간 분석 및 전문 AI 보고서")
    st.info(f"데이터 업데이트 시각: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    display_cols = ['품목', '평균시세', '단위', '증감률']
    def format_df(df):
        return df[display_cols].style.format({'평균시세': '{:,.2f}', '증감률': '{:+.2f}%'})

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
    # 4. 전문 AI 분석 섹션 (Gemini 2.0 적용)
    # ============================================================================
    st.header("📝 이슈 구매 품목 종합 보고서 (Gemini 단가 예측)")
    
    critical_items = get_critical_items(weekly_df, monthly_df, yearly_df)
    st.write(f"🔔 **AI 분석 대상 ({len(critical_items)}개 품목):** {', '.join(critical_items)}")

    if st.button("🔥 전체 품목 전문 Gemini 분석 시작"):
        if not os.environ.get("GEMINI_API_KEY") or not os.environ.get("SERPER_API_KEY"):
            st.error("🚨 API 키 설정을 확인해 주세요 (Secrets).")
        else:
            search_tool = SerperDevTool()
            
            # **[수정] 2026년 기준 가장 안정적인 모델 명칭: gemini/gemini-2.0-flash**
            # 이전의 1.5 모델은 deprecated 되어 404 에러가 발생했던 것입니다.
            gemini_llm = LLM(
                model="gemini/gemini-2.0-flash", 
                api_key=os.environ["GEMINI_API_KEY"]
            )

            with st.status("최신 Gemini 엔진으로 분석 보고서를 생성 중입니다...", expanded=True) as status:
                analyst = Agent(
                    role="시장 수급 및 단가 예측 전문가",
                    goal="뉴스와 데이터를 기반으로 품목별 향후 1~3개월 단가 예측",
                    backstory="뉴스 근거를 바탕으로 단가 변동 인과관계를 분석하는 전문 분석가입니다.",
                    llm=gemini_llm,
                    tools=[search_tool],
                    verbose=True
                )
                procurement = Agent(
                    role="전략적 구매 관리 전문가",
                    goal="예측된 단가 흐름에 따라 최적의 구매 시점 및 대응 전략 수립",
                    backstory="원가 절감과 공급 안정성을 설계하는 전략가입니다.",
                    llm=gemini_llm,
                    verbose=True
                )

                all_reports = []
                progress_bar = st.progress(0)
                
                for idx, item in enumerate(critical_items):
                    st.write(f"⏳ ({idx+1}/{len(critical_items)}) **{item}** 분석 중...")
                    
                    t1 = Task(description=f"{item}의 최신 뉴스 기반 향후 3개월 단가 예측", expected_output="원인 및 예측", agent=analyst)
                    t2 = Task(description=f"{item}의 단가 예측에 따른 구매 실행 전략", expected_output="구매 가이드", agent=procurement)
                    
                    # max_rpm=10으로 상향 조정 (Gemini 2.0은 속도가 더 빠름)
                    crew = Crew(agents=[analyst, procurement], tasks=[t1, t2], max_rpm=10, verbose=True)
                    
                    try:
                        report_out = crew.kickoff()
                        all_reports.append(report_out.raw)
                    except Exception as e:
                        st.error(f"❌ {item} 분석 중 오류: {str(e)}")
                        all_reports.append(f"### {item}\n분석 중 에러가 발생하여 생략되었습니다.")
                    
                    time.sleep(1) # API 보호를 위한 짧은 휴식
                    progress_bar.progress((idx + 1) / len(critical_items))

                final_report_md = f"# 📑 구매부서 종합 이슈 보고서 ({datetime.date.today()})\n\n" + "\n\n---\n\n".join(all_reports)
                status.update(label="✅ 모든 품목 분석 완료!", state="complete", expanded=False)

            st.markdown(final_report_md)
            docx_file = markdown_to_docx_stream(final_report_md)
            st.download_button(label="📄 Gemini 보고서 다운로드 (Word)", data=docx_file, file_name=f"Market_Report_{datetime.date.today()}.docx")
