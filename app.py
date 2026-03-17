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
# 1. 환경 및 보안 설정
# ============================================================================
st.set_page_config(page_title="구매지원팀 통합 분석 시스템 (Gemini)", layout="wide")

if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
    os.environ["GEMINI_API_KEY"] = api_key
    os.environ["GOOGLE_API_KEY"] = api_key 
    os.environ["SERPER_API_KEY"] = st.secrets.get("SERPER_API_KEY", "")
else:
    st.error("⚠️ Streamlit Cloud의 Secrets 설정을 완료해 주세요.")

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
# 2. 통계 계산 및 전체 이슈 품목 추출 (주/월/연 각 2개씩 선정)
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
    # 각 기간별로 가장 많이 오른 품목 1개, 가장 많이 내린 품목 1개씩 선정 (총 2개씩)
    w_top = w_df[w_df['기간'] == w_df['기간'].max()].nlargest(1, '증감률')['품목'].tolist()
    w_bot = w_df[w_df['기간'] == w_df['기간'].max()].nsmallest(1, '증감률')['품목'].tolist()
    
    m_top = m_df[m_df['기간'] == m_df['기간'].max()].nlargest(1, '증감률')['품목'].tolist()
    m_bot = m_df[m_df['기간'] == m_df['기간'].max()].nsmallest(1, '증감률')['품목'].tolist()
    
    y_top = y_df[y_df['기간'] == y_df['기간'].max()].nlargest(1, '증감률')['품목'].tolist()
    y_bot = y_df[y_df['기간'] == y_df['기간'].max()].nsmallest(1, '증감률')['품목'].tolist()
    
    # 리스트 통합 및 중복 제거 (최대 6개 품목)
    return list(set(w_top + w_bot + m_top + m_bot + y_top + y_bot))

# ============================================================================
# 3. 메인 대시보드 (기존 3x3 레이아웃)
# ============================================================================
if df_raw is not None:
    weekly_df, monthly_df, yearly_df = calculate_all_stats(df_raw)
    st.title("📊 원자재 시세 실시간 분석 및 전문 AI 보고서")
    st.info(f"데이터 업데이트 시각: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    display_cols = ['품목', '평균시세', '단위', '증감률']
    def format_df(df): return df[display_cols].style.format({'평균시세': '{:,.2f}', '증감률': '{:+.2f}%'})

    for title, func in [("🔍 기간별 전체 시세 현황", None), ("📈 가격 상승 TOP 5", "nlargest"), ("📉 가격 하락 TOP 5", "nsmallest")]:
        st.header(title)
        c1, c2, c3 = st.columns(3)
        for col, data, p_name in zip([c1, c2, c3], [weekly_df, monthly_df, yearly_df], ["🗓️ 주간", "📅 월간", "📂 연간"]):
            with col:
                st.subheader(p_name)
                curr = data[data['기간'] == data['기간'].max()]
                disp = getattr(curr, func)(5, '증감률') if func else curr
                st.dataframe(format_df(disp), use_container_width=True, hide_index=True)
        st.divider()

    # ============================================================================
    # 4. 전문 AI 분석 섹션 (분석 대상 압축 및 안정화 로직 유지)
    # ============================================================================
    st.header("📝 이슈 구매 품목 종합 보고서 (Gemini 2.0)")
    critical_items = get_critical_items(weekly_df, monthly_df, yearly_df)
    st.write(f"🔔 **AI 정밀 분석 대상 (핵심 {len(critical_items)}개 품목):** {', '.join(critical_items)}")

    if st.button("🔥 핵심 품목 정밀 분석 시작"):
        if not os.environ.get("GEMINI_API_KEY"):
            st.error("🚨 API 키가 설정되지 않았습니다.")
        else:
            search_tool = SerperDevTool()
            gemini_llm = LLM(model="gemini/gemini-2.0-flash", api_key=os.environ["GEMINI_API_KEY"])

            with st.status("핵심 품목 순차 분석 중...", expanded=True) as status:
                analyst = Agent(role="시장 예측가", goal="뉴스 기반 단가 예측", backstory="데이터 분석가", llm=gemini_llm, tools=[search_tool])
                procurement = Agent(role="구매 전략가", goal="전략 수립", backstory="구매 전문가", llm=gemini_llm)

                all_reports = []
                progress_bar = st.progress(0)
                
                for idx, item in enumerate(critical_items):
                    st.write(f"🔍 **{item}** 분석 중... ({idx+1}/{len(critical_items)})")
                    
                    t1 = Task(description=f"{item}의 최신 뉴스 기반 3개월 단가 예측", expected_output="원인/예측", agent=analyst)
                    t2 = Task(description=f"{item} 구매 전략 제안", expected_output="전략", agent=procurement)
                    crew = Crew(agents=[analyst, procurement], tasks=[t1, t2], max_rpm=1)

                    success = False
                    for attempt in range(3):
                        try:
                            report_out = crew.kickoff()
                            all_reports.append(report_out.raw)
                            success = True
                            break
                        except Exception as e:
                            if "429" in str(e):
                                wait_time = 15 * (attempt + 1)
                                st.warning(f"⏳ 사용량 한도 초과! {wait_time}초 후 다시 시도합니다... ({attempt+1}/3)")
                                time.sleep(wait_time)
                            else:
                                st.error(f"❌ {item} 오류: {str(e)}")
                                break
                    
                    if not success:
                        all_reports.append(f"### {item}\n분석 한도 초과로 리포트 생성에 실패했습니다.")
                    
                    time.sleep(7)
                    progress_bar.progress((idx + 1) / len(critical_items))

                final_report_md = f"# 📑 핵심 품목 구매 종합 보고서 ({datetime.date.today()})\n\n" + "\n\n---\n\n".join(all_reports)
                status.update(label="✅ 핵심 품목 분석 완료!", state="complete", expanded=False)

            st.markdown(final_report_md)
            docx_file = markdown_to_docx_stream(final_report_md)
            st.download_button(label="📄 Word 다운로드", data=docx_file, file_name=f"Critical_Market_Report_{datetime.date.today()}.docx")
