import streamlit as st
import pandas as pd
import os
import datetime
import re
import io
import time # **API 호출 간격 조절을 위해 추가**
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from crewai import Agent, Task, Crew, LLM
from crewai_tools import SerperDevTool

# ============================================================================
# 1. 보안 및 환경 설정
# ============================================================================
st.set_page_config(page_title="구매지원팀 통합 분석 시스템", layout="wide")

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
# 2. 통계 및 모든 이슈 품목 추출 로직 (누락 없이 전체 수집)
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
    """주/월/연 전체에서 모든 상승 및 하락 TOP 품목을 추출"""
    w_top = w_df[w_df['기간'] == w_df['기간'].max()].nlargest(5, '증감률')['품목'].tolist()
    w_bot = w_df[w_df['기간'] == w_df['기간'].max()].nsmallest(5, '증감률')['품목'].tolist()
    m_top = m_df[m_df['기간'] == m_df['기간'].max()].nlargest(5, '증감률')['품목'].tolist()
    m_bot = m_df[m_df['기간'] == m_df['기간'].max()].nsmallest(5, '증감률')['품목'].tolist()
    y_top = y_df[y_df['기간'] == y_df['기간'].max()].nlargest(5, '증감률')['품목'].tolist()
    y_bot = y_df[y_df['기간'] == y_df['기간'].max()].nsmallest(5, '증감률')['품목'].tolist()
    
    # 모든 리스트 통합 후 중복 제거
    combined = list(set(w_top + w_bot + m_top + m_bot + y_top + y_bot))
    return combined

# ============================================================================
# 3. 메인 대시보드 (9개 테이블 레이아웃 100% 유지)
# ============================================================================
if df_raw is not None:
    weekly_df, monthly_df, yearly_df = calculate_all_stats(df_raw)

    st.title("📊 원자재 시세 실시간 분석 및 전문 AI 보고서")
    st.info(f"데이터 업데이트 시각: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    display_cols = ['품목', '평균시세', '단위', '증감률']
    def format_df(df):
        return df[display_cols].style.format({'평균시세': '{:,.2f}', '증감률': '{:+.2f}%'})

    # 3x3 레이아웃 틀 유지
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
    # 4. 전문 AI 분석 섹션 (전체 품목 분석 + 에러 방지 최적화)
    # ============================================================================
    st.header("📝 이슈 구매 품목 종합 보고서 (AI 단가 예측)")
    
    critical_items = get_critical_items(weekly_df, monthly_df, yearly_df)
    st.write(f"🔔 **AI 분석 대상 ({len(critical_items)}개 품목):** {', '.join(critical_items)}")
    st.caption("※ 모든 상승/하락 이슈 품목을 분석합니다. 서버 한도 보호를 위해 순차적으로 진행됩니다.")

    if st.button("🔥 전체 품목 전문 AI 분석 시작"):
        if not os.environ.get("OPENAI_API_KEY") or not os.environ.get("SERPER_API_KEY"):
            st.error("보안 설정을 완료해주세요 (API Key 미입력).")
        else:
            search_tool = SerperDevTool()
            
            # **안정성을 위해 한도가 넉넉한 gpt-4o-mini 모델 권장**
            llm_model = LLM(model="gpt-4o-mini")

            with st.status("전체 품목을 정밀 분석 중입니다. 잠시만 기다려주세요...", expanded=True) as status:
                analyst = Agent(
                    role="시장 수급 및 단가 예측 전문가",
                    goal="최신 뉴스와 데이터를 기반으로 품목별 향후 1~3개월 단가 예측",
                    backstory="뉴스 근거를 바탕으로 단가 변동의 인과관계를 밝히는 20년 경력 분석가입니다.",
                    llm=llm_model,
                    tools=[search_tool],
                    verbose=True
                )
                procurement = Agent(
                    role="전략적 구매 관리자",
                    goal="예측된 단가 흐름에 따라 최적의 구매 시점 및 대응 전략 수립",
                    backstory="리스크를 최소화하고 원가를 절감하는 구매 실행 가이드를 작성합니다.",
                    llm=llm_model,
                    verbose=True
                )

                all_reports = []
                progress_bar = st.progress(0)
                
                # **핵심 수정: 모든 품목을 순차적으로 분석하며 짧은 휴식(Sleep) 부여**
                for idx, item in enumerate(critical_items):
                    st.write(f"⏳ ({idx+1}/{len(critical_items)}) **{item}** 분석 중...")
                    
                    t1 = Task(description=f"{item}의 최신 뉴스 근거 기반 향후 3개월 단가 예측", expected_output="원인 및 예측", agent=analyst)
                    t2 = Task(description=f"{item}의 단가 예측에 따른 구매 실행 전략", expected_output="구매 전략", agent=procurement)
                    
                    # max_rpm 설정을 통해 내부적인 속도 제한 추가
                    crew = Crew(agents=[analyst, procurement], tasks=[t1, t2], max_rpm=2, verbose=True)
                    
                    try:
                        report_out = crew.kickoff()
                        all_reports.append(report_out.raw)
                    except Exception as e:
                        st.error(f"❌ {item} 분석 중 오류 발생: {e}")
                        all_reports.append(f"### {item}\n분석 중 오류가 발생했습니다.")
                    
                    # 품목 간 5초의 휴식 시간을 주어 API Rate Limit 회피
                    if idx < len(critical_items) - 1:
                        time.sleep(5)
                    
                    progress_bar.progress((idx + 1) / len(critical_items))

                final_report_md = f"# 📑 구매부서 종합 이슈 보고서 ({datetime.date.today()})\n\n"
                final_report_md += f"본 보고서는 주간/월간/연간 변동성이 큰 총 **{len(critical_items)}개** 품목을 분석했습니다.\n\n"
                final_report_md += "\n\n---\n\n".join(all_reports)
                
                status.update(label="✅ 모든 품목 분석 완료!", state="complete", expanded=False)

            st.markdown(final_report_md)
            docx_file = markdown_to_docx_stream(final_report_md)
            st.download_button(label="📄 전문 보고서 다운로드 (Word)", data=docx_file, file_name=f"Full_Market_Report_{datetime.date.today()}.docx")
