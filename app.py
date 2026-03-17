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
    # 4. 전문 AI 에이전트 분석 섹션
    # ============================================================================
    st.divider()
    st.header("📝 구매부서 전용 심층 마켓 보고서 (AI)")

    # 분석 대상 선정: 주간 상승률 상위 품목
    top_items = weekly_df[weekly_df['기간'] == weekly_df['기간'].max()].nlargest(3, '증감률')['품목'].tolist()

    if st.button("🔥 전문 AI 팀 분석 시작"):
        if not os.environ.get("OPENAI_API_KEY"):
            st.error("사이드바에 OpenAI API Key를 입력하거나 Secrets를 설정해주세요.")
        else:
            with st.status("전문 분석팀이 작동 중입니다...", expanded=True) as status:
                # 1. 에이전트 설정
                analyst = Agent(
                    role="시장 변동 원인분석 전문가",
                    goal="가격 변동의 원인을 5대 지표(공급, 수요, 환경, 유통, 연관시장) 관점에서 분석",
                    backstory="15년 경력의 베테랑 시장 분석가입니다.",
                    llm=LLM(model="gpt-4o"),
                    verbose=True
                )
                procurement = Agent(
                    role="구매 인사이트 전문가",
                    goal="분석 결과를 바탕으로 구매 대응 전략 수립",
                    backstory="대기업 원료 구매팀장 출신입니다.",
                    llm=LLM(model="gpt-4o"),
                    verbose=True
                )

                all_reports = []
                for item in top_items:
                    st.write(f"🔍 **{item}** 정밀 분석 중...")
                    t1 = Task(description=f"{item}의 최근 시세 급등 원인을 정밀 분석하세요.", expected_output="원인 분석 보고서", agent=analyst)
                    t2 = Task(description=f"분석된 내용을 토대로 {item} 구매 대응 가이드를 작성하세요.", expected_output="구매 가이드", agent=procurement)
                    
                    crew = Crew(agents=[analyst, procurement], tasks=[t1, t2])
                    report_out = crew.kickoff()
                    all_reports.append(report_out.raw)

                final_report_md = "# 📑 구매부서 종합 시장 분석 보고서\n\n" + "\n\n---\n\n".join(all_reports)
                status.update(label="✅ 분석 완료!", state="complete", expanded=False)

            st.markdown(final_report_md)
            
            # 다운로드 버튼
            docx_file = markdown_to_docx_stream(final_report_md)
            st.download_button(label="📄 Word 보고서 다운로드", data=docx_file, file_name=f"Report_{datetime.date.today()}.docx")
