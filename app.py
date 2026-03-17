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
# 1. 환경 설정 및 데이터 로드
# ============================================================================
st.set_page_config(page_title="구매지원팀 통합 시세 분석 시스템", layout="wide")

# OpenAI API 키 (보안을 위해 환경 변수나 Streamlit secrets 사용 권장)
os.environ["OPENAI_API_KEY"] = "sk-proj-..." 

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

# Word 변환 유틸리티
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
        elif line.startswith('- ') or line.startswith('* '): doc.add_paragraph(line[2:], style='List Bullet')
        else: doc.add_paragraph(line)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

df_raw = load_data()

# ============================================================================
# 2. 통계 계산 함수 (기존 로직 유지)
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
# 3. 메인 대시보드 UI (기존 3x3 레이아웃 완벽 복구)
# ============================================================================
if df_raw is not None:
    weekly_df, monthly_df, yearly_df = calculate_all_stats(df_raw)

    st.title("📊 원자재 시세 실시간 분석 및 전문 AI 보고서")
    st.success("✅ 데이터가 성공적으로 업데이트되었습니다.")
    st.info(f"업데이트 시각: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # 공통 포맷 함수
    display_cols = ['품목', '평균시세', '단위', '증감률']
    def format_df(df):
        return df[display_cols].style.format({
            '평균시세': '{:,.2f}',
            '증감률': '{:+.2f}%'
        })

    # --- Section 1. 기간별 전체 시세 현황 ---
    st.header("🔍 기간별 전체 시세 현황")
    m_col1, m_col2, m_col3 = st.columns(3)
    with m_col1:
        st.subheader("🗓️ 주간 평균")
        st.dataframe(format_df(weekly_df[weekly_df['기간'] == weekly_df['기간'].max()]), use_container_width=True, hide_index=True)
    with m_col2:
        st.subheader("📅 월간 평균")
        st.dataframe(format_df(monthly_df[monthly_df['기간'] == monthly_df['기간'].max()]), use_container_width=True, hide_index=True)
    with m_col3:
        st.subheader("📂 연간 평균")
        st.dataframe(format_df(yearly_df[yearly_df['기간'] == yearly_df['기간'].max()]), use_container_width=True, hide_index=True)

    st.divider()

    # --- Section 2. 가격 상승 TOP 5 ---
    st.header("📈 가격 상승 TOP 5")
    up_col1, up_col2, up_col3 = st.columns(3)
    with up_col1:
        st.subheader("🗓️ 주간 상승")
        st.dataframe(format_df(weekly_df[weekly_df['기간'] == weekly_df['기간'].max()].nlargest(5, '증감률')), use_container_width=True, hide_index=True)
    with up_col2:
        st.subheader("📅 월간 상승")
        st.dataframe(format_df(monthly_df[monthly_df['기간'] == monthly_df['기간'].max()].nlargest(5, '증감률')), use_container_width=True, hide_index=True)
    with up_col3:
        st.subheader("📂 연간 상승")
        st.dataframe(format_df(yearly_df[yearly_df['기간'] == yearly_df['기간'].max()].nlargest(5, '증감률')), use_container_width=True, hide_index=True)

    # --- Section 3. 가격 하락 TOP 5 ---
    st.header("📉 가격 하락 TOP 5")
    down_col1, down_col2, down_col3 = st.columns(3)
    with down_col1:
        st.subheader("🗓️ 주간 하락")
        st.dataframe(format_df(weekly_df[weekly_df['기간'] == weekly_df['기간'].max()].nsmallest(5, '증감률')), use_container_width=True, hide_index=True)
    with down_col2:
        st.subheader("📅 월간 하락")
        st.dataframe(format_df(monthly_df[monthly_df['기간'] == monthly_df['기간'].max()].nsmallest(5, '증감률')), use_container_width=True, hide_index=True)
    with down_col3:
        st.subheader("📂 연간 하락")
        st.dataframe(format_df(yearly_df[yearly_df['기간'] == yearly_df['기간'].max()].nsmallest(5, '증감률')), use_container_width=True, hide_index=True)

    # ============================================================================
    # 4. 전문 AI 에이전트 보고서 섹션 (강화된 로직)
    # ============================================================================
    st.divider()
    st.header("📝 구매부서 전용 심층 마켓 보고서 (AI)")

    # 분석 대상 선정 (주간 상승률이 높은 품목들 자동 추출)
    top_items = weekly_df[weekly_df['기간'] == weekly_df['기간'].max()].nlargest(5, '증감률')['품목'].tolist()

    if st.button("🔥 전문 AI 에이전트 가동 (심층 분석 보고서 생성)"):
        with st.status("전문 분석가 팀이 품목별 인과관계를 추적 중입니다...", expanded=True) as status:
            
            # 에이전트 페르소나 설정
            analyst = Agent(
                role="농축수산물 시장 변동 원인분석 전문가",
                goal="가격 변동의 근본 원인을 5대 영역(공급/수요/정책/유통/연관시장)에서 분석",
                backstory="15년 경력의 시장 분석가. 단순 현상을 넘어선 구조적 변화를 포착합니다.",
                llm=LLM(model="gpt-4o"),
                verbose=True
            )
            
            procurement = Agent(
                role="구매담당자를 위한 시장 인사이트 전문가",
                goal="분석 결과를 바탕으로 실무적 대응 전략과 모니터링 포인트 제시",
                backstory="대기업 구매부서 베테랑. 리스크 관리와 구매 타이밍 최적화 전문가입니다.",
                llm=LLM(model="gpt-4o"),
                verbose=True
            )

            all_reports = []
            for item in top_items:
                st.write(f"🔍 **{item}** 품목 심층 분석 중...")
                
                t1 = Task(
                    description=f"{item}의 최근 시세 급등 원인을 공급/수요/환경/유통/연관성 5개 지표로 상세 분석하세요.",
                    expected_output=f"{item} 원인 분석 보고서",
                    agent=analyst
                )
                t2 = Task(
                    description=f"위 분석을 바탕으로 {item} 구매 담당자를 위한 전략(So What?)과 주의 신호를 도출하세요.",
                    expected_output=f"{item} 구매 전략 가이드",
                    agent=procurement
                )

                crew = Crew(agents=[analyst, procurement], tasks=[t1, t2])
                report_out = crew.kickoff()
                all_reports.append(report_out.raw)

            final_report_md = "# 📑 구매부서 종합 시장 분석 보고서\n\n" + "\n\n---\n\n".join(all_reports)
            status.update(label="✅ 보고서 작성 완료!", state="complete", expanded=False)

        # 결과 출력 및 다운로드
        st.markdown("---")
        st.markdown(final_report_md)

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            docx_file = markdown_to_docx_stream(final_report_md)
            st.download_button(
                label="📄 Word 보고서 다운로드 (.docx)",
                data=docx_stream,
                file_name=f"Market_Report_{datetime.date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        with col_dl2:
            st.download_button(
                label="📝 마크다운 파일 저장 (.md)",
                data=final_report_md,
                file_name=f"Market_Report_{datetime.date.today()}.md",
                mime="text/markdown"
            )
