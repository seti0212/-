import streamlit as st
import pandas as pd
import os
import datetime
from crewai import Agent, Task, Crew, LLM

# ============================================================================
# 1. 환경 설정 및 데이터 로드
# ============================================================================

st.set_page_config(page_title="구매지원팀 시세 분석 시스템", layout="wide")

# OpenAI API 키 설정
os.environ["OPENAI_API_KEY"] = "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"

# 구글 스프레드시트 게시용 URL
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
# 3. 메인 화면 구성
# ============================================================================

if df_raw is not None:
    weekly_df, monthly_df, yearly_df = calculate_all_stats(df_raw)

    st.title("📊 원자재 시세 실시간 분석 및 AI 보고서")
    st.success("✅ 구글 시트와 성공적으로 연결되었습니다.")
    st.info(f"업데이트 시각: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # --- 공통 설정 ---
    display_cols = ['품목', '평균시세', '단위', '증감률']
    def format_df(df, is_top=False):
        target_df = df if not is_top else df
        return target_df[display_cols].style.format({
            '평균시세': '{:,.2f}',
            '증감률': '{:+.2f}%'
        })

    # --- 1. 기간별 전체 시세 현황 ---
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

    # --- 2. 변동률 TOP 5 (상승) ---
    st.header("📈 가격 상승 TOP 5 (주간/월간/연간)")
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

    # --- 3. 변동률 TOP 5 (하락) ---
    st.header("📉 가격 하락 TOP 5 (주간/월간/연간)")
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
    # 4. AI 에이전트 보고서 작성
    # ============================================================================
    st.divider()
    st.header("📝 구매부서 전용 종합 마켓 보고서 (AI)")

    if st.button("전문 AI 분석 보고서 생성"):
        with st.spinner("전문 보고서 작성자가 데이터를 종합 분석 중입니다..."):
            report_writer = Agent(
                role="구매부서를 위한 시장 인사이트 보고서 전문 작성자",
                goal="주요 품목의 변동 원인을 파악하고 전략적 시사점 도출",
                backstory="글로벌 식품 기업의 베테랑 시장 분석가입니다.",
                llm=LLM(model="gpt-4o"),
                verbose=True
            )

            report_task = Task(
                description="최신 주간/월간/연간 상승 및 하락 데이터를 바탕으로 통합 리포트를 작성하세요.",
                agent=report_writer,
                expected_output="마크다운 형식의 보고서"
            )

            crew = Crew(agents=[report_writer], tasks=[report_task])
            report_result = crew.kickoff()

            st.markdown("---")
            st.markdown(report_result.raw)

            st.download_button(
                label="보고서 저장하기 (.md)",
                data=report_result.raw,
                file_name=f"Market_Analysis_Report.md",
                mime="text/markdown"
            )
