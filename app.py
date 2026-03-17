import streamlit as st
import pandas as pd
import os
import datetime
import re
import io
import SerperDevTool
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
# 4. 전문 AI 에이전트 분석 섹션 (예측 및 뉴스 근거 강화)
# ============================================================================
st.divider()
st.header("📝 이슈 구매 품목 보고서 (AI)")

# 검색 도구 초기화 (인터넷 뉴스 검색용)
search_tool = SerperDevTool()

def get_critical_items(w_df, m_df, y_df):
    w_top = w_df[w_df['기간'] == w_df['기간'].max()].nlargest(3, '증감률')['품목'].tolist()
    w_bot = w_df[w_df['기간'] == w_df['기간'].max()].nsmallest(3, '증감률')['품목'].tolist()
    m_top = m_df[m_df['기간'] == m_df['기간'].max()].nlargest(2, '증감률')['품목'].tolist()
    m_bot = m_df[m_df['기간'] == m_df['기간'].max()].nsmallest(2, '증감률')['품목'].tolist()
    y_top = y_df[y_df['기간'] == y_df['기간'].max()].nlargest(2, '증감률')['품목'].tolist()
    
    combined_items = list(set(w_top + w_bot + m_top + m_bot + y_top))
    return combined_items

critical_items = get_critical_items(weekly_df, monthly_df, yearly_df)

st.write(f"🔔 **AI 분석 대상 후보:** {', '.join(critical_items)}")
st.caption("※ 최신 뉴스와 시세 데이터를 바탕으로 향후 단가 예측 및 근거를 분석합니다.")

if st.button("🔥 전문 AI 팀 분석 시작"):
    if not os.environ.get("OPENAI_API_KEY"):
        st.error("사이드바에 OpenAI API Key를 입력하거나 Secrets를 설정해주세요.")
    else:
        with st.status("전문 분석팀이 실시간 뉴스와 변동 추이를 분석 중입니다...", expanded=True) as status:
            
            # 1. 에이전트 설정 (뉴스 분석 및 예측 기능 강화)
            analyst = Agent(
                role="농축수산물 시장 수급 예측 전문가",
                goal=f"{datetime.date.today()} 기준 최신 뉴스와 데이터를 바탕으로 품목별 향후 시세 방향성을 예측",
                backstory="""당신은 인터넷 뉴스, 기상보도, 정부 발표 자료를 종합하여 
                단기 및 중기 가격 변동을 정확하게 예측하는 데이터 분석가입니다. 
                단순 현상 나열이 아닌 '뉴스 근거'를 바탕으로 한 예측에 강점이 있습니다.""",
                llm=LLM(model="gpt-4o"),
                tools=[search_tool], # 인터넷 검색 도구 부여
                verbose=True
            )
            
            procurement = Agent(
                role="구매 전략 및 리스크 관리 전문가",
                goal="예측된 단가 변화에 따른 최적의 구매 시점과 대응 시나리오 제시",
                backstory="""당신은 분석가의 예측을 바탕으로 '지금 사야 할지, 기다려야 할지'를 결정합니다. 
                상승 예측 시 선매수 전략을, 하락 예측 시 재고 최소화 전략을 수립합니다.""",
                llm=LLM(model="gpt-4o"),
                verbose=True
            )

            all_reports = []
            
            for item in critical_items:
                st.write(f"🔍 **{item}** (뉴스 검색 및 미래 단가 예측 중...)")
                
                # 태스크 1: 뉴스 기반 원인 분석 및 미래 예측
                t1 = Task(
                    description=f"""
                    작업: {item} 품목에 대한 '최신 인터넷 뉴스'와 '수급 데이터'를 검색하여 분석하세요.
                    내용:
                    1. 최근 시세 변동의 결정적 원인 (뉴스 기사 제목 및 내용 인용)
                    2. 향후 1~3개월간의 단가 예측 (상승 / 하락 / 보합 중 선택)
                    3. 예측 근거: 기후 이변, 생산지 소식, 정부 정책 등 뉴스에서 확인된 구체적 사실 제시
                    """,
                    expected_output=f"{item}의 최신 뉴스 기반 원인 분석 및 향후 단가 예측 보고서",
                    agent=analyst
                )
                
                # 태스크 2: 예측에 따른 구매 실행 가이드
                t2 = Task(
                    description=f"""
                    작업: 분석가가 제시한 {item}의 단가 예측(상승/하락)에 따라 구체적인 구매 행동 지침을 작성하세요.
                    내용:
                    - 예상 단가 흐름: [예: 다음 주부터 완만한 상승세 예상]
                    - 대응 전략: [예: 재고 2주분 추가 확보 필요]
                    - 근거 요약: 뉴스에서 언급된 어떤 리스크 때문에 이 전략을 택했는지 설명
                    """,
                    expected_output=f"{item} 구매 대응 가이드 및 예측 요약",
                    agent=procurement
                )
                
                crew = Crew(agents=[analyst, procurement], tasks=[t1, t2])
                report_out = crew.kickoff()
                all_reports.append(report_out.raw)

            # 최종 보고서 통합
            final_report_md = f"# 📑 구매부서 종합 마켓 이슈 보고서 ({datetime.date.today()})\n\n"
            final_report_md += "## 🎯 핵심 요약: 향후 단가 예측 및 뉴스 근거\n"
            final_report_md += f"본 보고서는 **{', '.join(critical_items)}** 품목에 대한 최신 뉴스 기반 예측을 담고 있습니다.\n\n"
            final_report_md += "\n\n---\n\n".join(all_reports)
            
            status.update(label="✅ 분석 및 단가 예측 완료!", state="complete", expanded=False)

        # 결과 출력 및 다운로드 버튼
        st.markdown(final_report_md)
        
        docx_file = markdown_to_docx_stream(final_report_md)
        st.download_button(
            label="📄 전문 분석 보고서 다운로드 (Word)", 
            data=docx_file, 
            file_name=f"Market_Analysis_Report_{datetime.date.today()}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
