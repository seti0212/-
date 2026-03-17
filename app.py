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

    # 시세표 레이아웃 (기존 틀 유지)
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
    # 4. 전문 AI 에이전트 분석 섹션 (예측 및 뉴스 근거 강화)
    # ============================================================================
    st.header("📝 이슈 구매 품목 보고서 (AI 단가 예측)")
    
    critical_items = get_critical_items(weekly_df, monthly_df, yearly_df)
    st.write(f"🔔 **AI 분석 대상 후보:** {', '.join(critical_items)}")
    st.caption("※ 실시간 뉴스와 변동 데이터를 결합하여 향후 단가 예측 및 전략적 근거를 도출합니다.")

    # 버튼 클릭 시에만 도구 및 에이전트 초기화 (속도 및 보안 최적화)
    if st.button("🔥 전문 AI 팀 분석 시작"):
        if not os.environ.get("OPENAI_API_KEY") or not os.environ.get("SERPER_API_KEY"):
            st.error("사이드바에 OpenAI 및 Serper API Key를 입력해주세요.")
        else:
            search_tool = SerperDevTool() # 버튼 클릭 후 키가 로드된 상태에서 초기화
            
            with st.status("실시간 뉴스 검색 및 단가 예측 시나리오 작성 중...", expanded=True) as status:
                # 1. 원인 분석 및 단가 예측 에이전트
                analyst = Agent(
                    role="농축수산물 수급 및 단가 예측 전문가",
                    goal=f"{datetime.date.today()} 기준 최신 인터넷 뉴스를 검색하여 품목별 단가 변동의 근거를 찾고 향후 시세를 예측",
                    backstory="""당신은 기상보도, 수출입 통계, 뉴스 기사를 종합 분석하여 
                    미래 단가를 예측하는 수석 분석가입니다. 특히 뉴스의 핵심 내용을 인용하여 
                    예측의 신뢰성을 높이는 데 탁월한 능력이 있습니다.""",
                    llm=LLM(model="gpt-4o"),
                    tools=[search_tool],
                    verbose=True
                )

                # 2. 구매 대응 전략 에이전트
                procurement = Agent(
                    role="구매 전략 및 원가 관리 전문가",
                    goal="분석가의 예측 결과를 바탕으로 최적의 구매 시점과 대응 가이드 제시",
                    backstory="분석된 단가 흐름에 따라 선매수, 구매 대기, 혹은 대체재 확보 등 실질적인 구매 액션 플랜을 설계합니다.",
                    llm=LLM(model="gpt-4o"),
                    verbose=True
                )

                all_reports = []
                for item in critical_items:
                    st.write(f"🔍 **{item}** 분석 및 미래 단가 예측 중...")
                    
                    # Task 1: 뉴스 근거 기반 분석 및 예측
                    t1 = Task(
                        description=f"""
                        1. {item} 품목에 대한 최신 뉴스(최근 1개월 이내)를 검색하세요.
                        2. 뉴스에서 언급된 가격 변동의 결정적 원인(기사 제목 혹은 핵심 내용)을 요약하세요.
                        3. 이를 바탕으로 향후 1~3개월간의 단가가 '상승', '하락', '보합' 중 어떻게 변할지 예측하고 그 근거를 제시하세요.
                        """,
                        expected_output=f"{item}의 뉴스 근거 중심 원인 분석 및 단가 예측 보고서",
                        agent=analyst
                    )
                    
                    # Task 2: 전략 수립
                    t2 = Task(
                        description=f"""
                        분석가가 예측한 {item}의 단가 흐름에 따라 구매 부서가 취해야 할 구체적인 행동(구매 시점, 확보 물량 등)을 제안하세요.
                        """,
                        expected_output=f"{item} 구매 대응 가이드",
                        agent=procurement
                    )
                    
                    crew = Crew(agents=[analyst, procurement], tasks=[t1, t2])
                    report_out = crew.kickoff()
                    all_reports.append(report_out.raw)

                final_report_md = f"# 📑 구매부서 종합 마켓 이슈 보고서 ({datetime.date.today()})\n\n" + "\n\n---\n\n".join(all_reports)
                status.update(label="✅ 모든 품목 분석 및 단가 예측 완료!", state="complete", expanded=False)

            # 결과 화면 출력
            st.markdown(final_report_md)
            
            # 워드 다운로드 버튼
            docx_file = markdown_to_docx_stream(final_report_md)
            st.download_button(
                label="📄 전문 분석 보고서 다운로드 (Word)", 
                data=docx_file, 
                file_name=f"Market_Analysis_Report_{datetime.date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
