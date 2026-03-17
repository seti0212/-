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
# 1. 환경 설정 및 데이터 로드 (기존 로직 유지)
# ============================================================================
st.set_page_config(page_title="구매지원팀 전문 시세 분석 시스템", layout="wide")

# API 키 설정
os.environ["OPENAI_API_KEY"] = "사용자님의_API_키" # 보안을 위해 관리 주의

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

# Word 변환 유틸리티 (메모리 스트림 방식)
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
# 2. 대시보드 통계 화면 (기존 로직 유지)
# ============================================================================
if df_raw is not None:
    # 통계 계산 로직 (기존과 동일)
    df = df_raw.copy()
    df['연주'] = df['날짜'].dt.to_period('W').astype(str)
    df['연월'] = df['날짜'].dt.to_period('M').astype(str)
    df['연도'] = df['날짜'].dt.year

    def get_stats(df_group, col_name):
        grouped = df_group.groupby(['품목', '단위', col_name])['y'].mean().reset_index()
        grouped.columns = ['품목', '단위', '기간', '평균시세']
        grouped['이전시세'] = grouped.groupby('품목')['평균시세'].shift(1)
        grouped['증감률'] = ((grouped['평균시세'] - grouped['이전시세']) / grouped['이전시세'] * 100).round(2)
        return grouped.fillna(0)

    weekly_df, monthly_df, yearly_df = get_stats(df, '연주'), get_stats(df, '연월'), get_stats(df, '연도')

    st.title("📊 원자재 시세 분석 및 전문 AI 보고서")
    st.info(f"마지막 업데이트: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # 시세 현황 테이블 출력 (생략: 사용자님의 기존 UI 코드와 동일)
    # ... (주간/월간/연간 TOP 5 테이블 코드) ...
    st.header("🔍 실시간 시세 변동 현황")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📈 주간 상승 TOP 5")
        top_risers = weekly_df[weekly_df['기간'] == weekly_df['기간'].max()].nlargest(5, '증감률')
        st.table(top_risers[['품목', '평균시세', '증감률']])
    with col2:
        st.subheader("📉 주간 하락 TOP 5")
        top_fallers = weekly_df[weekly_df['기간'] == weekly_df['기간'].max()].nsmallest(5, '증감률')
        st.table(top_fallers[['품목', '평균시세', '증감률']])

    # ============================================================================
    # 3. 전문 AI 분석 섹션 (요청하신 상세 버전으로 교체)
    # ============================================================================
    st.divider()
    st.header("📝 구매부서 전용 심층 마켓 보고서 (Multi-Agent)")
    
    # 분석 대상 품목 선정 (상승률이 높은 상위 품목들)
    target_items = top_risers['품목'].tolist()

    if st.button("🔥 전문 AI 에이전트 팀 가동 (심층 분석)"):
        with st.status("전문 분석가 팀(시장/구매/전략)이 데이터를 분석 중입니다...", expanded=True) as status:
            
            # 1. 전문 에이전트 정의
            market_analyst = Agent(
                role="농축수산물 시장 변동 원인분석 전문가",
                goal="공급망, 기후, 정책 등 5대 영역에서 가격 변동의 근본 원인을 분석",
                backstory="15년 경력의 시장 분석가. 단순 현상이 아닌 인과관계를 추적합니다.",
                llm=LLM(model="gpt-4o"),
                verbose=True
            )

            procurement_expert = Agent(
                role="구매담당자를 위한 시장 인사이트 전문가",
                goal="분석 결과를 바탕으로 구매 실무자가 취해야 할 전략적 시사점 도출",
                backstory="대기업 구매부서 베테랑. 리스크 조기 경보와 대응 방안 마련에 능숙합니다.",
                llm=LLM(model="gpt-4o"),
                verbose=True
            )

            # 2. 태스크 정의 (상세 버전)
            all_reports = []
            for item in target_items:
                st.write(f"🔎 {item} 품목 정밀 진단 중...")
                
                analysis_task = Task(
                    description=f"{item}의 최근 시세 변동을 1.공급, 2.수요, 3.외부환경, 4.유통구조, 5.연관시장 관점에서 분석하세요.",
                    expected_output=f"{item} 심층 원인 분석서",
                    agent=market_analyst
                )
                
                insight_task = Task(
                    description=f"위 분석을 바탕으로 {item} 구매 담당자를 위한 'So What', '모니터링 포인트', '대응 시점'을 도출하세요.",
                    expected_output=f"{item} 구매 전략 제안서",
                    agent=procurement_expert
                )

                crew = Crew(agents=[market_analyst, procurement_expert], tasks=[analysis_task, insight_task])
                result = crew.kickoff()
                all_reports.append(result.raw)

            final_report_text = "# 📑 구매부서 종합 시장 분석 보고서\n\n" + "\n\n---\n\n".join(all_reports)
            status.update(label="✅ 분석 완료!", state="complete", expanded=False)

        # 결과 표시 및 다운로드
        st.markdown(final_report_text)
        
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            docx_file = markdown_to_docx_stream(final_report_text)
            st.download_button("📄 Word 보고서 다운로드", data=docx_file, file_name="Market_Report.docx")
        with col_dl2:
            st.download_button("📝 MD 파일 저장", data=final_report_text, file_name="Market_Report.md")
