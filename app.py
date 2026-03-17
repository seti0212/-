__import__('pysqlite3')
import sys
sys.modules['sqlite3'] = sys.modules.pop('pysqlite3')

import streamlit as st
import pandas as pd
import os
import datetime
import io
import re
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from crewai import Agent, Task, Crew, LLM

# ============================================================================
# 1. 환경 설정 및 API 키 보안 로드
# ============================================================================

st.set_page_config(page_title="구매지원팀 시세 분석 시스템", layout="wide")

# API 키 보안 처리: st.secrets에서 가져옵니다.
if "OPENAI_API_KEY" in st.secrets:
    os.environ["OPENAI_API_KEY"] = st.secrets["OPENAI_API_KEY"]
else:
    st.error("⚠️ API 키가 설정되지 않았습니다. .streamlit/secrets.toml 파일 혹은 Streamlit Cloud의 Secrets 설정을 확인해 주세요.")
    st.stop()

# 구글 스프레드시트 데이터 URL
DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vST3eDNhF1GLc231d4RdAnSCb8DnSznnZ4lJfPxxmtIHIcuEXbvFmrBI9LRdbURog-ik09vSOHTOAMp/pub?output=csv"

@st.cache_data(ttl=600)
def load_data():
    try:
        data = pd.read_csv(DATA_URL)
        data['날짜'] = pd.to_datetime(data['날짜'])
        return data.sort_values(['품목', '날짜'])
    except Exception as e:
        st.error(f"데이터 로드 실패: {e}")
        return None

df_raw = load_data()

# ============================================================================
# 2. 분석용 데이터 통계 함수
# ============================================================================

def calculate_stats(df):
    df = df.copy()
    df['연주'] = df['날짜'].dt.to_period('W').astype(str)
    df['연월'] = df['날짜'].dt.to_period('M').astype(str)
    
    def get_period_stats(df_group, col_name):
        grouped = df_group.groupby(['품목', '단위', col_name])['y'].mean().reset_index()
        grouped.columns = ['품목', '단위', '기간', '평균시세']
        grouped['이전시세'] = grouped.groupby('품목')['평균시세'].shift(1)
        grouped['증감률'] = ((grouped['평균시세'] - grouped['이전시세']) / grouped['이전시세'] * 100).round(2)
        return grouped.fillna(0)

    return get_period_stats(df, '연주'), get_period_stats(df, '연월')

# ============================================================================
# 3. 마크다운 -> Word 변환 엔진
# ============================================================================

def markdown_to_word_stream(markdown_text):
    doc = Document()
    
    # 문서 기본 여백 설정
    for section in doc.sections:
        section.top_margin = section.bottom_margin = Inches(0.8)
        section.left_margin = section.right_margin = Inches(0.8)

    def process_text(paragraph, text):
        # **볼드체** 및 [숫자] 인용구 서식 적용
        pattern = r'(\*\*.*?\*\*|\[\d+\])'
        parts = re.split(pattern, text)
        for part in parts:
            if not part: continue
            if part.startswith('**') and part.endswith('**'):
                paragraph.add_run(part[2:-2]).bold = True
            elif re.match(r'^\[\d+\]$', part):
                run = paragraph.add_run(part)
                run.font.size, run.font.superscript = Pt(9), True
                run.font.color.rgb = RGBColor(0, 0, 255)
            else:
                paragraph.add_run(part)

    lines = markdown_text.split('\n')
    for line in lines:
        line = line.strip()
        if not line: continue
        
        if line.startswith('# '):
            h = doc.add_heading(line[2:], level=0)
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif line.startswith('## '):
            doc.add_heading(line[3:], level=1)
        elif line.startswith('### '):
            doc.add_heading(line[4:], level=2)
        elif line.startswith('- ') or line.startswith('* '):
            p = doc.add_paragraph(line[2:], style='List Bullet')
            process_text(p, line[2:])
        else:
            process_text(doc.add_paragraph(), line)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ============================================================================
# 4. 메인 대시보드 UI
# ============================================================================

if df_raw is not None:
    weekly_df, monthly_df = calculate_stats(df_raw)

    st.title("👨‍💻 구매지원팀: 원자재 시세 분석 및 전략 보고서")
    st.info(f"💡 현재 시각: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | 데이터 출처: Google Sheets")

    # 탭 구성
    tab1, tab2 = st.tabs(["📊 시세 현황판", "🤖 AI 심층 분석"])

    with tab1:
        st.header("🔍 최근 주간/월간 변동 요약")
        c1, c2 = st.columns(2)
        
        latest_week = weekly_df['기간'].max()
        latest_month = monthly_df['기간'].max()

        with c1:
            st.subheader(f"🗓️ 주간 평균 ({latest_week})")
            st.dataframe(weekly_df[weekly_df['기간'] == latest_week].style.format({'평균시세': '{:,.0f}', '증감률': '{:+.2f}%'}), hide_index=True, use_container_width=True)
        with c2:
            st.subheader(f"📅 월간 평균 ({latest_month})")
            st.dataframe(monthly_df[monthly_df['기간'] == latest_month].style.format({'평균시세': '{:,.0f}', '증감률': '{:+.2f}%'}), hide_index=True, use_container_width=True)

    with tab2:
        st.header("📝 전문가 AI 협업 리포트 생성")
        
        all_items = weekly_df['품목'].unique().tolist()
        selected_items = st.multiselect("상세 분석이 필요한 품목을 선택하세요", all_items, default=all_items[:1])

        if st.button("🚀 전문가 분석 시작 (Word 생성)"):
            if not selected_items:
                st.warning("품목을 선택해 주세요.")
            else:
                with st.spinner("AI 전문가 그룹이 시장 지표를 분석하고 보고서를 작성 중입니다..."):
                    
                    # LLM 설정 (GPT-4o)
                    gpt4o = LLM(model="gpt-4o")

                    # 에이전트 1: 시장 원인분석 전문가
                    analyst = Agent(
                        role="농축수산물 시장 변동 원인분석 전문가",
                        goal="가격 변동 원인을 공급, 수요, 외부 환경, 유통, 연관 시장의 5개 영역에서 분석",
                        backstory="15년 경력의 시장 분석가로, 기후 및 국제 공급망 데이터를 해석하는 능력이 탁월합니다.",
                        llm=gpt4o, verbose=True
                    )

                    # 에이전트 2: 구매 실무 인사이트 전문가
                    procurer = Agent(
                        role="구매담당자를 위한 시장 인사이트 전문가",
                        goal="분석된 원인을 바탕으로 'So What?', 모니터링 포인트, 대응 시점을 도출",
                        backstory="대기업 구매부서 10년 경력자로, 가격 변동에 따른 실질적인 구매 전략을 수립합니다.",
                        llm=gpt4o, verbose=True
                    )

                    # 에이전트 3: 통합 보고서 작성자
                    writer = Agent(
                        role="구매부서 보고서 전문 작성자",
                        goal="개별 품목 분석을 종합하여 전략적인 시장 보고서 완성",
                        backstory="경영진 보고용 대외비 리포트를 작성하는 전문가로, 핵심 내용을 구조화하는 데 능숙합니다.",
                        llm=gpt4o, verbose=True
                    )

                    # 품목별 분석 실행
                    item_reports = []
                    for item in selected_items:
                        st.write(f"🔎 {item} 분석 중...")
                        
                        task1 = Task(
                            description=f"품목: {item}. 5개 영역(공급, 수요, 외부, 유통, 연관)에서 가격 변동 원인을 2025-2026 최신 트렌드를 반영하여 상세히 분석하세요.",
                            agent=analyst,
                            expected_output="5대 영역별 상세 분석 마크다운"
                        )
                        
                        task2 = Task(
                            description=f"{item}의 분석 내용을 바탕으로 구매담당자를 위한 'So What?', 'Early Warning' 신호, 'When to Act' 지침을 작성하세요.",
                            agent=procurer,
                            expected_output="구매 대응 전략 및 모니터링 가이드"
                        )

                        crew = Crew(agents=[analyst, procurer], tasks=[task1, task2])
                        result = crew.kickoff()
                        item_reports.append(f"# {item} 시장 심층 분석\n\n{result.raw}")

                    # 최종 보고서 통합
                    final_task = Task(
                        description="분석된 모든 품목 리포트를 통합하고, 전체 시장의 공통 리스크와 조기 경보 시스템(EWS)을 포함한 최종 보고서를 작성하세요.",
                        agent=writer,
                        expected_output="통합 시장 분석 및 전략 보고서 (마크다운)"
                    )
                    
                    final_crew = Crew(agents=[writer], tasks=[final_task])
                    final_report_md = final_crew.kickoff(inputs={"contents": "\n\n".join(item_reports)}).raw

                    st.divider()
                    st.markdown(final_report_md)

                    # Word 다운로드 생성
                    word_doc = markdown_to_word_stream(final_report_md)
                    st.download_button(
                        label="📄 최종 분석 보고서 다운로드 (.docx)",
                        data=word_doc,
                        file_name=f"Raw_Material_Strategy_Report_{datetime.datetime.now().strftime('%Y%m%d')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

else:
    st.error("데이터 로드에 실패했습니다. 구글 시트 URL을 확인해 주세요.")
