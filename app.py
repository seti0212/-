import streamlit as st
import datetime
import re
import io
import os
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from crewai import Agent, Task, Crew, LLM

# ============================================================================
# 1. 문서 변환 및 서식 유틸리티 (Word 파일 생성용)
# ============================================================================

def add_formatted_text(paragraph, text):
    """마크다운 서식을 Word 서식으로 변환 (굵게, 첨자 등)"""
    pattern = r'(\*\*.*?\*\*|\*.*?\*|`.*?`|\[.*?\]\(.*?\)|\[\d+(?:,\d+)*\])'
    parts = re.split(pattern, text)
    
    for part in parts:
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            paragraph.add_run(part[2:-2]).bold = True
        elif re.match(r'^\[\d+(?:,\d+)*\]$', part):
            run = paragraph.add_run(part)
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(0, 0, 255)
            run.font.superscript = True
        else:
            paragraph.add_run(part)

def create_table(doc, table_lines):
    """마크다운 표 형식을 Word 테이블로 변환"""
    rows = [line.split('|')[1:-1] for line in table_lines if '|' in line and '---' not in line]
    if not rows: return
    
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = 'Table Grid'
    
    for r_idx, row in enumerate(rows):
        for c_idx, cell_text in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            # 셀 내 텍스트 서식 적용
            add_formatted_text(cell.paragraphs[0], cell_text.strip())

def markdown_to_docx_stream(markdown_text):
    """전체 마크다운 텍스트를 Word 바이너리 스트림으로 변환"""
    doc = Document()
    
    # 문서 기본 여백 설정
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)

    lines = markdown_text.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        
        # 제목 및 본문 처리
        if line.startswith('# '):
            h = doc.add_heading(line[2:], level=0)
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif line.startswith('## '):
            doc.add_heading(line[3:], level=1)
        elif line.startswith('### '):
            doc.add_heading(line[4:], level=2)
        elif line.startswith('- ') or line.startswith('* '):
            doc.add_paragraph(line[2:], style='List Bullet')
        elif '|' in line and i + 1 < len(lines) and '|--' in lines[i+1]:
            table_lines = []
            while i < len(lines) and '|' in lines[i]:
                table_lines.append(lines[i])
                i += 1
            create_table(doc, table_lines)
            continue
        else:
            p = doc.add_paragraph()
            add_formatted_text(p, line)
        i += 1
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ============================================================================
# 2. 에이전트 및 태스크 정의 (공유해주신 상세 로직 반영)
# ============================================================================

def run_ai_analysis(items, dates):
    """CrewAI를 실행하여 품목별 분석 및 통합 보고서 생성"""
    
    # 에이전트 설정
    market_analyst = Agent(
        role="농축수산물 시장 변동 원인분석 전문가",
        goal="가격 급등락의 근본 원인을 5개 영역에서 체계적으로 분석하여 인사이트 제공",
        backstory="15년 경력의 시장 분석가. 추측을 배제하고 2025년 최신 데이터를 기반으로 인과관계를 분석합니다.",
        llm=LLM(model="gpt-4o"),
        verbose=True
    )

    procurement_expert = Agent(
        role="구매담당자를 위한 시장 인사이트 전문가",
        goal="원인분석을 바탕으로 구매 실무자가 알아야 할 핵심 인사이트와 'So What' 도출",
        backstory="대기업 구매부서 10년 경력. 리스크 신호와 구매 타이밍을 포착하는 데 전문가입니다.",
        llm=LLM(model="gpt-4o"),
        verbose=True
    )

    writer = Agent(
        role="구매부서를 위한 시장 인사이트 보고서 전문 작성자",
        goal="개별 분석을 종합하여 전체 시장의 구조적 트렌드와 조기 경보 시스템 구축",
        backstory="식품 기업 구매팀의 전략 보고서 담당자. 경영진이 한눈에 파악할 수 있는 요약 능력이 뛰어납니다.",
        llm=LLM(model="gpt-4o"),
        verbose=True
    )

    all_item_results = []

    # 1단계: 품목별 개별 분석
    for item, date in zip(items, dates):
        st.write(f"🔍 **{item}** 품목에 대한 심층 분석을 수행 중입니다...")
        
        analysis_task = Task(
            description=f"""
            품목: {item} (기준일: {date})
            다음 5개 영역에서 가격 변동 원인을 분석하세요:
            1. 공급 측면(생산량, 비용, 재고)
            2. 수요 측면(소비 패턴, 구매주체 변화)
            3. 외부 환경(정책, 무역, 기후)
            4. 유통 구조(단계별 마진, 물류비)
            5. 연관 시장(대체재, 보완재 영향)
            """,
            expected_output=f"{item} 가격 변동 원인 분석서",
            agent=market_analyst
        )

        insight_task = Task(
            description=f"위 분석을 바탕으로 {item} 구매 담당자가 주의해야 할 신호와 향후 3개월 전망을 도출하세요.",
            expected_output=f"{item} 구매 전략 가이드",
            agent=procurement_expert
        )

        item_crew = Crew(agents=[market_analyst, procurement_expert], tasks=[analysis_task, insight_task])
        item_result = item_crew.kickoff()
        all_item_results.append(f"### {item} 분석 결과\n\n{item_result.raw}")

    # 2단계: 통합 보고서 생성
    st.write("📂 모든 데이터를 종합하여 **전략 보고서**를 작성 중입니다...")
    
    integration_task = Task(
        description=f"""
        다음 개별 품목 분석 결과를 바탕으로 '구매부서 종합 분석 보고서'를 작성하세요.
        내용에는 Executive Summary, 거시적 환경 분석, 품목군별 종합 전망, 조기 경보 시스템(Early Warning)이 포함되어야 합니다.
        
        대상 데이터:
        {chr(10).join(all_item_results)}
        """,
        expected_output="종합 시장 분석 보고서 (마크다운 형식)",
        agent=writer
    )

    final_crew = Crew(agents=[writer], tasks=[integration_task])
    final_report = final_crew.kickoff()

    return final_report.raw

# ============================================================================
# 3. Streamlit 메인 화면 인터페이스
# ============================================================================

st.set_page_config(page_title="AI 마켓 리포트 생성기", layout="wide")
st.divider()
st.header("📝 구매부서 전용 종합 마켓 보고서 (AI)")

# 기존 데이터(result)가 세션에 있는지 확인
if 'result' not in st.session_state:
    st.warning("⚠️ 분석할 품목 데이터가 없습니다. 먼저 상단에서 데이터를 로드해 주세요.")
else:
    items = st.session_state.result['품목'].tolist()
    dates = st.session_state.result['마지막일'].tolist()

    if st.button("🚀 전문 AI 심층 보고서 생성"):
        with st.status("전문 분석팀이 작업을 시작합니다...", expanded=True) as status:
            
            # AI 분석 실행
            final_md_report = run_ai_analysis(items, dates)
            
            status.update(label="✅ 보고서 작성 완료!", state="complete", expanded=False)

        # 4. 결과 출력 및 다운로드
        st.markdown("---")
        st.subheader("📊 생성된 보고서 미리보기")
        st.markdown(final_md_report)

        st.divider()
        col1, col2 = st.columns(2)
        
        with col1:
            # Word 파일 생성 및 다운로드 버튼
            docx_stream = markdown_to_docx_stream(final_md_report)
            st.download_button(
                label="📄 Word 보고서 다운로드 (.docx)",
                data=docx_stream,
                file_name=f"Market_Analysis_Report_{datetime.date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        with col2:
            # 마크다운 파일 다운로드 버튼
            st.download_button(
                label="📝 마크다운 파일 저장 (.md)",
                data=final_md_report,
                file_name=f"Market_Analysis_Report_{datetime.date.today()}.md",
                mime="text/markdown"
            )
