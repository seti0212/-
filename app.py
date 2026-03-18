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
        # [핵심 수정] 특정 품목명을 가독성 좋게 변경
        data['품목'] = data['품목'].replace('가수분해소고기농축물(호주)', '호주산 쇠고기 (가수분해농축물)')
        
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
# 2. 통계 계산 및 핵심 이슈 품목 추출 (주/월/연 각 2개씩 선정)
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
    w_top = w_df[w_df['기간'] == w_df['기간'].max()].nlargest(1, '증감률')['품목'].tolist()
    w_bot = w_df[w_df['기간'] == w_df['기간'].max()].nsmallest(1, '증감률')['품목'].tolist()
    m_top = m_df[m_df['기간'] == m_df['기간'].max()].nlargest(1, '증감률')['품목'].tolist()
    m_bot = m_df[m_df['기간'] == m_df['기간'].max()].nsmallest(1, '증감률')['품목'].tolist()
    y_top = y_df[y_df['기간'] == y_df['기간'].max()].nlargest(1, '증감률')['품목'].tolist()
    y_bot = y_df[y_df['기간'] == y_df['기간'].max()].nsmallest(1, '증감률')['품목'].tolist()
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
    # 4. 전문 AI 분석 섹션 (미래 예측 기능 유지)
    # ============================================================================
    st.header("🔮 미래 단가 예측 전문 보고서 (Gemini 2.5 Flash)")
    critical_items = get_critical_items(weekly_df, monthly_df, yearly_df)
    
    today_str = datetime.date.today().strftime('%Y년 %m월 %d일')
    future_target = (datetime.date.today() + datetime.timedelta(days=120)).strftime('%Y년 %m월')

    st.write(f"🔔 **AI 미래 예측 대상:** {', '.join(critical_items)}")
    st.caption(f"※ {today_str} 기준 정보를 바탕으로 **{future_target}까지의 미래 시세**를 예측합니다.")

    if st.button("🚀 미래 단가 예측 시작"):
        if not os.environ.get("GEMINI_API_KEY"):
            st.error("🚨 API 키가 설정되지 않았습니다.")
        else:
            search_tool = SerperDevTool()
            gemini_llm = LLM(model="gemini/gemini-2.5-flash", api_key=os.environ["GEMINI_API_KEY"])

            with st.status("미래 시장 시나리오 분석 중...", expanded=True) as status:
                analyst = Agent(
                    role="미래 시장 수급 예측가", 
                    goal=f"오늘({today_str}) 이후의 뉴스 및 기후 데이터를 분석하여 향후 3~6개월의 단가를 예측", 
                    backstory="당신은 과거의 수치보다 미래의 변동 가능성에 집중합니다. 현재 발생한 사건이 미래의 어느 시점에 가격으로 반영될지를 정확히 짚어냅니다.", 
                    llm=gemini_llm, 
                    tools=[search_tool]
                )
                procurement = Agent(
                    role="전략적 미래 구매 설계자", 
                    goal="예측된 미래 단가에 따른 최적의 선매수 및 재고 확보 시점 제안", 
                    backstory="당신은 미래 단가 상승이 예상될 때 지금 바로 사야 할 양을 정하고, 하락이 예상될 때 구매를 늦추는 타이밍의 대가입니다.", 
                    llm=gemini_llm
                )

                all_reports = []
                progress_bar = st.progress(0)
                
                for idx, item in enumerate(critical_items):
                    st.write(f"🔮 **{item}** 미래 전망 분석 중... ({idx+1}/{len(critical_items)})")
                    
                    t1 = Task(
                        description=f"""
                        품목: {item}
                        현재 날짜: {today_str}
                        미션: 
                        1. 오늘 이후의 최신 뉴스(기후, 전쟁, 정책 등)를 검색하여 {item}의 미래 시세를 분석하세요.
                        2. **[미래 시나리오]** 섹션을 만들어 향후 단가 흐름을 월별 혹은 분기별로 예측하세요.
                        3. 과거 데이터 요약은 최소화하고, '앞으로 일어날 일'과 그로 인한 '미래 가격 범위'를 예측값으로 제시하세요.
                        """, 
                        expected_output=f"{item}의 미래 단가 변동 시나리오 및 예측 시점 보고서", 
                        agent=analyst
                    )
                    
                    t2 = Task(
                        description=f"""
                        위의 {item} 미래 예측 결과를 바탕으로 구매 전략을 수립하세요.
                        - **구매 적기**: 구체적인 미래 시점(예: 2026년 8월 등)을 제안하세요.
                        - **재고 전략**: 미래 가격 변동폭에 따른 물량 확보 비중을 제안하세요.
                        - **위기 알림**: 예측이 빗나갈 수 있는 변수를 지정하세요.
                        """, 
                        expected_output=f"{item}의 미래 대비 구매 실행 가이드", 
                        agent=procurement
                    )
                    
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
                                time.sleep(15 * (attempt + 1))
                            else:
                                break
                    
                    if not success:
                        all_reports.append(f"### {item}\n분석 한도 초과로 미래 리포트 생성에 실패했습니다.")
                    
                    time.sleep(7)
                    progress_bar.progress((idx + 1) / len(critical_items))

                final_report_md = f"# 📑 [전략보고] 미래 단가 예측 및 구매 로드맵 ({today_str} 발행)\n\n" + "\n\n---\n\n".join(all_reports)
                status.update(label="✅ 모든 핵심 품목 미래 예측 완료!", state="complete", expanded=False)

            st.markdown(final_report_md)
            docx_file = markdown_to_docx_stream(final_report_md)
            st.download_button(label="📄 미래 예측 보고서 다운로드 (Word)", data=docx_file, file_name=f"Future_Prediction_Report_{datetime.date.today()}.docx")
