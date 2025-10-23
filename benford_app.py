
"""
벤포드 분석 웹 애플리케이션 (Streamlit)
"""

import streamlit as st
import pandas as pd
import numpy as np
from scipy import stats
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
import io
from datetime import datetime

# 페이지 설정
st.set_page_config(
    page_title="벤포드 분석 도구",
    page_icon="🔍",
    layout="wide"
)


class BenfordAnalyzer:
    """벤포드 분석 클래스"""
    
    def __init__(self, file_data, confidence=0.95):
        self.file_data = file_data
        self.confidence = confidence
        self.z_critical = stats.norm.ppf((1 + confidence) / 2)
        self.sigma_thresholds = [1.0, 1.5, 2.0]
        
        self.wb = None
        self.data = None
        self.keywords = None
        self.results = {}
        
    def benford_probability(self, digit):
        """벤포드 법칙 확률 계산"""
        return np.log10(1 + 1 / digit)
    
    def leading_two_digits(self, number):
        """숫자의 앞 두자리 추출"""
        number = abs(number)
        if number == 0:
            return None
        
        while number >= 100:
            number /= 10
        while number < 10:
            number *= 10
        
        return int(number)
    
    def load_data(self):
        """엑셀 데이터 로딩"""
        self.wb = load_workbook(self.file_data)
        
        # 분개장 데이터
        df = pd.read_excel(self.file_data, sheet_name='분개장입력')
        
        # 금액 숫자 변환
        df['금액'] = pd.to_numeric(df['금액'], errors='coerce')
        df = df.dropna(subset=['금액'])
        
        # 앞 두자리 추출
        df['두자리수'] = df['금액'].apply(self.leading_two_digits)
        df = df[(df['두자리수'] >= 10) & (df['두자리수'] <= 99)]
        
        self.data = df
        
        # 키워드 로딩
        try:
            keywords_df = pd.read_excel(self.file_data, sheet_name='적요 키워드')
            self.keywords = keywords_df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        except:
            self.keywords = []
    
    def analyze_benford(self, data, label="전체"):
        """벤포드 분석 수행"""
        n = len(data)
        if n == 0:
            return None
        
        results = []
        
        for digit in range(10, 100):
            observed = (data['두자리수'] == digit).sum()
            expected_prob = self.benford_probability(digit)
            expected_count = expected_prob * n
            
            obs_ratio = observed / n if n > 0 else 0
            std_dev = np.sqrt(n * expected_prob * (1 - expected_prob))
            z_value = (observed - expected_count) / std_dev if std_dev > 0 else 0
            chi_sq = (observed - expected_count) ** 2 / expected_count if expected_count > 0 else 0
            
            flag = "초과" if abs(z_value) >= self.z_critical else ""
            
            results.append({
                '구분': label,
                '두자리수': digit,
                '관측빈도': observed,
                '관측비율': obs_ratio,
                '기대비율': expected_prob,
                '기대빈도': expected_count,
                '(O-E)^2/E': chi_sq,
                '표본수': n,
                '표준편차': std_dev,
                'Z값': z_value,
                '신뢰수준판정': flag
            })
        
        return pd.DataFrame(results)
    
    def run_analysis(self):
        """전체 분석 실행"""
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # 1. 전체 분석
        status_text.text("1/8 전체 데이터 분석 중...")
        progress_bar.progress(12)
        self.results['01_벤포드_전체'] = self.analyze_benford(self.data, "전체")
        
        # 2. 손익별
        if '손익' in self.data.columns:
            status_text.text("2/8 손익별 분석 중...")
            progress_bar.progress(25)
            dfs = []
            for pnl in self.data['손익'].unique():
                if pd.notna(pnl):
                    df = self.analyze_benford(self.data[self.data['손익'] == pnl], pnl)
                    if df is not None:
                        dfs.append(df)
            if dfs:
                self.results['02_벤포드_손익'] = pd.concat(dfs, ignore_index=True)
        
        # 3. 대분류별
        if '대분류' in self.data.columns:
            status_text.text("3/8 대분류별 분석 중...")
            progress_bar.progress(37)
            dfs = []
            for cat in self.data['대분류'].unique():
                if pd.notna(cat):
                    df = self.analyze_benford(self.data[self.data['대분류'] == cat], cat)
                    if df is not None:
                        dfs.append(df)
            if dfs:
                self.results['03_벤포드_대분류'] = pd.concat(dfs, ignore_index=True)
        
        # 4. 중분류별
        if '중분류' in self.data.columns:
            status_text.text("4/8 중분류별 분석 중...")
            progress_bar.progress(50)
            dfs = []
            for cat in self.data['중분류'].unique():
                if pd.notna(cat):
                    df = self.analyze_benford(self.data[self.data['중분류'] == cat], cat)
                    if df is not None:
                        dfs.append(df)
            if dfs:
                self.results['04_벤포드_중분류'] = pd.concat(dfs, ignore_index=True)
        
        # 5. 소분류별
        if '소분류' in self.data.columns:
            status_text.text("5/8 소분류별 분석 중...")
            progress_bar.progress(62)
            dfs = []
            for cat in self.data['소분류'].unique():
                if pd.notna(cat):
                    df = self.analyze_benford(self.data[self.data['소분류'] == cat], cat)
                    if df is not None:
                        dfs.append(df)
            if dfs:
                self.results['05_벤포드_소분류'] = pd.concat(dfs, ignore_index=True)
        
        # 6. 일자별 요약
        status_text.text("6/8 일자별 요약 중...")
        progress_bar.progress(75)
        daily = self.data.groupby(['일자', '손익']).agg({
            '금액': ['count', 'sum']
        }).reset_index()
        daily.columns = ['일자', '손익', '전표갯수', '금액합계']
        self.results['06_일자요약'] = daily
        
        # 7. 거래처별 요약
        status_text.text("7/8 거래처별 요약 중...")
        progress_bar.progress(87)
        vendor = self.data.groupby(['거래처', '손익']).agg({
            '금액': ['count', 'sum']
        }).reset_index()
        vendor.columns = ['거래처', '손익', '전표갯수', '금액합계']
        self.results['07_거래처요약'] = vendor
        
        # 8. 키워드 필터링
        if self.keywords:
            status_text.text("8/8 키워드 필터링 중...")
            progress_bar.progress(100)
            keyword_data = []
            for keyword in self.keywords:
                mask = self.data['적요'].astype(str).str.contains(keyword, case=False, na=False)
                matched = self.data[mask].copy()
                if len(matched) > 0:
                    matched.insert(0, '키워드', keyword)
                    keyword_data.append(matched)
            
            if keyword_data:
                cols = ['키워드', '일자', '손익', '금액', '거래처', '적요']
                if '전표번호' in self.data.columns:
                    cols.insert(2, '전표번호')
                if '대분류' in self.data.columns:
                    cols.insert(-3, '대분류')
                if '중분류' in self.data.columns:
                    cols.insert(-3, '중분류')
                if '소분류' in self.data.columns:
                    cols.insert(-3, '소분류')
                
                available_cols = [col for col in cols if col in pd.concat(keyword_data).columns]
                self.results['08_키워드필터'] = pd.concat(keyword_data, ignore_index=True)[available_cols]
        
        # 9. 중점검토 시나리오
        self.results['09_중점검토'] = self.create_scenario_sheet()
        
        # 10. 분석 해설
        self.results['10_분석해설'] = self.create_explanation_sheet()
        
        progress_bar.progress(100)
        status_text.text("✅ 분석 완료!")
    
    def create_scenario_sheet(self):
        """중점검토 시나리오 생성"""
        top_findings = []
        
        for key, df in self.results.items():
            if key.startswith('0') and '벤포드' in key:
                if 'Z값' in df.columns:
                    df_sorted = df.nlargest(15, 'Z값', keep='all')
                    df_sorted['범위'] = key.replace('_벤포드_', ' - ').replace('0', '').replace('_', '')
                    top_findings.append(df_sorted)
        
        if top_findings:
            all_findings = pd.concat(top_findings, ignore_index=True)
            all_findings['Z값_절대'] = all_findings['Z값'].abs()
            top15 = all_findings.nlargest(15, 'Z값_절대').copy()
            
            cols = ['범위', '구분', '두자리수', '관측빈도', '기대빈도', '표본수', 'Z값', '신뢰수준판정']
            return top15[[col for col in cols if col in top15.columns]]
        
        return None
    
    def create_explanation_sheet(self):
        """분석 해설 시트 생성"""
        explanations = []
        
        # 전체 분석 결과
        if '01_벤포드_전체' in self.results:
            df = self.results['01_벤포드_전체']
            
            # Z값 기준 심각도
            red_count = len(df[df['Z값'].abs() >= 2.0])
            orange_count = len(df[(df['Z값'].abs() >= 1.5) & (df['Z값'].abs() < 2.0)])
            yellow_count = len(df[(df['Z값'].abs() >= 1.0) & (df['Z값'].abs() < 1.5)])
            
            severity_score = red_count * 3 + orange_count * 2 + yellow_count * 1
            
            if severity_score > 100:
                risk = "🔴 높음"
                action = "즉각 조사 필요"
            elif severity_score > 50:
                risk = "🟠 중간"
                action = "주의 깊은 검토 필요"
            else:
                risk = "🟡 낮음"
                action = "일반 모니터링"
            
            explanations.append({
                '항목': '종합 위험도',
                '값': f'{severity_score}/270',
                '평가': risk,
                '조치': action
            })
            
            explanations.append({
                '항목': '매우 이상 (Z≥2)',
                '값': f'{red_count}개',
                '평가': '🔴',
                '조치': '전수 조사'
            })
            
            explanations.append({
                '항목': '주의 (Z≥1.5)',
                '값': f'{orange_count}개',
                '평가': '🟠',
                '조치': '샘플 조사'
            })
            
            explanations.append({
                '항목': '관심 (Z≥1)',
                '값': f'{yellow_count}개',
                '평가': '🟡',
                '조치': '모니터링'
            })
            
            # Top 5 이상
            top5 = df.nlargest(5, 'Z값')
            explanations.append({
                '항목': '',
                '값': '',
                '평가': '',
                '조치': ''
            })
            explanations.append({
                '항목': '가장 이상한 5개',
                '값': '',
                '평가': '',
                '조치': ''
            })
            
            for i, (_, row) in enumerate(top5.iterrows(), 1):
                digit = int(row['두자리수'])
                obs = int(row['관측빈도'])
                exp = row['기대빈도']
                z = row['Z값']
                diff = obs - exp
                
                explanations.append({
                    '항목': f'{i}. {digit}으로 시작',
                    '값': f'{obs}건 (정상:{exp:.0f}건)',
                    '평가': f'Z={z:.1f}',
                    '조치': f'+{diff:.0f}건 초과'
                })
        
        # 데이터 통계
        explanations.append({
            '항목': '',
            '값': '',
            '평가': '',
            '조치': ''
        })
        explanations.append({
            '항목': '데이터 통계',
            '값': '',
            '평가': '',
            '조치': ''
        })
        
        total_entries = len(self.data)
        explanations.append({
            '항목': '총 전표 수',
            '값': f'{total_entries:,}개',
            '평가': '',
            '조치': ''
        })
        
        if '일자' in self.data.columns:
            unique_dates = self.data['일자'].nunique()
            explanations.append({
                '항목': '분석 기간',
                '값': f'{unique_dates}일',
                '평가': '',
                '조치': f'일평균 {total_entries/unique_dates:.0f}건'
            })
        
        if '거래처' in self.data.columns:
            unique_vendors = self.data['거래처'].nunique()
            explanations.append({
                '항목': '총 거래처 수',
                '값': f'{unique_vendors:,}개',
                '평가': '',
                '조치': ''
            })
        
        return pd.DataFrame(explanations)
    
    def apply_conditional_formatting(self, ws, start_row, end_row, z_col_idx):
        """조건부 서식 적용"""
        colors = {
            'red': PatternFill(start_color='FF9999', end_color='FF9999', fill_type='solid'),
            'orange': PatternFill(start_color='FFCC66', end_color='FFCC66', fill_type='solid'),
            'yellow': PatternFill(start_color='FFFF99', end_color='FFFF99', fill_type='solid')
        }
        
        for row_idx in range(start_row, end_row + 1):
            z_cell = ws.cell(row_idx, z_col_idx)
            if z_cell.value is not None:
                try:
                    z_val = abs(float(z_cell.value))
                    
                    if z_val >= 2.0:
                        fill = colors['red']
                    elif z_val >= 1.5:
                        fill = colors['orange']
                    elif z_val >= 1.0:
                        fill = colors['yellow']
                    else:
                        continue
                    
                    # 전체 행에 색상 적용
                    for col_idx in range(1, ws.max_column + 1):
                        ws.cell(row_idx, col_idx).fill = fill
                except:
                    pass
    
    def save_results(self):
        """결과를 엑셀로 저장"""
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # 원본 데이터 먼저 (두자리수 컬럼 제거)
            original_data = self.data.drop(columns=['두자리수'], errors='ignore')
            original_data.to_excel(writer, sheet_name='00_원본데이터', index=False)
            
            # 분석 결과들
            for sheet_name, df in self.results.items():
                if df is not None and len(df) > 0:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # 조건부 서식 적용
        output.seek(0)
        wb = load_workbook(output)
        
        # 벤포드 시트들에 색상 적용
        for sheet_name in wb.sheetnames:
            if '벤포드' in sheet_name:
                ws = wb[sheet_name]
                if ws.max_row > 1:
                    # Z값 컬럼 찾기
                    z_col_idx = None
                    for col_idx in range(1, ws.max_column + 1):
                        if ws.cell(1, col_idx).value == 'Z값':
                            z_col_idx = col_idx
                            break
                    
                    if z_col_idx:
                        self.apply_conditional_formatting(ws, 2, ws.max_row, z_col_idx)
            
            # 중점검토 시트도
            if sheet_name == '09_중점검토':
                ws = wb[sheet_name]
                if ws.max_row > 1:
                    for col_idx in range(1, ws.max_column + 1):
                        if ws.cell(1, col_idx).value == 'Z값':
                            self.apply_conditional_formatting(ws, 2, ws.max_row, col_idx)
                            break
        
        # 메모리에 저장
        final_output = io.BytesIO()
        wb.save(final_output)
        final_output.seek(0)
        
        return final_output


def main():
    """메인 앱"""
    
    # 제목
    st.title("🔍 벤포드 분석 도구")
    st.markdown("---")
    
    # 사이드바 - 설명
    with st.sidebar:
        st.header("📖 사용 방법")
        st.markdown("""
        ### 1️⃣ 파일 준비
        - '벤포드_분개장_틀.xlsx'에 데이터 입력
        - 필수 컬럼: 일자, 손익, 금액, 거래처, 적요
        
        ### 2️⃣ 파일 업로드
        - 아래에서 파일 선택
        
        ### 3️⃣ 분석 실행
        - 자동으로 분석 시작
        
        ### 4️⃣ 결과 다운로드
        - '벤포드 분석결과.xlsx' 다운로드
        """)
        
        st.markdown("---")
        st.markdown("### 🎯 Z값 의미")
        st.markdown("""
        - **|Z| < 1.0**: 정상 범위
        - **|Z| ≥ 1.0**: 🟡 약간 주의
        - **|Z| ≥ 1.5**: 🟠 주의 필요
        - **|Z| ≥ 2.0**: 🔴 조사 필요
        - **|Z| ≥ 3.0**: 🔴🔴 즉각 조사
        """)
    
    # 메인 영역
    st.header("📁 파일 업로드")
    
    uploaded_file = st.file_uploader(
        "벤포드 분개장 데이터 파일을 선택하세요 (.xlsx)",
        type=['xlsx'],
        help="'벤포드_분개장_틀.xlsx'에 데이터를 입력한 파일"
    )
    
    if uploaded_file is not None:
        try:
            # 파일 정보
            st.success(f"✅ 파일 업로드 완료: {uploaded_file.name}")
            
            # 미리보기
            with st.expander("📋 데이터 미리보기"):
                df_preview = pd.read_excel(uploaded_file, sheet_name='분개장입력', nrows=10)
                st.dataframe(df_preview)
                st.info(f"총 {len(pd.read_excel(uploaded_file, sheet_name='분개장입력')):,}행")
            
            # 분석 버튼
            if st.button("🚀 벤포드 분석 시작", type="primary"):
                st.markdown("---")
                st.header("⚙️ 분석 중...")
                
                # 분석 실행
                analyzer = BenfordAnalyzer(uploaded_file)
                
                with st.spinner("데이터 로딩 중..."):
                    analyzer.load_data()
                
                st.success(f"✅ 데이터 로딩 완료: {len(analyzer.data):,}건")
                
                # 분석 실행
                analyzer.run_analysis()
                
                # 결과 저장
                with st.spinner("결과 파일 생성 중..."):
                    result_file = analyzer.save_results()
                
                st.markdown("---")
                st.header("✅ 분석 완료!")
                
                # 간단한 요약
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("총 전표 수", f"{len(analyzer.data):,}건")
                
                with col2:
                    if '일자' in analyzer.data.columns:
                        days = analyzer.data['일자'].nunique()
                        st.metric("분석 기간", f"{days}일")
                
                with col3:
                    if '거래처' in analyzer.data.columns:
                        vendors = analyzer.data['거래처'].nunique()
                        st.metric("거래처 수", f"{vendors:,}개")
                
                # 주요 발견사항
                if '01_벤포드_전체' in analyzer.results:
                    df = analyzer.results['01_벤포드_전체']
                    red_count = len(df[df['Z값'].abs() >= 2.0])
                    
                    if red_count > 10:
                        st.warning(f"⚠️ 매우 이상한 패턴 {red_count}개 발견! 상세 결과 확인 필요")
                    elif red_count > 0:
                        st.info(f"ℹ️ 주의가 필요한 패턴 {red_count}개 발견")
                    else:
                        st.success("✅ 심각한 이상 패턴 없음")
                
                # 다운로드 버튼
                st.download_button(
                    label="📥 벤포드 분석결과.xlsx 다운로드",
                    data=result_file,
                    file_name=f"벤포드_분석결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                st.markdown("---")
                st.info("""
                💡 **다음 단계:**
                1. 다운로드한 파일을 엑셀에서 열기
                2. '10_분석해설' 시트에서 종합 평가 확인
                3. '09_중점검토' 시트에서 이상치 Top 15 확인
                4. 🔴빨강, 🟠주황 표시된 항목 우선 조사
                """)
                
        except Exception as e:
            st.error(f"❌ 오류 발생: {str(e)}")
            st.exception(e)
    
    else:
        # 안내 메시지
        st.info("""
        👆 위에서 파일을 선택하세요.
        
        **파일 준비 방법:**
        1. '벤포드_분개장_틀.xlsx' 다운로드
        2. '분개장입력' 시트에 데이터 입력
        3. 여기에 업로드
        """)


if __name__ == "__main__":
    main()
