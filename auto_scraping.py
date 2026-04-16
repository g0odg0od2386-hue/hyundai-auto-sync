
import time
import pandas as pd
import os
from io import StringIO
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# --- [사번 및 이름 리스트 설정] ---
members = [
    {"id": "H0196287", "name": "남광일"},
    {"id": "H0300607", "name": "정진성"},
    {"id": "H0300617", "name": "김문희"},
    {"id": "H0300682", "name": "손영주"}
]

start_date = datetime(2026, 4, 1)
end_date = datetime(2026, 4, 20)
save_path = r"C:\Users\J"
# -------------------------------

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
options.add_argument("--start-maximized")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
all_combined_data = [] 
date_list = [start_date + timedelta(days=x) for x in range((end_date - start_date).days + 1)]

try:
    print(f"🌌 4인 통합 수집 모드 시작 (총 인원: {len(members)}명)")
    driver.get("https://www.ehyundai.com/newPortal/card/TB/TB000018_V.do")
    time.sleep(2)

    for member in members:
        emp_id = member["id"]
        emp_name = member["name"]
        print(f"\n👤 {emp_name}({emp_id}) 데이터 낚는 중...", end=" ", flush=True)

        for current_date in date_list:
            date_str = current_date.strftime("%Y-%m-%d")
            
            # JS로 조회 명령 전송 (0.2초 리듬)
            driver.execute_script(f"""
                document.getElementById('wrtRecommenNum').value = '{emp_id}';
                var s = document.getElementById('wrtStartDate');
                var e = document.getElementById('wrtFinishDate');
                s.value = '{date_str}'; e.value = '{date_str}';
                $('#Form').submit();
            """)
            
            time.sleep(0.2) 

            try:
                html_content = driver.page_source
                if '접수점' in html_content:
                    df_list = pd.read_html(StringIO(html_content), flavor='lxml')
                    target_df = [d for d in df_list if '접수점' in d.columns][0]
                    
                    if not target_df.empty:
                        # 로우 데이터 레이블링
                        target_df['성명'] = emp_name
                        target_df['사번'] = emp_id
                        target_df['날짜'] = date_str
                        # 실적 행만 필터링 (점포 데이터)
                        target_df = target_df[target_df['접수점'].str.contains('점|본사|더현대', na=False)]
                        all_combined_data.append(target_df)
                        print(f"🔥", end="", flush=True)
                else:
                    print("⚠️", end="", flush=True)
            except:
                print("❌", end="", flush=True)

    # 엑셀 저장 단계
    if all_combined_data:
        final_df = pd.concat(all_combined_data, ignore_index=True)
        
        # 열 순서 예쁘게 배치 (날짜, 성명, 사번 순)
        cols = ['날짜', '성명', '사번'] + [c for c in final_df.columns if c not in ['날짜', '성명', '사번']]
        final_df = final_df[cols]

        timestamp = datetime.now().strftime("%H%M%S")
        file_name = f"현대_4월실적_4인통합_{timestamp}.xlsx"
        full_path = os.path.join(save_path, file_name)
        
        final_df.to_excel(full_path, index=False)
        print(f"\n\n🏁 전체 수집 완료! (총 {len(final_df)}개 행 수집)")
        print(f"📂 저장 위치: {full_path}")
    else:
        print("\n⚠️ 수집된 데이터가 없습니다.")

except Exception as e:
    print(f"\n☢️ 시스템 과열: {e}")

finally:
    print("\n" + "="*50)
    print("모든 작업이 끝났습니다. 통합 엑셀 파일을 확인해 보세요!")
    print("="*50)
    while True: time.sleep(1)
