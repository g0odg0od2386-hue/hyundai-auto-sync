import time
import pandas as pd
import os
import json
from io import StringIO
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- [설정 구간] ---
members = [
    {"id": "H0196287", "name": "남광일"},
    {"id": "H0300607", "name": "정진성"},
    {"id": "H0300617", "name": "김문희"},
    {"id": "H0300682", "name": "손영주"}
]

# 실행일 기준 어제 날짜 수집 (새벽 실행 고려 한국시간 기준)
target_date = (datetime.now() + timedelta(hours=9) - timedelta(days=1)).strftime("%Y-%m-%d")

# ⚠️ 주의: 시트 이름과 탭 이름을 정확히 확인하세요!
SPREADSHEET_NAME = "현대백화점 접수 내용"
WORKSHEET_NAME = "2026.04RawData"  # 아까 사진에서 확인한 탭 이름으로 수정했습니다.

google_key_json = os.environ.get("GOOGLE_SERVICE_KEY")
# -------------------------------

options = Options()
options.add_argument("--headless") 
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
all_combined_data = [] 

try:
    print(f"🚀 {target_date} 데이터 수집 시작")
    driver.get("https://www.ehyundai.com/newPortal/card/TB/TB000018_V.do")
    
    # 화면이 뜰 때까지 최대 15초 대기
    wait = WebDriverWait(driver, 15)

    for member in members:
        emp_id, emp_name = member["id"], member["name"]
        print(f"🔍 {emp_name}({emp_id}) 조회 중...")
        
        # 입력칸이 나타날 때까지 기다림 (에러 방지 핵심!)
        wait.until(EC.presence_of_element_located((By.ID, "wrtRecommenNum")))
        
        driver.execute_script(f"""
            document.getElementById('wrtRecommenNum').value = '{emp_id}';
            document.getElementById('wrtStartDate').value = '{target_date}';
            document.getElementById('wrtFinishDate').value = '{target_date}';
            $('#Form').submit();
        """)
        
        time.sleep(3) # 서버 응답 대기 시간 증가

        html_content = driver.page_source
        if '접수점' in html_content:
            df_list = pd.read_html(StringIO(html_content), flavor='lxml')
            target_df = [d for d in df_list if '접수점' in d.columns][0]
            if not target_df.empty and '데이터가 없습니다' not in html_content:
                target_df['성명'], target_df['사번'], target_df['날짜'] = emp_name, emp_id, target_date
                # '합계' 행 제외 및 유효 지점 필터링
                target_df = target_df[target_df['접수점'].str.contains('점|본사|더현대', na=False)]
                all_combined_data.append(target_df)
                print(f"✅ {emp_name} 실적 발견!")
            else:
                print(f"⚪ {emp_name} 실적 없음")

    if all_combined_data:
        final_df = pd.concat(all_combined_data, ignore_index=True).fillna("")
        
        # 구글 시트 업로드
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = json.loads(google_key_json)
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        
        doc = client.open(SPREADSHEET_NAME)
        worksheet = doc.worksheet(WORKSHEET_NAME)
        
        # 데이터 업로드
        worksheet.append_rows(final_df.values.tolist())
        print(f"🎉 {target_date} 총 {len(final_df)}건 업데이트 완료!")
    else:
        print(f"📢 {target_date} 수집된 실적이 없습니다.")

except Exception as e:
    print(f"❌ 오류 발생: {e}")
finally:
    driver.quit()
