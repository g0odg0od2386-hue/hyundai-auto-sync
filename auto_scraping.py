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

# 한국 시간 기준 어제 날짜
target_date = (datetime.now() + timedelta(hours=9) - timedelta(days=1)).strftime("%Y-%m-%d")

SPREADSHEET_NAME = "현대백화점 접수 내용"
WORKSHEET_NAME = "2026.04RawData" 

google_key_json = os.environ.get("GOOGLE_SERVICE_KEY")
# -------------------------------

options = Options()
options.add_argument("--headless") 
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
all_combined_data = [] 

try:
    print(f"🚀 {target_date} 수집 시작 (강제 대기 모드)")
    wait = WebDriverWait(driver, 30) # 최대 30초 대기

    for member in members:
        emp_id, emp_name = member["id"], member["name"]
        print(f"🔍 {emp_name}({emp_id}) 데이터 조회 시도...")
        
        # 1. 페이지를 매번 새로 불러와서 초기화
        driver.get("https://www.ehyundai.com/newPortal/card/TB/TB000018_V.do")
        time.sleep(5) # 페이지가 완전히 안착할 때까지 무조건 5초 대기

        try:
            # 2. 입력칸이 '실제로' 보일 때까지 대기 후 찾기
            # wrtRecommenNum 라는 ID를 가진 요소를 찾음
            id_box = wait.until(EC.element_to_be_clickable((By.ID, "wrtRecommenNum")))
            start_date_box = driver.find_element(By.ID, "wrtStartDate")
            end_date_box = driver.find_element(By.ID, "wrtFinishDate")

            # 3. 자바스크립트 에러를 피하기 위해 하나씩 직접 입력
            driver.execute_script("arguments[0].value = '';", id_box) # 기존값 지우기
            id_box.send_keys(emp_id)
            
            driver.execute_script(f"arguments[0].value = '{target_date}';", start_date_box)
            driver.execute_script(f"arguments[0].value = '{target_date}';", end_date_box)

            # 4. 조회 버튼 클릭 (자바스크립트로 강제 클릭)
            search_btn = driver.find_element(By.CSS_SELECTOR, "a.btn.btn_srch")
            driver.execute_script("arguments[0].click();", search_btn)
            
            time.sleep(3) # 데이터 로딩 대기

            # 5. 데이터 수집
            html_content = driver.page_source
            if '접수점' in html_content:
                df_list = pd.read_html(StringIO(html_content), flavor='lxml')
                target_df = [d for d in df_list if '접수점' in d.columns][0]
                if not target_df.empty and '데이터가 없습니다' not in html_content:
                    target_df['성명'], target_df['사번'], target_df['날짜'] = emp_name, emp_id, target_date
                    target_df = target_df[target_df['접수점'].str.contains('점|본사|더현대', na=False)]
                    all_combined_data.append(target_df)
                    print(f"✅ {emp_name} 실적 수집 성공")
                else:
                    print(f"⚪ {emp_name} 실적 없음")
            
        except Exception as sub_e:
            print(f"⚠️ {emp_name} 단계에서 오류(스킵): {sub_e}")
            continue

    # 6. 구글 시트 전송
    if all_combined_data:
        final_df = pd.concat(all_combined_data, ignore_index=True).fillna("")
        creds_dict = json.loads(google_key_json)
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        
        doc = client.open(SPREADSHEET_NAME)
        worksheet = doc.worksheet(WORKSHEET_NAME)
        worksheet.append_rows(final_df.values.tolist())
        print(f"🎉 총 {len(final_df)}건 시트 업데이트 완료!")
    else:
        print("📢 수집된 데이터가 없습니다.")

except Exception as e:
    print(f"❌ 치명적 오류: {e}")
finally:
    driver.quit()
