import time
import pandas as pd
import os
import json
from io import StringIO
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
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

# 실행일 기준 어제 날짜 수집 (자동화용)
target_date = (datetime.now() + timedelta(hours=9) - timedelta(days=1)).strftime("%Y-%m-%d")

SPREADSHEET_NAME = "현대백화점 접수 내용"
WORKSHEET_NAME = "(4월)단순계산기"

# GitHub Secrets에서 보안키 불러오기
google_key_json = os.environ.get("GOOGLE_SERVICE_KEY")
# -------------------------------

options = Options()
options.add_argument("--headless") # 서버용 (창 안 띄움)
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
all_combined_data = [] 

try:
    print(f"🚀 {target_date} 데이터 수집 시작")
    driver.get("https://www.ehyundai.com/newPortal/card/TB/TB000018_V.do")
    time.sleep(2)

    for member in members:
        emp_id, emp_name = member["id"], member["name"]
        driver.execute_script(f"""
            document.getElementById('wrtRecommenNum').value = '{emp_id}';
            document.getElementById('wrtStartDate').value = '{target_date}';
            document.getElementById('wrtFinishDate').value = '{target_date}';
            $('#Form').submit();
        """)
        time.sleep(1)

        html_content = driver.page_source
        if '접수점' in html_content:
            df_list = pd.read_html(StringIO(html_content), flavor='lxml')
            target_df = [d for d in df_list if '접수점' in d.columns][0]
            if not target_df.empty:
                target_df['성명'], target_df['사번'], target_df['날짜'] = emp_name, emp_id, target_date
                target_df = target_df[target_df['접수점'].str.contains('점|본사|더현대', na=False)]
                all_combined_data.append(target_df)

    if all_combined_data:
        final_df = pd.concat(all_combined_data, ignore_index=True).fillna("")
        
        # 구글 시트 업로드
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds_dict = json.loads(google_key_json)
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        
        worksheet = client.open(SPREADSHEET_NAME).worksheet(WORKSHEET_NAME)
        # 마지막 행 다음에 추가
        worksheet.append_rows(final_df.values.tolist())
        print(f"✅ {target_date} 실적 업데이트 완료!")

except Exception as e:
    print(f"❌ 오류: {e}")
finally:
    driver.quit()