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
# 💡 핵심: 서버 환경을 일반 PC처럼 보이게 만드는 속성
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
options.add_argument("lang=ko_KR")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
all_combined_data = [] 

try:
    print(f"🚀 {target_date} 수집 시작 (서버 최적화 모드)")
    driver.get("https://www.ehyundai.com/newPortal/card/TB/TB000018_V.do")
    
    # 1. 페이지가 완전히 로딩될 때까지 넉넉히 대기
    wait = WebDriverWait(driver, 30)

    for member in members:
        emp_id, emp_name = member["id"], member["name"]
        print(f"🔍 {emp_name}({emp_id}) 데이터 조회 시도...")
        
        try:
            # 2. 입력칸이 나타날 때까지 대기 (ID가 없으면 NAME으로라도 찾음)
            wait.until(EC.presence_of_element_located((By.XPATH, "//*[@name='wrtRecommenNum' or @id='wrtRecommenNum']")))
            
            # 3. 안전한 JavaScript 입력 방식 (서버 에러 방지용)
            driver.execute_script(f"""
                var idInput = document.getElementsByName('wrtRecommenNum')[0] || document.getElementById('wrtRecommenNum');
                var startInput = document.getElementsByName('wrtStartDate')[0] || document.getElementById('wrtStartDate');
                var endInput = document.getElementsByName('wrtFinishDate')[0] || document.getElementById('wrtFinishDate');
                
                if(idInput) idInput.value = '{emp_id}';
                if(startInput) startInput.value = '{target_date}';
                if(endInput) endInput.value = '{target_date}';
            """)
            
            # 4. 조회 버튼 클릭
            search_btn = driver.find_element(By.XPATH, "//a[contains(text(), '조회')] | //a[contains(@class, 'btn_srch')]")
            driver.execute_script("arguments[0].click();", search_btn)
            
            time.sleep(4) # 데이터 로딩 대기

            # 5. 데이터 긁기
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
            print(f"⚠️ {emp_name} 단계에서 지연 혹은 오류 발생: {sub_e}")
            continue

    # 6. 구글 시트 업로드
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
        print("📢 오늘은 수집할 새로운 실적이 없습니다.")

except Exception as e:
    print(f"❌ 치명적 오류: {e}")
finally:
    driver.quit()
