import time
import datetime
import random
import requests
import os
import ctypes
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

CONFIG_FILE = "config.txt"

def load_config():
    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            f.write("帳號=你的帳號\n密碼=你的密碼\n測試模式=是\n隱藏視窗=否\n")
        ctypes.windll.user32.MessageBoxW(0, f"已自動建立 {CONFIG_FILE}，請填入帳密後再執行。", "初始化", 0x40)
        os._exit(0)
    
    config = {}
    with open(CONFIG_FILE, "r", encoding="utf-8-sig") as f:
        for line in f:
            if "=" in line:
                k, v = line.split("=", 1)
                config[k.strip()] = v.strip().replace("\n", "").replace("\r", "").replace("\t", "")
    return config

def is_taiwan_workday(date_to_check):
    """檢查今天是否為工作日"""
    y, m, d = date_to_check.year, date_to_check.month, date_to_check.day
    api_url = f"https://api.pin-yi.me/taiwan-calendar/{y}/{m}/{d}"
    try:
        response = requests.get(api_url, timeout=5)
        if response.status_code == 200:
            data = response.json()
            day_info = data[0] if isinstance(data, list) else data
            return not day_info.get('isHoliday', True)
        return date_to_check.weekday() < 5
    except:
        return date_to_check.weekday() < 5

def log_and_notify(msg, is_error=False):
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    content = f"[{timestamp}] {msg}"
    print(content)
    with open("punch_history.txt", "a", encoding="utf-8") as f:
        f.write(content + "\n")
    if "成功" in msg or "失敗" in msg:
        ctypes.windll.user32.MessageBoxW(0, msg, "ERP 打卡幫手", 0x10 if is_error else 0x40)

def run_punch():
    cfg = load_config()
    USER_ID = cfg.get("帳號", "")
    PASSWORD = cfg.get("密碼", "")
    TEST_MODE = True if cfg.get("測試模式") == "是" else False
    HEADLESS_MODE = True if cfg.get("隱藏視窗") == "是" else False
    ERP_URL = "https://zc8662-login.aoacloud.com.tw/Home/DeskAuthIndex"

    now = datetime.datetime.now()
    today = datetime.date.today()

    # --- 1. 假日判斷 ---
    if not TEST_MODE and not is_taiwan_workday(today):
        print(f"[{now.strftime('%H:%M:%S')}] 今天是假日，跳過打卡。")
        return

    # --- 2. 打卡時間段判斷 ---
    target_id = ""
    punch_name = ""

    if TEST_MODE:
        target_id = "btnclock1"
        punch_name = "簽到(測試)"
    else:
        # 上午簽到時間段：08:00 ~ 09:30 (可自行微調)
        if 8 <= now.hour < 10: 
            target_id = "btnclock1"
            punch_name = "上班簽到"
        # 下午簽退時間段：18:30 之後 (到 23:59)
        elif now.hour >= 18:
            if now.hour == 18 and now.minute < 30: # 如果是 18:00~18:29 則不執行
                print(f"[{now.strftime('%H:%M:%S')}] 尚未到達 18:30，暫不簽退。")
                return
            target_id = "btnclock2"
            punch_name = "下班簽退"
        else:
            print(f"[{now.strftime('%H:%M:%S')}] 非打卡時段，程式結束。")
            return

        # 隨機延遲 (避免準點打卡太假)
        wait_sec = random.randint(30, 180)
        print(f"符合 {punch_name} 時段，預計 {wait_sec} 秒後執行...")
        time.sleep(wait_sec)

    # --- 3. 執行 Selenium ---
    options = Options()
    if HEADLESS_MODE: options.add_argument('--headless')
    options.add_argument('--start-maximized')
    options.add_argument("--disable-blink-features=AutomationControlled")
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    try:
        print(f"正在連線 ERP 並執行 {punch_name}...")
        driver.get(ERP_URL)
        time.sleep(3)
        
        driver.find_element(By.ID, "login_name").send_keys(USER_ID)
        driver.find_element(By.ID, "password").send_keys(PASSWORD + Keys.ENTER)
        time.sleep(12)

        menu_js = """
        var elements = document.querySelectorAll('span, a, div');
        for (var el of elements) {
            if (el.textContent.trim() === '出勤線上打卡') {
                el.click();
                return true;
            }
        }
        """
        driver.execute_script(menu_js)
        time.sleep(10)

        found = False
        driver.switch_to.default_content()
        if len(driver.find_elements(By.ID, target_id)) > 0:
            found = True
        else:
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            for i in range(len(iframes)):
                driver.switch_to.default_content()
                driver.switch_to.frame(i)
                if len(driver.find_elements(By.ID, target_id)) > 0:
                    found = True
                    break
        
        if found:
            driver.execute_script(f"document.getElementById('{target_id}').click();")
            log_and_notify(f"✅ {punch_name} 成功！")
            time.sleep(5)
        else:
            log_and_notify(f"❌ 找不到{punch_name}按鈕", is_error=True)
            driver.save_screenshot("failed_debug.png")

    except Exception as e:
        log_and_notify(f"❌ 發生異常: {str(e)[:50]}", is_error=True)
    finally:
        driver.quit()

if __name__ == "__main__":
    run_punch()