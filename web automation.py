https://stats.labs.apnic.net/ipv6/MN
import os                                                   # Operating system-тэй холбоотой функцуудад ашиглана
import sys                                                  # System-ийн параметрүүд болон функуудад ашиглана
from datetime import datetime                               # Огноо болон цагийн мэдээлэл
from io import StringIO                                     # Файлтай ажиллаж буй мэт String-тэй ажиллах 
import pandas as pd                                         # HTML дээр буй хүснэгтийш уншиж excel файлд бичих
from selenium import webdriver                              # Browser-тэй  ажиллах
from selenium.webdriver.chrome.options import Options       # Chrome-ийг хүссэнээрээ тохируулах
from bs4 import BeautifulSoup                               # HTML/XML хуулна
import time                                                 
import traceback                                            # Алдааны лог

# --- Тохиргоо ---
# Хүснэгт авах URL-ийг тодорхойлох (APNIC statistics page)
URL = "https://stats.labs.apnic.net/ipv6/MN" 
# Файлаа хадгалах directory тодорхойлох
OUTPUT_DIR = r"C:\Users\sugar\Desktop\Test"
# Хадгалах файлын нэр
EXCEL_FILENAME = "apnic_tables_history.xlsx"

def scrape_tables():
    
    chrome_options = Options()
    # browser-ийг headless mode-д ажиллуулах (GUI-гүйгээр)
    chrome_options.add_argument("--headless=new") 
    # Аюулгүй байдлын үүднээс sandbox mode-ийг унтраах 
    chrome_options.add_argument("--no-sandbox")   
    # Shared memory (/dev/shm) ашиглалтыг унтраана.
    chrome_options.add_argument("--disable-dev-shm-usage") 
    # GPU hardware acceleration-ийг унтраана.
    chrome_options.add_argument("--disable-gpu")  
    chrome_options.add_argument("window-size=1920x1080")
    # Custom user agent. Зарим вэбсайт headless Chrome-ийг илрүүлээд блоклодог.
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")

    # Chrome WebDriver-ийг дээр тохируулснаар ажиллуулах
    driver = webdriver.Chrome(options=chrome_options)
    # page-ийг ачааллах maximum time
    driver.set_page_load_timeout(30)
    # Элементүүд гарч ирэхийг хүлээх (3 сек)
    driver.implicitly_wait(3)    
    
    try:
        print(f"[{datetime.now()}] 1: Вебийг нээж байна...")
        # URL-аар нь дамжуулан веб рүү хандах
        driver.get(URL)
        time.sleep(5)  
        
        print(f"[{datetime.now()}] 2: Вебийг хуулж авч байна...")
        # BeautifulSoup ашиглан HTML кодыг ялган авах
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        
        # Ялгасан HTML кодоос хүснэгтэн мэдээллийг ялгаад pandas DataFrame list-рүү хувиргах
        tables = pd.read_html(StringIO(str(soup)))
        
        # page дээр хүснэгт олдоогүй бол
        if not tables:
            print("Хүснэгт олдоогүй.")
            driver.quit()
            sys.exit(0)  

        # Хэдэн хүснэгт олдсныг хэвлэх
        print(f"{len(tables)} хүснэгт оллоо...")
        
        # Excel file-ийн бүтэн path-ийг үүсгэх
        excel_file = os.path.join(OUTPUT_DIR, EXCEL_FILENAME)
        
        # Файлд ямар эрхтэй хандахыг заах
        # Файл үүсчихсэн бол append ('a'), шинээр бол шинэ файл үүсгэнэ ('w')
        if os.path.exists(excel_file):
            writer_mode = 'a'
        else:
            writer_mode = 'w'

        # sheet бүрийг нэрлэхэд ашиглахдаа цаг хугацааг ашиглана
        timestamp_str = datetime.now().strftime("%Y-%m-%d_%H%M_%S")
        
        try:
            # Тохирох горимд openpyxl ашиглан Excel writer объект үүсгэх
            with pd.ExcelWriter(excel_file, engine='openpyxl', mode=writer_mode) as writer:
                # Олдсон бүх хүснэгтийг унших
                for i, df in enumerate(tables):
                    # Хоосон эсвэл 2-оос бага мөртэй хүснэгтийг алгасах
                    if df.empty or df.shape[0] < 2: 
                        continue
                    # Огноо, цаг болон хүснэгтийн дугаараар шинэ sheet үүсгэх
                    sheet_name = f"{timestamp_str}_T{i+1}"
                    # DataFrame-ийг Excel-д шинэ sheet болгон нэмэх
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Амжилттай! Шинэ хүснэгтийг {excel_file}-д нэмсэн")
            
        except PermissionError:
            # PermissionError буюу уг файл өөр процессоор lock-логдсон
            # backup файлд зориулан Огноо болон цагаар нэр үүсгэх
            alt_timestamp = datetime.now().strftime("%H%M%S")
            # backup file-ийн нэрийг lock error-той байгааг илтгэх байдлаар үүсгэх
            backup_filename = f"apnic_tables_history_LOCKED_{alt_timestamp}.xlsx"
            # excel_file path-ийг шинэ backup файл болгон өөрчлөх
            excel_file = os.path.join(OUTPUT_DIR, backup_filename)
            print(f"Файл lock-логдсон! backup файлд мэдээллийг хадгаллаа: {excel_file}")
            
            # backup файлд мэдээллийг оруулах ('w' mode-той шинэ файл үүсгэнэ)
            with pd.ExcelWriter(excel_file, engine='openpyxl', mode='w') as writer:
                # Хүснэгтүүдийг уншаад backup file-д бичих
                for i, df in enumerate(tables):
                    # хоосон хүснэгт алгасах
                    if df.empty or df.shape[0] < 2: 
                        continue
                    # Огноо болон цагийн нэртэй шинэ sheet үүсгэх
                    sheet_name = f"{timestamp_str}_T{i+1}"
                    # DataFrame-ийг backup Excel файлд оруулах
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        driver.quit()
        print("Процесс дууслаа.")
        sys.exit(0)

    except Exception as e:
        # Програм ажиллаж байх үед гарсан бусад алдаа
        # Алдааны лог файлын path
        log_path = os.path.join(OUTPUT_DIR, "error_log.txt")
        print(f"ERROR! {log_path}-ийг шалга")
        
        try:
            # Алдааны лог файлыг UTF-8 аргаар encode хийх
            with open(log_path, "w", encoding="utf-8") as log_file:
                # Алдааны огноо болон цагыг бичих
                log_file.write(f"Timestamp: {datetime.now()}\n")
                # Алдааг бичих
                log_file.write(f"Error Message: {str(e)}\n\n")
                log_file.write(traceback.format_exc())
        except:
            pass 
            
        finally:
            driver.quit()
            sys.exit(1) 

if __name__ == "__main__":
    scrape_tables()

# chmod +x web-automation.py
# crontab -e
# * 12 * * * /usr/bin/python3 /home/23b1num1379/web-automation.py
