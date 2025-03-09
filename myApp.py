import os
import glob
import sys
import random
import time
from colorama import *
from PyQt6.QtCore import QStringListModel, QAbstractItemModel
from PyQt6.QtGui import QStandardItemModel, QStandardItem
import pandas as pd
import random
#cài config selenium
from seleniumwire import webdriver  # Chặn request
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.alert import Alert
import openpyxl
import requests
from selenium_stealth import stealth

from time import sleep

from PyQt6 import uic
from PyQt6.QtWidgets import QApplication, QMainWindow, QMessageBox, QFileDialog, QMessageBox
from myAppUi import Ui_Form  
class UI(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # uic.loadUi("myApp.ui", self)
        self.ui = Ui_Form  ()
        self.ui.setupUi(self)
        # self.bntChooseFile.clicked.connect(self.openFileDialog)
        # self.bntAction.clicked.connect(self.process)
        self.ui.bntChooseFile.clicked.connect(self.openFileDialog)
        self.ui.bntAction.clicked.connect(self.process)

        self.file_path = ""
            
    def openFileDialog(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Chọn file", "", "All Files (*);;Text Files (*.txt)")
        if file_path:
            self.ui.lbLinkfile.setText(f"Đã chọn: {file_path}")
            self.file_path = file_path
        else:
            self.ui.lbLinkfile.setText("Chưa chọn file")

    def process(self):
        if not self.file_path:
            QMessageBox.information(self, "Thông báo", "Chưa chọn file")
            return
        try:            
            df = pd.read_excel(self.file_path, engine="openpyxl")

            # Hoặc hiển thị từng dòng
            for index, row in df.iterrows():
                data = row.to_dict()  
                IRSFormPage().action_form(data, index+1, self.file_path)
        except Exception as e:
            QMessageBox.information(self, "Thông báo", f"Có lỗi xảy ra: {str(e)}")
            
class IRSFormPage:
    def __init__(self):
        
        download_path = r"C:\Tutorial\down"
        
        chrome_options  = Options()
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # Ẩn chế độ tự động
        chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36")
        chrome_options.add_argument("--ignore-certificate-errors")
        chrome_options.add_argument("--allow-running-insecure-content")
        chrome_options.add_argument("--start-maximized")  # Mở rộng cửa sổ
        prefs = {
            "download.default_directory": download_path,  # Đặt thư mục tải file
            "download.prompt_for_download": False,          # Không hỏi xác nhận tải file
            "download.directory_upgrade": True,             # Cập nhật thư mục nếu cần
            "plugins.always_open_pdf_externally": True        # Nếu tải PDF, sẽ tải về thay vì mở trong trình duyệt
        }
        # Tắt thông báo "Chrome is being controlled..."
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        # Tắt extension tự động hóa
        chrome_options.add_experimental_option("useAutomationExtension", False)
        chrome_options.add_experimental_option("prefs",prefs);
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        stealth(self.driver,
                languages=["en-US", "en"],
                vendor="Google Inc.",
                platform="Win32",
                webgl_vendor="Intel Inc.",
                renderer="Intel Iris OpenGL Engine",
                fix_hairline=True)
        
    def action_form(self, data, index, file_path):
        try:
            
            self.driver.delete_all_cookies() 
                      
            self.driver.get("https://sa.www4.irs.gov/modiein/individual/index.jsp")
            time.sleep(2)

            #identifi Page
            try:
                alert = self.driver.switch_to.alert  # Chuyển sang alert
                print("Popup message:", alert.text)  # In ra nội dung popup
                alert.accept()  # Nhấn OK
                print("Đã nhấn OK trên popup.")
            except:
                print("Không tìm thấy popup.")

            
            self.click_button("//input[@type='submit' and @value='Begin Application >>']", "Begin Application")
        
            self.select_radio("//input[@type='radio' and @id='sole']")
                
            self.click_button("//input[@type='submit' and @value='Continue >>']", "Continue")
    
            self.select_radio("//input[@type='radio' and @value='30' and @id='sole']")
                
            self.click_button("//input[@type='submit' and @value='Continue >>']", "Continue")
        
            self.click_button("//input[@type='submit' and @value='Continue >>']", "Continue")
                
            self.select_radio("//input[@type='radio' and @name='radioReasonForApplying']")
            
            self.click_button("//input[@type='submit' and @value='Continue >>']", "Continue")
        
            self.auto_fill_form_Authenticate_2(data)
            self.auto_fill_form_Addresses_3(data)
            self.auto_fill_form_Details_4()
            self.auto_fill_form_EIN_Confirmation_5(index, file_path)
        except Exception as e:
            print(f"❌ Dừng xử lý dòng {index} do lỗi: {e}")
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            self.driver.delete_all_cookies()  # Xóa toàn bộ cookies
            # Xóa Local Storage & Session Storage
            self.driver.execute_script("window.localStorage.clear();")
            self.driver.execute_script("window.sessionStorage.clear();")

            self.driver.quit()
            return  # Thoát action_form(), quay lại vòng lặp tiếp theo
                   
    def auto_fill_form_Authenticate_2(self, data):
        self.fill_form_input("applicantFirstName", data["FN"])
        sleep(0.5)
        self.fill_form_input("applicantLastName", data["LN"])
        sleep(0.5)
        # Chuyển SSN thành chuỗi và đảm bảo đủ 9 số
        ssn_str = str(data["SSN"]).zfill(9)

        # Tách thành 3 phần
        ssn_part1 = ssn_str[:3]  # 3 số đầu
        ssn_part2 = ssn_str[3:5]  # 2 số giữa
        ssn_part3 = ssn_str[5:]   # 4 số cuối
        self.fill_form_input("applicantSSN3", ssn_part1)
        sleep(0.5)
        self.fill_form_input("applicantSSN2", ssn_part2)
        sleep(0.5)
        self.fill_form_input("applicantSSN4", ssn_part3)
        
        #click checkbox
        self.select_radio("//input[@type='radio' and @name='tpdQuestion' and @id='iamsole']")
        sleep(0.5)
        self.click_button("//input[@type='submit' and @value='Continue >>']", "Continue") 
    
    def auto_fill_form_Addresses_3(self, data):
        self.fill_form_input("physicalAddressStreet", data["ADD"])
        sleep(0.5)
        self.fill_form_input("physicalAddressCity", data["CITY"])
        sleep(0.5)
        self.select_state(data["STATE"])
        sleep(0.5)
        self.fill_form_input("physicalAddressZipCode", data["ZIP"])
        
        first3 = str(random.randint(100, 999))  # 3 số đầu
        middle3 = str(random.randint(100, 999))  # 3 số giữa
        last4 = str(random.randint(1000, 9999))  # 4 số cuối
        
        self.fill_form_input("phoneFirst3", first3)
        sleep(0.5)
        self.fill_form_input("phoneMiddle3", middle3)
        sleep(0.5)
        self.fill_form_input("phoneLast4", last4)
        sleep(0.5)
        
        self.select_radio("//input[@type='radio' and @name='radioAnotherAddress' and @id='radioAnotherAddress_n' and @value='false']")
        sleep(0.5)
        self.click_button("//input[@type='submit' and @value='Continue >>']", "Continue") 
        
    def auto_fill_form_Details_4(self):
        month_ramdom = str(random.randint(1,9))
        year_ramdom = str(random.randint(2022,2024))
        
        #page 1
        self.select_month(month_ramdom)
        self.fill_form_input("BUSINESS_OPERATIONAL_YEAR_ID", year_ramdom)
        self.click_button("//input[@type='submit' and @value='Continue >>']", "Continue") 

        #page 2
        self.select_radio("//input[@type='radio' and @name='radioTrucking' and @id='radioTrucking_n' and @value='false']")
        self.select_radio("//input[@type='radio' and @name='radioInvolveGambling' and @id='radioInvolveGambling_n' and @value='false']")
        self.select_radio("//input[@type='radio' and @name='radioExciseTax' and @id='radioExciseTax_n' and @value='false']")
        self.select_radio("//input[@type='radio' and @name='radioSellTobacco' and @id='radioSellTobacco_n' and @value='false']")
        self.select_radio("//input[@type='radio' and @name='radioHasEmployees' and @id='radioHasEmployees_n' and @value='false']")
        
        self.click_button("//input[@type='submit' and @value='Continue >>']", "Continue") 
            
        #page 3
        self.select_radio("//input[@type='radio' and @name='radioPrincipalActivity' and @id='other' and @value='15']")
        self.click_button("//input[@type='submit' and @value='Continue >>']", "Continue") 
        
        #page 4
        self.select_radio("//input[@type='radio' and @name='radioPrincipalService' and @id='other' and @value='line15OpenEntry']")
        self.fill_form_input("pleasespecify", "digital")
        self.click_button("//input[@type='submit' and @value='Continue >>']", "Continue") 

        #page 5
        #check ô đầu tiên
        self.select_radio("//input[@type='radio' and @name='radioLetterOption' and @id='receiveonline' and @value='online']")
        self.click_button("//input[@type='submit' and @value='Continue >>']", "Continue") 

        #page 6
        self.click_button("//input[@type='submit' and @value='Submit'and @name='Submit']", "Submit completed") 
        
    def auto_fill_form_EIN_Confirmation_5(self, index, file_path):
        print("page 5, para, index excell  ", index,"  file path: ", file_path)
        ein_assigned = self.driver.find_element(By.XPATH, "//td[text()='EIN Assigned:']/following-sibling::td/b").text

        # Lấy Legal Name
        legal_name = self.driver.find_element(By.XPATH, "//td[text()='Legal Name:']/following-sibling::td/b").text


        print(f"EIN Assigned: {ein_assigned}")
        print(f"Legal Name: {legal_name}")
        
        # Mở file Excel
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active

        # Ghi dữ liệu vào đúng dòng
        sheet.cell(row=index+1, column=1, value=ein_assigned)  # Cột 1 (EIN)
        sheet.cell(row=index+1, column=2, value=legal_name)  # Cột 2 (Legal Name)
        
        # Lưu lại file Excel
        workbook.save(file_path)
        workbook.close()

        try:
            # download_path = os.path.abspath("downloads")
            # if not os.path.exists(download_path):
            #     os.makedirs(download_path)
            
            pdf_link_element = self.driver.find_element(By.XPATH, "//*[contains(@href, '.pdf')]")
            pdf_url = pdf_link_element.get_attribute("href")
            print(f"📌 Link PDF: {pdf_url}")
            # Click để tải file
            pdf_link_element.click()
            time.sleep(5)  # Chờ tải xong

            # Tìm nút tải file CSV (theo CSS selector) và click vào đó
            downloadcsv = self.driver.find_element(By.CSS_SELECTOR, '#baseSvg')
            downloadcsv.click()

            

            # # Lấy file mới nhất trong thư mục
            # pdf_files = glob.glob(os.path.join(download_path, "*.pdf"))
            # if pdf_files:
            #     latest_file = max(pdf_files, key=os.path.getctime)
            #     new_filename = os.path.join(download_path, "document.pdf")  # Đổi tên file
            #     os.rename(latest_file, new_filename)
            #     print(f"✅ File đã được đổi tên thành: {new_filename}")
            # else:
            #     print("❌ Không tìm thấy file PDF để đổi tên")
        except Exception as e:
            print(f"❌ Lỗi khi tải PDF: {e}")
    
    def select_state(self, state):
        try:
            dropdown = Select(self.driver.find_element(By.ID, "physicalAddressState"))
            dropdown.select_by_value(state)
        except Exception as e:
            raise Exception(f"❌ Lỗi khi chọn state: {e}")
            
    def select_month(self, month):
        try:
            dropdown = Select(self.driver.find_element(By.ID, "BUSINESS_OPERATIONAL_MONTH_ID"))
            dropdown.select_by_value(month)
        except Exception as e:
            raise Exception(f"❌ Lỗi khi chọn state: {e}")
                
    def fill_form_input(self, xpath, data):
        try:
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, xpath))).send_keys(data)
            print(f"Đã điền {data} vào {xpath}.")
        except Exception as e:
            raise Exception(f"❌ Lỗi khi điền {data} vào {xpath}.")
     
    def select_radio(self, xpath):
        try:
            radio_button = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath)))
            radio_button.click()
            print(f"Đã chọn radio.")
        except Exception as e:
            raise Exception(f"❌ Lỗi khi chọn radio: {e}")
    
    def click_button(self, xpath, description, repeat=1):
        for _ in range(repeat):
            try:
                button = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath)))
                button.click()
                print(f"Đã nhấn nút {description}.")
                time.sleep(2)
            except Exception as e:
                raise Exception(f"❌ Lỗi khi nhấn nút {description}: {e}")

                
            
if __name__ == "__main__":
    try:
        app = QApplication([])
        window = UI()
        window.show()
        app.exec()
    except KeyboardInterrupt:
        sys.exit()
            