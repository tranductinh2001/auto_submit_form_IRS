from selenium import webdriver

chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": r"C:\Users\Admin\Downloads",  # Thư mục lưu file
    "download.prompt_for_download": False,  # Không hỏi trước khi tải
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True  # Mở PDF ngoài trình duyệt để tải ngay
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=chrome_options)
driver.get("https://sampletestfile.com/pdf/")

# Tìm nút tải xuống và click
download_button = driver.find_element("xpath", "#post-131 > div > div > div > div.elementor-element.elementor-element-4946a7d.e-flex.e-con-boxed.e-con.e-parent > div > a")
download_button.click()
