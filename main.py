from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from time import sleep
import openpyxl
from settings import EMAIL, PASSWORD

FILE_NAME = "example.xlsx"
COMPANY = "servicenow"

try:
    # varolan dosyayı yükle
    wb = openpyxl.load_workbook(FILE_NAME)
    sheet = wb.active
except:
    # dosya yoksa oluştur
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.append([
        "Owner of the post (name)",
        "URL of the owner of the post (name)",
        "date",
        "text"
    ])


class Browser:

    def __init__(self):
        self.browser = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        self.browser.maximize_window()

    def get(self):
        return self.browser


driver = Browser().get()
driver.get(f"https://www.linkedin.com/company/{COMPANY}/posts/")
sleep(2)
# giriş yap
driver.find_element(By.ID, "username").send_keys(EMAIL)
driver.find_element(By.ID, "password").send_keys(PASSWORD)
driver.find_element(By.XPATH, "//button[@type='submit']").click()

# gönderileri al
# sonsuz kaydırma işlemi için sayfayı aşağı kaydır
posts_count = 0
while True:
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    posts = driver.find_elements(By.XPATH, "//div[contains(@class,'occludable-update')]")
    if posts_count == len(posts):
        break

    posts_count = len(posts)
    sleep(1)


# gönderileri al
for post in posts:

    # gönderi sahibi linkini al
    try:
        post_owner_link = post.find_element(By.XPATH, ".//a[contains(@class, 'update-components-actor__container-link')]").get_attribute("href").split("?")[0]
        post_owner = post.find_element(By.XPATH, ".//span[contains(@class, 'update-components-actor__name')]").text
        post_date = post.find_element(By.XPATH, ".//span[contains(@class, 'update-components-actor__sub-description')]").text.split(" • ")[0]
        post_text = post.find_element(By.XPATH, ".//div[contains(@class, 'feed-shared-update-v2__description-wrapper')]").text
    except:
        # eğer yoksa kişiye ait değil [reklam vb.] geç
        continue

    sheet.append([post_owner, post_owner_link, post_date, post_text])



# tarayıcıyı kapat
driver.close()
wb.save(FILE_NAME)
