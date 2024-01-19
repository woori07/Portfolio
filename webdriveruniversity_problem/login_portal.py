from selenium import webdriver
from selenium.webdriver.common.by import By
import time

class LoginPortal_Problem():
    def xpath_based_click(self):
        self.driver = webdriver.Chrome()        # xpath 기반 클릭 기법
        self.driver.get("https://webdriveruniversity.com/Click-Buttons/index.html")
        self.driver.find_element(By.XPATH, '//*[@id="button1"]').click()
        time.sleep(1)
        self.driver.find_element(By.XPATH, '//*[@id="myModalClick"]/div/div/div[3]/button').click()
        
    def css_based_click(self):      # css 기반 클릭 기법
        self.driver.find_element(By.CSS_SELECTOR, '#button2').click()
        time.sleep(1)
        self.driver.find_element(By.CSS_SELECTOR, '#myModalJSClick > div > div > div.modal-footer > button').click()
        
    def anylocator_click(self):     # ID 기반 클릭 기법
        self.driver.find_element(By.ID, 'button3').click()
        time.sleep(1)
        self.driver.find_element(By.XPATH, '//*[@id="myModalMoveClick"]/div/div/div[3]/button').click()
        time.sleep(1)

if __name__ == "__main__":
    # 3가지의 클릭 기법 사용 후 팝업 닫기
    clicktest = LoginPortal_Problem()
    clicktest.xpath_based_click()
    clicktest.css_based_click()
    clicktest.anylocator_click()
    
    