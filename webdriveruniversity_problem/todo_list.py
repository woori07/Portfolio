from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time


class ManagerList():
    def __init__(self):
        self.driver = webdriver.Chrome()
        self.driver.get('https://webdriveruniversity.com/To-Do-List/index.html')
    
    def add_todo(self, str):    # 일정 추가
        self.driver.find_element(By.XPATH, '//*[@id="container"]/input').send_keys(str)
        self.driver.find_element(By.XPATH, '//*[@id="container"]/input').send_keys(Keys.RETURN)
        time.sleep(1)
        
    def remove_todo(self, int):     # 일정 삭제
        hoverable = self.driver.find_element(By.XPATH, f'//*[@id="container"]/ul/li[{int}]/span/i')
        ActionChains(self.driver).move_to_element(hoverable).perform()
        self.driver.find_element(By.XPATH, f'//*[@id="container"]/ul/li[{int}]/span/i').click()
        time.sleep(3)
    
if __name__ == "__main__":
    # 일정 2개 추가 후, 4번째 일정 삭제
    managerlist = ManagerList()
    managerlist.add_todo('new_work')
    managerlist.add_todo('old_work')
    managerlist.remove_todo(4)
        
    