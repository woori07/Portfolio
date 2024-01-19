from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import pymysql
import time

# tr class 태그의 개수를 카운트하는 테스트 함수

class OpenBrowser:
    def startBrowser(self):
        self.driver = webdriver.Chrome()
        self.driver.get('https://gs.thinkwise.co.kr/')
        self.driver.find_element(By.NAME, 'user_id').send_keys('secret')
        self.driver.find_element(By.NAME, 'user_pass').send_keys('secret!')
        self.driver.find_element(By.ID, 'btn-confirm').click()
        alert = self.driver.switch_to.alert
        alert.accept()
        keyword = "회원관리"

        self.driver.find_element(By.XPATH, "//*[contains(text(),'" + keyword + "')]").click()
        time.sleep(3)
        
        
    def getmemeberID_DB(self):
        
        ### DB 연결 ###
        db = pymysql.connect(host='secret', user='secret', password='secret', db='tw_colman',charset='utf8')
        curs = db.cursor()

        ### Select Query ###
        sql = """select * from collaboration_user"""
        curs.execute(sql)
        select = list(curs.fetchall())
        db.commit()

        global text_values
        text_values = []

        for i in range(0, len(select)):
            text_values.append(select[i][1])

        db.close()
        return text_values
        
    def checkData(self, text_values):

        ### 현재 페이지 행 개수 파악 ###
        tr_tags = self.driver.find_elements(By.CSS_SELECTOR, 'tr.tta_ctr')

        ### for문 돌며 데이터가 일치하는지 확인 ###
        for i in range(0, len(tr_tags)):
            element_value = self.driver.find_element(By.XPATH, f'//*[@id="content-wrap"]/div/table/tbody/tr[{(i)+2}]/td[5]').text

            ### 1. assert 검증 방법 ###
            
            try:
                assert(element_value == text_values[i])
            except AssertionError:
                print(text_values[i])

            ### 2. if 검증 방법 ###

            # print(element_value, text_values[i])
            # if (element_value == text_values[i]):
            #     print("correct")
            # else:
            #     print("Inconsistency")
                

if __name__=="__main__":
    openbrowser = OpenBrowser()
    openbrowser.startBrowser()
    openbrowser.getmemeberID_DB()
    openbrowser.checkData(text_values)
