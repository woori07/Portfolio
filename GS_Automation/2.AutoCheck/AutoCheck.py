from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as expected_conditions
import time
import pandas as pd 

# gsNum 입력
gsNumList = input()
gsNumIndex = gsNumList.split()

# gsNum을 result_df 데이터프레임에 저장
gsFrame = pd.DataFrame()
gsFrame['인증번호'] = gsNumIndex

# 컬럼 선언
gsFrame['저작권 확인서'] = ''
gsFrame['프로그램 등록부'] = ''

# 브라우저 디버거 모드
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
driver = webdriver.Chrome(options=chrome_options)

# 분당ECM 탐색
driver.get('secret')
time.sleep(1.5)

jjlist = []
prolist = []

# for문으로 파일 체크
for a in range(0, len(gsNumIndex)):

    # gsNum[i]값을 가져와서 '-' 기준으로 split하여 배열(gsSp[3])로 나눔.
    block = gsNumIndex[a].split('-')
 
    # A-23
    if(block[1]=="A" and block[2]=="23"):

        time.sleep(2)
        print(gsNumIndex[a])
        driver.find_element(By.XPATH, '//*[@id="edm-folder"]/ul/li[3]/ul/li[23]/ins').click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//*[contains(text(),'" + "GS시험인증(1등급)" + "')]").click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//*[contains(text(),'" + gsNumIndex[a] + "')]").click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//*[contains(text(),'" + "인증관련" + "')]").click()
        time.sleep(1.5)
        try:
            if (driver.find_element(By.CLASS_NAME, "list-no-data").is_displayed()):
                print("확인서 및 등록증이 존재하지 않습니다.")
                jjlist.insert(a,'x')
                prolist.insert(a,'x')
                print(jjlist)
                print(prolist)
                
        except:
                print("확인서 및 등록증이 존재합니다.")
                jjlist.insert(a,'o')
                prolist.insert(a,'o')
                print(jjlist)
                print(prolist)
                
        driver.find_element(By.XPATH, '//*[@id="edm-folder"]/ul/li[3]/ul/li[23]/ins').click()

    # A-22
    elif(block[1]=="A" and block[2]=="22"):

        time.sleep(2)
        print(gsNumIndex[a])
        driver.find_element(By.XPATH, '//*[@id="edm-folder"]/ul/li[3]/ul/li[22]/ins').click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//*[contains(text(),'" + "GS시험인증(1등급)" + "')]").click()
        time.sleep(1.3)
        driver.find_element(By.XPATH, "//*[contains(text(),'" + gsNumIndex[a] + "')]").click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//*[contains(text(),'" + "인증관련" + "')]").click()
        time.sleep(1.5)
        try:
            if (driver.find_element(By.CLASS_NAME, "list-no-data").is_displayed()):
                print("확인서 및 등록증이 존재하지 않습니다.")
                jjlist.insert(a,'x')
                prolist.insert(a,'x')
                print(jjlist)
                print(prolist)
                
        except:
            
                print("확인서 및 등록증이 존재합니다.")
                jjlist.insert(a,'o')
                prolist.insert(a,'o')
                print(jjlist)
                print(prolist)
    
        driver.find_element(By.XPATH, '//*[@id="edm-folder"]/ul/li[3]/ul/li[22]/ins').click()

    # B or C 일 경우
    else:
        pass

print(jjlist)
print(prolist)


# 상암ECM 탐색 
driver.get('secret')

# 로그인 페이지 대기
wait = WebDriverWait(driver, 15)
element = wait.until(expected_conditions.element_to_be_clickable((By.ID, "wrap")))

driver.find_element(By.NAME, 'user_id').send_keys('manager')
driver.find_element(By.NAME, 'password').send_keys('manager123')
driver.find_element(By.XPATH, '//*[@id="form-login"]/div[2]').click()

# for 문으로 파일 체크
for a in range(0, len(gsNumIndex)):

    # gsNum[i]값을 가져와서 '-' 기준으로 split하여 배열(gsSp[3])로 나눔.
    block = gsNumIndex[a].split('-')

    # B-23
    if(block[1]=="B" and block[2]=="23"):

        time.sleep(2)
        print(gsNumIndex[a])
        driver.find_element(By.XPATH, '//*[@id="edm-folder"]/ul/li[4]/ul/li[17]/ins').click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="edm-folder"]/ul/li[4]/ul/li[17]/ul/li[2]/ins').click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//*[contains(text(),'" + gsNumIndex[a] + "')]").click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//*[contains(text(),'" + "인증관련" + "')]").click()
        time.sleep(1.5)
        try:
            if (driver.find_element(By.CLASS_NAME, "list-no-data").is_displayed()):
                print("확인서 및 등록증이 존재하지 않습니다.")
                jjlist.insert(a,'x')
                prolist.insert(a,'x')
                print(jjlist)
                print(prolist)
                
        except:
                print("확인서 및 등록증이 존재합니다.")
                jjlist.insert(a,'o')
                prolist.insert(a,'o')
                print(jjlist)
                print(prolist)
            
        driver.find_element(By.XPATH, '//*[@id="edm-folder"]/ul/li[4]/ul/li[17]/ins').click()

    # B-22
    elif(block[1]=="B" and block[2]=="22"):

        time.sleep(2)
        print(gsNumIndex[a])
        driver.find_element(By.XPATH, '//*[@id="edm-folder"]/ul/li[4]/ul/li[16]/ins').click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="edm-folder"]/ul/li[4]/ul/li[16]/ul/li[2]/ins').click()
        time.sleep(1.3)
        driver.find_element(By.XPATH, "//*[contains(text(),'" + gsNumIndex[a] + "')]").click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//*[contains(text(),'" + "인증관련" + "')]").click()
        time.sleep(1.5)
        try:
            if (driver.find_element(By.CLASS_NAME, "list-no-data").is_displayed()):
                print("확인서 및 등록증이 존재하지 않습니다.")
                jjlist.insert(a,'x')
                prolist.insert(a,'x')
                print(jjlist)
                print(prolist)
                
        except:
                print("확인서 및 등록증이 존재합니다.")
                jjlist.insert(a,'o')
                prolist.insert(a,'o')
                print(jjlist)
                print(prolist)
    
        driver.find_element(By.XPATH, '//*[@id="edm-folder"]/ul/li[4]/ul/li[16]/ins').click()

    # C-23
    elif(block[1]=="C" and block[2]=="23"):

        time.sleep(2)
        print(gsNumIndex[a])
        driver.find_element(By.XPATH, '//*[@id="edm-folder"]/ul/li[4]/ul/li[19]/ins').click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="edm-folder"]/ul/li[4]/ul/li[19]/ul/li[3]/ins').click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//*[contains(text(),'" + "GS시험인증(1등급)" + "')]").click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//*[contains(text(),'" + gsNumIndex[a] + "')]").click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//*[contains(text(),'" + "인증관련" + "')]").click()
        time.sleep(1.5)
        try:
            if (driver.find_element(By.CLASS_NAME, "list-no-data").is_displayed()):
                print("확인서 및 등록증이 존재하지 않습니다.")
                jjlist.insert(a,'x')
                prolist.insert(a,'x')
                print(jjlist)
                print(prolist)
                
        except:
                print("확인서 및 등록증이 존재합니다.")
                jjlist.insert(a,'o')
                prolist.insert(a,'o')
                print(jjlist)
                print(prolist)
            
        driver.find_element(By.XPATH, '//*[@id="edm-folder"]/ul/li[4]/ul/li[19]/ins').click()

    # C-22
    elif(block[1]=="C" and block[2]=="22"):

        time.sleep(2)
        print(gsNumIndex[a])
        driver.find_element(By.XPATH, '//*[@id="edm-folder"]/ul/li[4]/ul/li[19]/ins').click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="edm-folder"]/ul/li[4]/ul/li[19]/ul/li[2]/ins').click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//*[contains(text(),'" + "GS시험인증(1등급)" + "')]").click()
        time.sleep(1.3)
        driver.find_element(By.XPATH, "//*[contains(text(),'" + gsNumIndex[a] + "')]").click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//*[contains(text(),'" + "인증관련" + "')]").click()
        time.sleep(1.5)
        try:
            if (driver.find_element(By.CLASS_NAME, "list-no-data").is_displayed()):
                print("확인서 및 등록증이 존재하지 않습니다.")
                jjlist.insert(a,'x')
                prolist.insert(a,'x')
                print(jjlist)
                print(prolist)
                
        except:
                print("확인서 및 등록증이 존재합니다.")
                jjlist.insert(a,'o')
                prolist.insert(a,'o')
                print(jjlist)
                print(prolist)
    
        driver.find_element(By.XPATH, '//*[@id="edm-folder"]/ul/li[4]/ul/li[19]/ins').click()

    # A 일 경우
    else:
        pass

# 배열을 프레임에 저장
gsFrame['저작권 확인서'] = jjlist
gsFrame['프로그램 등록부'] = prolist

# 상태 체크         
print(gsFrame)

# 경로 설정 및 저장
gsFrame.to_excel(r"C:/Users/이건우/Downloads/GSexcel.xlsx", index=False)