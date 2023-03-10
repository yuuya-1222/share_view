from selenium import webdriver
from selenium.webdriver.common.by import By
#import gspread
import time
import datetime
import re


''
#-------------------------------------------------------------------
# colab用のgoogle認証なのでcolab外で動かすなら変更必須、スプレッドシート開くのもここで

from google.colab import auth
from oauth2client.client import GoogleCredentials
import gspread
auth.authenticate_user()
gc = gspread.authorize(GoogleCredentials.get_application_default())

ss_url = "https://docs.google.com/spreadsheets/d/1uzX47207xDxl_no8uC7BpJQns0hRPkOHrBRmFMDdVQ8/edit#gid=0"
wb = gc.open_by_url(ss_url)

wb.add_worksheet(title='●日(●)')
#ws = wb.worksheet('●日(●)')
ws = wb.worksheet('テストシート')


''
#-------------------------------------------------------------------
# maildealerパラメータ(垢使い分けない限りは操作不要)

url = 'https://mdchorus.maildealer.jp/index.php'
login = 'CRAVEARKS様'
password = 'cravearks0424'


#-------------------------------------------------------------------
# ドライバー設定及びMaildealerログイン(colab用になってるので注意)

options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome('chromedriver', options=options)
driver.implicitly_wait(5)
driver.get(url)
time.sleep(3)

driver.find_element(By.XPATH, '//*[@id="fUName"]').send_keys(login)
driver.find_element(By.XPATH, '//*[@id="fPassword"]').send_keys(password)
driver.find_element(By.XPATH, '//*[@id="d_login_input"]/input[2]').click()
time.sleep(5)


#-------------------------------------------------------------------
# Rit様フォルダ開いて(今テストなので試しに不明フォルダ展開あり)、メインフレームへ移動まで

side_iframe = driver.find_element(By.ID, 'ifmSide')
driver.switch_to.frame(side_iframe)
driver.find_element(By.XPATH, '//*[@id="sec"]/div/div[1]/ul/li/ul[39]/li/div/button').click()
time.sleep(3)

# ↓ここテスト用途に応じて切り替え忘れずに()
#element = driver.find_element(By.XPATH, '//*[@id="sec"]/div/div[1]/ul/li/ul[39]/li/div[1]')
element = driver.find_element(By.XPATH, '//*[@id="sec"]/div/div[1]/ul/li/ul[39]/li/div[2]/ul/li/div')
driver.execute_script("arguments[0].click();", element)
time.sleep(1)

# これは後でswitch_frame関数(SideverとMainver)としてまとめたい
driver.switch_to.default_content()
main_iframe = driver.find_element(By.ID, 'ifmMain')
driver.switch_to.frame(main_iframe)


#-------------------------------------------------------------------
# mail件数取得→必要に応じてページ送り→一番古いメールのID取得→クリック

mail_count_total = driver.find_element(By.XPATH, '//*[@id="form-olv-p-maillist"]/div[1]/div/div[1]')
res = re.search(r'/(.*)', mail_count_total.text).group()
oldest_mail_number = int(re.sub(r"[/ 件]", "", res))

while oldest_mail_number > 50:
driver.find_element(By.XPATH, '//*[@id="form-olv-p-maillist"]/div[1]/div/div[2]/button[2]').click()
time.sleep(1)
oldest_mail_number -= 50

#mail_count = len(driver.find_elements(By.CLASS_NAME, 'olv-c-table__tbody'))
mail_ID_path = f'//*[@id="form-olv-p-maillist"]/div[2]/div/div[2]/table/tbody[{oldest_mail_number}]/tr/td[3]/div/span'
mail_ID = driver.find_element(By.XPATH, mail_ID_path)
print(mail_ID.text)
driver.execute_script("arguments[0].click();", mail_ID)
time.sleep(1)


#-------------------------------------------------------------------
# メール本文とアドレス取得→タイプ別に必要な情報抽出

mail_body = driver.find_element(By.XPATH, '//*[@id="form-olv-p-viewmail"]/div[2]/section/div[2]/pre')
print(mail_body.text)
mail_address = driver.find_element(By.XPATH, '//*[@id="form-olv-p-viewmail"]/div[2]/section/div[1]/div[1]/ul/li[1]/span')
print(mail_address.text)

if 'talkmation' in mail_address.text:
res = re.search(r'電話番号：(.*)\n', mail_body.text).group()
phone_number = re.sub(r"[電話番号： ]", "", res)
print(phone_number)
res = re.search(r'顧客名：(.*)\n', mail_body.text).group()
customer_name = re.sub(r"[顧客名： ]", "", res)
print(customer_name)
res = re.search(r'アドレス：(.*)\n', mail_body.text).group()
customer_address = re.sub(r"[アドレス： ]", "", res)
print(customer_address)



driver.quit()
