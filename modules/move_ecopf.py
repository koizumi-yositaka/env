import chromedriver_binary
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.keys import Keys
import time
# WebDriver のオプションを設定する
options = webdriver.ChromeOptions()
#options.add_argument('--headless')
 
print('connectiong to remote browser...')
driver = webdriver.Chrome(options=options)
 
driver.get('http://localhost/service/eco-pf/WebApp/DESolution/Login.aspx')
print(driver.current_url)

#
txt_user=driver.find_element(By.ID,"txtUser")
txt_user.clear()
txt_user.send_keys("admin")
#txt_pass=driver.find_element_by_id("txtUser")
txt_password=driver.find_element(By.ID,"txtPassword")
txt_password.clear()
txt_password.send_keys("test")
txt_password.send_keys(Keys.ENTER)
#ログイン完了
wait=WebDriverWait(driver,60)
wait.until(expected_conditions.frame_to_be_available_and_switch_to_it((By.ID,"FrameMain")))
wait.until(expected_conditions.frame_to_be_available_and_switch_to_it((By.ID,"FrameL")))
a_treeview=driver.find_element(By.ID,"Tree1n0")
a_treeview.click()
top_treeview_div = driver.find_element(By.ID, "Tree1n0Nodes")
a_init_class_node=top_treeview_div.find_element(By.CLASS_NAME, "Tree1_0")
a_home=a_init_class_node.find_element(By.TAG_NAME, "a")
#print(a_init_class[0])
wait.until(expected_conditions.element_to_be_clickable(a_home))
a_home.click()
print(driver.page_source)
driver.save_screenshot("aaaa.png")
#クラス編集画面の表示





driver.implicitly_wait(20)



# ブラウザを終了する
#driver.quit()