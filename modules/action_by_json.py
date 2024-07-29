import chromedriver_binary # nopa
import json
import time
import datetime
import os
import glob
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd

def reset_folder(folder_path):
    files=glob.glob(os.path.join(folder_path,"*"))
    for file in files:
        try:
            os.remove(file)
            print(f'{file} を削除しました。')
        except Exception as e:
            print(f'{file} の削除中にエラーが発生しました: {e}')



class ControlDisplay(object):
    def __init__(self,type):
        if type=="Chrome":
            options = webdriver.ChromeOptions()
            
            #options.add_argument('--headless')
            
            print('connectiong to remote browser...')
            self.driver = webdriver.Chrome(options=options)
    def read_scenario(self,path):
        with open(path,"r",encoding="UTF-8") as j:
            data=json.load(j)
            try:
                for action in data["Input"]:
                    self.control(action)
                    print(action,"完了")
                self.driver.quit()
            except Exception as e:

                print("Error",e)
                dt_now=datetime.datetime.now()
                self.driver.save_screenshot("Error.png")

    def get_target_item(self,data):
        if "target" in data:
            print("target",data["target"]) 
            target=data["target"]
            current_item=self.driver
            for t in target:
                current_item=self.get_local(current_item,t)
                print("current_item",current_item)
            return current_item
    

    def get_local(self,parent,target):
        print(target)
        if target == None:return None
        if "id" in target:
            return parent.find_element(By.ID,target["id"])
        elif "tag" in target:
            return parent.find_element(By.TAG_NAME,target["tag"])
        elif "key" in target:
            if target["key"] == "ENTER":
                return Keys.ENTER
        elif "class" in target:
            return parent.find_element(By.CLASS_NAME,target["class"])
        elif "link_text" in target:
            return parent.find_element(By.PARTIAL_LINK_TEXT,target["link_text"])
        elif "xpath" in target:
            return parent.find_element(By.XPATH,target["xpath"])
            
    def control(self,data):
        wait=WebDriverWait(self.driver,60)
        target_item=self.get_target_item(data)
        if data["action"] == "move":
            if "url" in data:
                print("move",data["url"])      
                self.driver.get(data["url"])  
            else:
                print("キーエラー:url")
        elif data["action"] =="move_frame":

            wait.until(expected_conditions.frame_to_be_available_and_switch_to_it(target_item))    
        elif data["action"]=="click":
            wait.until(expected_conditions.element_to_be_clickable(target_item))
            target_item.click()
        elif data["action"]=="clickKey":
            self.driver.send
        elif data["action"]=="input":
            if "text" in data:
                input_text=data["text"]
                if input_text.startswith("+"):
                    input_text=input_text[1:]
                else:
                   target_item.clear()
                target_item.send_keys(input_text)
        elif data["action"]=="move_parent_frame":
            self.driver.switch_to.parent_frame()
        elif data["action"]=="wait_all_element_load":
            WebDriverWait(self.driver, 10).until(expected_conditions.presence_of_all_elements_located)
        
        elif data["action"]=="mouse_over":
            actions = ActionChains(self.driver)
            actions.move_to_element(target_item)
            actions.perform()
        elif data["action"]=="sleep":
            time.sleep(data["value"])
        elif data["action"]=="screen_shot":
            print(target_item,"screen")
            if target_item is not None:
                wait.until(expected_conditions.element_to_be_clickable(target_item))
            # get width and height of the page
            w = self.driver.execute_script("return document.body.scrollWidth;")
            h = self.driver.execute_script("return document.body.scrollHeight;")

            # set window size
            self.driver.set_window_size(w,h)
            dt_now=datetime.datetime.now()
            # Get Screen Shot
            filePath=os.path.join(os.path.dirname(os.path.abspath(__file__)),"img",dt_now.isoformat()+".png")
            self.driver.save_screenshot(data["target"][0]["id"]+".png")

            print(filePath)
        elif data["action"]=="test_json":
            self.test()

    def test(self):
        path="functest/TestData.xlsx"

        success_img_folder="functest/result/success"
        error_img_folder="functest/result/error"
        reset_folder(success_img_folder)
        reset_folder(error_img_folder)
        
        target_xpath="/html/body/form/div[3]/div[2]/div/div/div/div/div[1]/div[1]/textarea"
        result_xpath="/html/body/form/div[3]/div[2]/div/div/div/div/div[1]/div[2]/textarea[1]"


        # Excelファイルからデータを読み込む
        df = pd.read_excel(path)
        



        # 行ごとにデータを取得して表示
        for index, row in df.iterrows():     
            if row["Flg"]==1: continue      
            try:
                target=self.driver.find_element(By.XPATH,target_xpath)
                exe_button=self.driver.find_element(By.XPATH,"/html/body/form/div[3]/div[2]/div/div/div/div/div[1]/div[1]/div/input[1]")
                result_text=self.driver.find_element(By.XPATH,result_xpath)
            
                target.clear()
                result_text.clear()
                
                target.send_keys(row['Definition'])
                ##df.at[index,"Result"]=str(index)+"sasdasdasdadas"
                exe_button.click()

                WebDriverWait(self.driver, 10).until(
                    lambda driver: driver.find_element(By.XPATH,result_xpath).get_attribute('value') != ""
                )  
                result=self.driver.find_element(By.XPATH, result_xpath).get_attribute("value")
                df.at[index,"Result"]=result
                self.driver.save_screenshot(success_img_folder+"/test"+"_"+str(index)+".png")
                if result != row['Expect']:
                    self.driver.save_screenshot(error_img_folder+"/test"+"_error_notequal_"+str(index)+".png")
            except TimeoutError:
                self.driver.save_screenshot(error_img_folder+"/test"+"_error_timeout_"+str(index)+".png")
                
            except  Exception as e:
                print(e)
                self.driver.save_screenshot(error_img_folder+"/test"+"_error_"+str(index)+".png")

            finally:
                modals=self.driver.find_elements(By.CLASS_NAME, "ui-dialog")
               
                if(len(modals)>0):
                    self.driver.find_element(By.XPATH,"/html/body/div[1]/div[3]/div/button").click()
                
        df.to_excel(path, index=False)
