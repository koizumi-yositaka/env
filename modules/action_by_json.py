import chromedriver_binary # nopa
import json
import time
import datetime
import os
import glob
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
timeout =10
def reset_folder(folder_path):
    files=glob.glob(os.path.join(folder_path,"*"))
    for file in files:
        try:
            os.remove(file)
            print(f'{file} を削除しました。')
        except Exception as e:
            print(f'{file} の削除中にエラーが発生しました: {e}')



class ControlDisplay(object):
    def __init__(self,type,env):
        if type=="Chrome":
            options = webdriver.ChromeOptions()
            
            #options.add_argument('--headless')
            
            print('connectiong to remote browser...')
            self.driver = webdriver.Chrome(options=options)
            self.driver.maximize_window()
            self.env=env
    def get_parameter(self,value):
        print("get_parameter",value)
        pattern = r'\${{(.*?)}}'
        matches = re.findall(pattern, value)
        if matches:
            print("Matches found:")
            for match in matches:
                print(match)
                
                params=match.split(".")
                result=None
                for index,param in enumerate(params):
                    print(index,param)
                    if index==0:
                        if param=="env":
                            #環境変数
                            result=self.env
                        elif param=="arg":
                            #引数
                            result=self.arg
                        else:
                            result=params[0]

                    else:
                        print(param in result)
                        #さらに
                        if param in result and result:
                            result=result[param]
                        else:
                            print(f"{param}が見つかりませんでした")
                if isinstance(result,str) :
                    value = value.replace("${{"+match+"}}",result)
                else:
                    value=result
                print(value)
            return value
                
        else:
            print("No matches found.")
            return value     

    def read_scenario(self,scenarios):
        for scenario in scenarios:

            with open(scenario["file"],"r",encoding="UTF-8") as j:
                data=json.load(j)
                try:
                    for action in data["Input"]:
                        self.arg=scenario["arg"]
                        self.control(action)
                        print(action,"完了")
                   
                except Exception as e:

                    print("Error",e)
                    dt_now=datetime.datetime.now()
                    self.driver.save_screenshot("Error.png")
                    self.driver.quit()
        self.driver.quit()
                     

    def get_target_item(self,data):
        if "target" in data:
            print("target",data["target"]) 
            target=data["target"]
            current_item=self.driver
            for t in target:
                current_item=self.get_local(current_item,t)
            return current_item
    

    def get_local(self,parent,target):
        if target == None:return None
        if "id" in target:
            return WebDriverWait(parent, timeout).until(
                lambda d: d.find_element(By.ID, self.get_parameter(target["id"]))
            )
        elif "tag" in target:
            return WebDriverWait(parent, timeout).until(
                lambda d: d.find_element(By.TAG_NAME, self.get_parameter(target["tag"]))
            )
        elif "key" in target:
            if target["key"] == "ENTER":
                return Keys.ENTER
        elif "class" in target:
            return WebDriverWait(parent, timeout).until(
                lambda d: d.find_element(By.CLASS_NAME, self.get_parameter(target["class"]))
            )
        elif "link_text" in target:
            return WebDriverWait(parent, timeout).until(
                lambda d: d.find_element(By.PARTIAL_LINK_TEXT, self.get_parameter(target["link_text"]))
            )
        elif "xpath" in target:
            return WebDriverWait(parent, timeout).until(
                lambda d: d.find_element(By.XPATH, self.get_parameter(target["xpath"]))
            )
            
    def control(self,data):
        wait=WebDriverWait(self.driver,60)
        target_item=self.get_target_item(data)

        if data["action"] == "move":
            if "url" in data:
                print(data["url"])
                self.driver.get(self.get_parameter(data["url"]))  
            else:
                print("キーエラー:url")
        elif data["action"] =="move_frame":
            if target_item is None:
                print("targetが見つかりません",data) 
                return
            wait.until(expected_conditions.frame_to_be_available_and_switch_to_it(target_item))     # type: ignore
        elif data["action"] =="move_display_frame":
            if target_item is None:
                print("targetが見つかりません",data) 
                return
            print("target_item",target_item)
            iframes = target_item.find_elements(By.TAG_NAME, "iframe")
            for iframe in iframes:
                print(iframe.get_attribute("id"))
                if iframe.get_attribute("style") and "display: none" not in iframe.get_attribute("style"):
                    print("iframe",iframe)
                    #wait.until(expected_conditions.frame_to_be_available_and_switch_to_it(iframe))     # type: ignore
                    self.driver.switch_to.frame(iframe)
                    print("Switched to the iframe.")
                    break
            
        elif data["action"]=="click":
            if target_item is None:
                print("targetが見つかりません",data) 
                return
            wait.until(expected_conditions.element_to_be_clickable(target_item))# type: ignore
            target_item.click()

        elif data["action"]=="input":
            if target_item is None:
                print("targetが見つかりません",data) 
                return
            if "text" in data:
                input_text=self.get_parameter(data["text"])
                if input_text.startswith("+"):
                    input_text=input_text[1:]
                else:
                   target_item.clear()
                target_item.send_keys(input_text)
        elif data["action"]=="move_parent_frame":
            print("parentframe")
            self.driver.switch_to.parent_frame()
        elif data["action"]=="wait_all_element_load":
            WebDriverWait(self.driver, 10).until(expected_conditions.presence_of_all_elements_located)
        
        elif data["action"]=="mouse_over":
            if target_item is None:
                print("targetが見つかりません",data) 
                return
            self.driver.execute_script('arguments[0].scrollIntoView(false);', target_item)
            actions = ActionChains(self.driver)    
            actions.move_to_element(target_item)
            actions.perform()
        elif data["action"]=="sleep":
            time.sleep(data["value"])
        elif data["action"]=="screen_shot":
            if target_item is None:
                print("targetが見つかりません",data) 
                return            
            wait.until(expected_conditions.element_to_be_clickable(target_item))
            dt_now=datetime.datetime.now()
            # Get Screen Shot
            filePath=os.path.join(os.path.dirname(os.path.abspath(__file__)),"img",dt_now.isoformat()+".png")
            self.driver.save_screenshot(data["target"][0]["id"]+".png")
        elif data["action"]=="test_json":
            self.test()
        elif data["action"]=="test_input":
            if "inputdata" not in data:
                print("inputdataがありません")
                return
            input_data=self.get_parameter(data["inputdata"])
            for input in input_data:
                elem=None
                value=""
                if "value" in input:
                    value=input["value"]
                
                if input["type"]=="text":
                    elem=self.driver.find_element(By.ID, input["id"])
                    elem.send_keys(value)
                elif input["type"]=="radio":
                    name=input["id"]
                    
                    elem = WebDriverWait(self.driver, 10).until(
                        expected_conditions.presence_of_element_located((By.XPATH, f"//input[@name='{name}' and @value='{value}']"))
                    )
                    if elem:
                        elem.click()
                    else:
                        print("ERROR ラジオボタンが見つかりません")
                        return
                    
                elif input["type"]=="button":
                    ##wait.until(expected_conditions.element_to_be_clickable(target_item))# type: ignore
                    elem=self.driver.find_element(By.ID, input["id"])
                    if elem is None:
                        print("ボタンがありません")
                        return
                    elem.click() 
                else:
                    elem=self.driver.find_element(By.ID, input["id"])
                    elem.send_keys(value)

                    
        
            
    def test(self):
        path="functest/TestData.xlsx"
        now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path=f"functest/result/output-excel/result-{now}.xlsx"
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
                
        df.to_excel(output_path, index=False)
