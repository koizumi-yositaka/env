from modules.action_by_json import *
def main():

    ##files=os.listdir(scenario_folder)
    with open("user/setting.json" ,'r',encoding="UTF-8") as f:
        data=json.load(f)
        env=data["env"]
        #test_dataは配列でテストデータをループしていく
        cd=ControlDisplay("Chrome",env)
        for plan in data["plans"]:
            cd.read_scenario(
                plan["scenario"],
                plan["testdatas"]
            )
    
    
    

