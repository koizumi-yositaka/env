from modules.action_by_json import *
def main():
    env={
        "url":"http://localhost/service/eco-pf/webapp/desolution/login.aspx",
        "user":"u011",
        "pass":"test"
    }
    #test_dataは配列でテストデータをループしていく
    cd=ControlDisplay("Chrome",env)
    files=os.listdir(scenario_folder)
    for file_name in files:
        cd.read_scenario(
            file_name[:file_name.index(".")],["test1","test2"]
        )
    

