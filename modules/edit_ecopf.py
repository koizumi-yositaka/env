from modules.action_by_json import ControlDisplay
def main():
    cd=ControlDisplay("Chrome")
    cd.read_scenario("testdata/demo.json")

