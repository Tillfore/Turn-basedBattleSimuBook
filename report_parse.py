import json


def get_report():
    reportFile = open("battleoutput.json")
    reportString = reportFile.read()
    battleReport = json.loads(reportString)
    print('report的类型为:', type(battleReport))
    print('report的值为:', battleReport)


if __name__ == '__main__':
    get_report()
