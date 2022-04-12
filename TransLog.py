from dateutil.parser import parse
import json
import copy
import os


# 输入时间格式
# a = parse('2019-10-30 23:43:10.123')
# b = parse("2019-10-28/09:08:13.56212")
#
# (a - b).days  # 获取天数的时间差
# (a - b).seconds  # 获取时间差中的秒数，也就是23:43:10到09:08:13，不包括前面的天数和秒后面的小数
# (a - b).total_seconds()  # 包括天数，小时，微秒等在内的所有秒数差
# (a - b).microseconds  # 秒小数点后面的差值

def t2ms(strTime):
    h, m, s, ms = strTime.strip().split(':')
    print(h, m, s, '0.' + ms)
    return ((int(h) * 3600 + int(m) * 60 + int(s)) + float('0.' + ms)) * 1000


def Replace_all_Ditincte(line, strFind):
    if strFind not in line:
        return line

    find_pos = line.find(strFind) + len(strFind) + 1
    find_r = line.find(")", find_pos)
    source_str = line[find_pos: find_r]
    if not DN_dict.get(source_str):
        return line
    transStr = DN_dict[source_str]
    if strFind == "nMainCmd" or strFind == "nSubCmd":
        transStr = "SC" + transStr[2:]

    line = line[:find_pos] + transStr + line[find_r:]

    # {"traceEvents": [
    #  { "ts":0, "tid":"TCP - iMainCmd", "name":"CS_LOGIN", "ph": "X","dur": 1000000,"pid":1 },
    findPos = line.find("'")
    currTime = t2ms(line[findPos + 1: line.find("'", findPos + 1)])
    # currTime = parse(line[findPos + 1: line.find("'", findPos + 1)])
    """这里用换算得到的时间单位为秒，乘以100万得到毫秒ms"""
    # DN_subJson["ts"] = (currTime - log_startTime).total_seconds() * 1000000
    print("开始时间：" + str(log_startTime))
    DN_subJson["ts"] = currTime - log_startTime
    print(DN_subJson["ts"])
    print("当前时间：" + str(currTime))
    DN_subJson["name"] = transStr
    DN_subJson["tid"] = strFind
    DN_subJson["dur"] = 100
    DN_subJson["args"]["ms"] += 1
    DN_subJson["args"]["阶段"] = "登录阶段"

    DN_json["traceEvents"].append(copy.deepcopy(DN_subJson))
    return line


def GetInfo(dn_path):
    if not dn_path:
        print("路径为空")
        return
    if not os.path.exists(dn_path):
        print("文件不存在，请重新输入")

    foo = open(dn_path)
    print(foo.name)
    coutReport = open(r"E:\Desktop\Test\TransferLog\x64\Debug\DN.log", "w")
    jsonReport = open(r"E:\Desktop\Test\TransferLog\x64\Debug\DN.json", "w")

    line = foo.readline()
    findPos = line.find("'")
    global log_startTime
    log_startTime = t2ms(line[findPos + 1: line.find("'", findPos + 1)])
    print(log_startTime)
    foo.seek(0, 0)
    for line in foo:
        line = Replace_all_Ditincte(line, "TCP - iMainCmd")
        line = Replace_all_Ditincte(line, "iSubCmd")
        line = Replace_all_Ditincte(line, "nMainCmd")
        line = Replace_all_Ditincte(line, "nSubCmd")
        coutReport.write(line)
    jsonReport.write(json.dumps(DN_json, sort_keys=False, indent=4))

    coutReport.close()
    jsonReport.close()
    foo.close()
    return


Datas = [
    "CS_LOGIN",
    "CS_SYSTEM",
    "CS_CHAR",
]


def loopAssignDict():
    i = 1
    for val in Datas:
        DN_dict[str(i)] = val
        i += 1


# 数据定义
DN_path = r"E:\Desktop\Test\TransferLog\x64\Debug\DragonNest.log"
DN_dict = {}
loopAssignDict()
DN_json = {"traceEvents": []}
DN_subJson = {"ts": 0.0, "tid": "TCP - iMainCmd", "name": "CS_LOGIN", "ph": "X", "dur": 1000000, "pid": 1,
              "args": {"ms": 121.6}}
log_startTime = 0

if __name__ == '__main__':
    # print(os.getcwd())
    # # DN_path = input("请输入文件路径：")
    # fn = input("请输入文件名字：")
    # DN_path = os.getcwd() + fn

    GetInfo(DN_path)

    # os.system("pause")
