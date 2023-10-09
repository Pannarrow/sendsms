#群发短信，读取表格中电话，信息请放在第一张表格（默认是Sheet1），列名须拥有“姓名”和“电话”
import os 
import pandas as pd
import time
import subprocess


#坐标点击
def ClickScreen(x,y):
    os.system(f'adb shell input tap {x} {y}')

#发送短信
def SendMessage(mobile,content):
    #向mobile编写content
    os.popen("adb shell am start -a android.intent.action.SENDTO -d sms:{} --es sms_body {}".format(mobile,content))
    time.sleep(3)
    #点击发送
    ClickScreen(970,1120)   #按钮坐标请写到这儿

def main():

    #激活adb接口
    adb_check=subprocess.Popen("adb nodaemon server",shell=True)

    #查询adb接口
    print("正在检查adb接口......")
    result=os.popen("adb devices")
    text=result.readlines()
    if text[1]=='\n':
        print("未检测到adb接口，请检查手机USB调试状态")
        os.system("pause")
    else:
        print(text[1])

        #读取模板
        fr = open('template.txt','r', encoding='utf-8')
        model = fr.read()
        fr.close()
        # model="$Name你好，请于$Time到楼上报道"
        if model is None:
            print("模板不能为空")
            return
        
        #匹配模板
        if "$Name" in model:
            nameIndex = model.find("$Name")
        elif "$name" in model:
            nameIndex = model.find("$name")
        if "$Time" in model:
            timeIndex = model.find("$Time")
        elif "$time" in model:
            timeIndex = model.find("$time")
        if nameIndex > 0:
            head=model[0:nameIndex]
        else:
            head=""
        tail=model[nameIndex+5:timeIndex]
        end=model[timeIndex+5:]
        content=""
        mobile=""

        #导入表格
        df = pd.read_excel('list.xlsx', sheet_name=0)

        #遍历表格，读取电话
        for index,row in df.iterrows():
            mobile=str(row['电话'])
            content=head+str(row['姓名'])+tail+str(row['时间'])+end
            SendMessage(mobile,content)
            content=""
    time.sleep(5)
    adb_check.kill()

if __name__ == '__main__':
    main()