from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
import datetime
from apscheduler.schedulers.asyncio import AsyncIOScheduler
try:
    import asyncio
except ImportError:
    import trollius as asyncio

option = webdriver.ChromeOptions()
# 保持浏览器界面不关闭
# option.add_experimental_option("detach", True)
# headless 模式 不打开UI界面的情况下使用 Chrome 浏览器
option.add_argument('headless')
# 谷歌文档提到需要加上这个属性来规避bug
option.add_argument('--disable-gpu')
# 不加载图片, 提升速度
option.add_argument('blink-settings=imagesEnabled=false')
# 以最高权限运行
option.add_argument('--no-sandbox')

# driver = webdriver.Chrome(chrome_options=option)  # Chrome浏览器配置
driver = webdriver.Chrome(options=option)  # 不打开Chrome浏览器UI界面配置


# 使用谷歌浏览器模拟执行
def run_driver():
    # 检查是否存在sn序列号
    sn = ""
    sn_list = []
    with open('./sn.txt', "r") as f:
        sn_list = f.readlines()
        list_len = len(sn_list)
        if list_len == 0:
            return False
        if list_len > 0:
            sn = sn_list[0]
        if list_len == 1:
            sn_list = []
        if list_len > 1:
            sn_list = sn_list[1:]

    # 打开网页
    driver.get("http://10.114.26.174:8002/booking.aspx")  # 打开url网页
    # input username
    TextBox3 = driver.find_element(By.NAME, "TextBox3")
    TextBox3.clear()  # 清除文本
    TextBox3.send_keys("Sunmoon_she@jabil.com")  # 模拟按键输入

    # inpout passwd
    TextBox4 = driver.find_element(By.NAME, "TextBox4")
    TextBox4.clear()
    TextBox4.send_keys("sunmoon")

    # click login button
    Button2 = driver.find_element(By.ID, "Button2")
    Button2.click()
    # input Booking S/N
    TextBox2 = driver.find_element(By.NAME, "TextBox2")
    TextBox2.clear()
    TextBox2.send_keys(sn)

    # select datetime
    DropDownList1 = Select(driver.find_element(By.NAME, "DropDownList1"))
    optionLen = len(DropDownList1.options)
    DropDownList1.select_by_index(optionLen-1)

    # Verification
    checkNumber = driver.find_element(By.ID, "Label4").text
    if checkNumber.find("+") >= 0:
        num1 = checkNumber.split('+')[0]
        num2 = checkNumber.split('+')[1].split('=')[0]
        TextBox5 = driver.find_element(By.NAME, "TextBox5")
        TextBox5.clear()
        TextBox5.send_keys(int(num1)+int(num2))

    # Click Booking S/N
    Button1 = driver.find_element(By.NAME, "Button1")
    Button1.click()

    # 响应文本
    res_txt = ""

    # 是否有alert告警弹窗
    alert = EC.alert_is_present()(driver)
    if alert:
        res_txt = alert.text
        alert.accept()

    # Model Number/ Message
    res_txt = res_txt + driver.find_element(By.ID, "Label1").text

    with open('./booking_log.txt', "a") as f:
        f.write(datetime.datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S")+"\t"+res_txt+"\n")

    # 判断是否预定成功,成功则从sn文件中删除
    if res_txt.find("MES Checking Succeed") >= 0:
        with open('./sn.txt', "w") as f:
            f.writelines(sn_list)
        with open('./succeed_sn.txt', "a") as f:
            f.write(datetime.datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S")+"\t"+sn+"\n")

    # 判断是否已被预定,是则从sn文件中删除
    if res_txt.find("has been booked  already") >= 0:
        with open('./sn.txt', "w") as f:
            f.writelines(sn_list)
        with open('./booked_sn.txt', "a") as f:
            f.write(datetime.datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S")+"\t"+sn+"\n")

    return False


def run_task():
    n = 0
    while n < 1000:
        begin = datetime.datetime.now()
        print("start time:", begin)
        res = run_driver()
        end = datetime.datetime.now()
        print("end time:", end)
        print("time spent", end-begin, "\n")

        if res:
            break
        else:
            n = n+1


if __name__ == '__main__':
  
    # Execution will block here until Ctrl+C (Ctrl+Break on Windows) is pressed.
    try:
        asyncio.get_event_loop().run_forever()
    except (KeyboardInterrupt, SystemExit):
        pass
