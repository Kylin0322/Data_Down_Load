from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import datetime
import time

class webdriver_chrome():
    def __init__(self) -> None:
        pass

    def init(self):
        self.option = webdriver.ChromeOptions()
        # 保持浏览器界面不关闭
        self.option.add_experimental_option("detach", True)
        # headless 模式 不打开UI界面的情况下使用 Chrome 浏览器
        # self.option.add_argument('headless')
        # 谷歌文档提到需要加上这个属性来规避bug
        self.option.add_argument('--disable-gpu')
        # 不加载图片, 提升速度
        self.option.add_argument('blink-settings=imagesEnabled=false')
        # 以最高权限运行
        self.option.add_argument('--no-sandbox')

        #self.option.page_load_strategy = 'normal'
        self.option.page_load_strategy = 'eager'    #WebDriver waits until DOMContentLoaded event fire is returned.
        

        # driver = webdriver.Chrome(chrome_options=option)  # Chrome浏览器配置
        
        self.driver = webdriver.Chrome(options=self.option)  # 打开Chrome浏览器UI界面配置
        

        self.web_address ="http://cnhuam0rptq01/reports/report/QM2%20Customized%20Reports/Defect%20Report_mirror"
        #self.web_address ="http://10.114.26.174:8002/"
        

    def save_page(self,page):
        with open(file="page.html",mode="w",encoding="utf-8") as f:
            f.write(page) 

    def wait_until_loaded(self,Ele="Element_to_be_found"):
        
        try:
            elem = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, Ele)) #This is a dummy element
            )
        finally:
            driver.quit()
    
    def wait_until_element_found(self,Element1="</html>",timeout=60):
        ''' 等待字符串str1在 page_source 中出现
        '''
        el=None
        t=1
        while el==None and t<timeout:
            if el==None:
                try:
                    el=self.driver.find_element(by=By.CLASS_NAME,value=Element1)
                except:
                    el=None
                finally:
                    pass

            if el==None:
                try:
                    el=self.driver.find_element(by=By.NAME,value=Element1)
                except:
                    el=None
                finally:
                    pass

            if el==None:
                try:
                    el=self.driver.find_element(by=By.ID,value=Element1)
                except:
                    el=None
                finally:
                    pass

            time.sleep(1)
            t=t+1
        return el


    # 使用谷歌浏览器模拟执行
    def open_web(self):

        print('打开网页')
        self.driver.get(self.web_address)  # 打开url网页
        self.driver.implicitly_wait(60)

        #self.driver_Org=self.driver

        # self.driver.refresh()

        print('Switch_to Frame')
        # iframe = self.driver.find_element(By.CSS_SELECTOR, "#modal > iframe")
        #iframe = self.driver.find_element(By.CLASS_NAME, "view")
        # switch to selected iframe
        self.driver.switch_to.frame(0)
        self.wait_until_element_found(Element1='ParamEntryCell')
        self.wait_until_element_found(Element1='ReportViewerControl$ctl04$ctl03$ddValue')
        time.sleep(2)


        print('Select Customer')
        TextBox3 = self.driver.find_element(By.NAME, "ReportViewerControl$ctl04$ctl03$ddValue")
        TextBox3.click()    #Ok.
        TextBox3.send_keys("HP/Printer Motherboard")



        TextBox3 = self.driver.find_element(By.NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.NAME, "xxxxxxxxxxxxx")

        TextBox3 = self.driver.find_element(By.CLASS_NAME, "ParamLabelCell")   #OK
        TextBox3 = self.driver.find_element(By.CLASS_NAME, "ParamEntryCell")   #OK
        TextBox3 = self.driver.find_element(By.CLASS_NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.CLASS_NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.CLASS_NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.CLASS_NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.CLASS_NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.CLASS_NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.CLASS_NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.CLASS_NAME, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.CLASS_NAME, "xxxxxxxxxxxxx")

        TextBox3 = self.driver.find_element(By.ID, "NavigationCorrector")   #OK
        TextBox3 = self.driver.find_element(By.ID, "NavigationCorrector_ctl00") #OK
        TextBox3 = self.driver.find_element(By.ID, "ReportViewerControl_ctl04_ctl03") #OK
        TextBox3 = self.driver.find_element(By.ID, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.ID, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.ID, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.ID, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.ID, "xxxxxxxxxxxxx")
        TextBox3 = self.driver.find_element(By.ID, "xxxxxxxxxxxxx")



  

        TextBox3.get_attribute()
        TextBox3.click()
        TextBox3.id.count()
        
        TextBox3.clear()  # 清除文本
        TextBox3.send_keys("Sunmoon_she@jabil.com")  # 模拟按键输入

        # # inpout passwd
        # TextBox4 = driver.find_element(By.NAME, "TextBox4")
        # TextBox4.clear()
        # TextBox4.send_keys("sunmoon")

        # # click login button
        # Button2 = driver.find_element(By.ID, "Button2")
        # Button2.click()
        # # input Booking S/N
        # TextBox2 = driver.find_element(By.NAME, "TextBox2")
        # TextBox2.clear()
        # TextBox2.send_keys(sn)

        # # select datetime
        # DropDownList1 = Select(driver.find_element(By.NAME, "DropDownList1"))
        # optionLen = len(DropDownList1.options)
        # DropDownList1.select_by_index(optionLen-1)

        # # Verification
        # checkNumber = driver.find_element(By.ID, "Label4").text
        # if checkNumber.find("+") >= 0:
        #     num1 = checkNumber.split('+')[0]
        #     num2 = checkNumber.split('+')[1].split('=')[0]
        #     TextBox5 = driver.find_element(By.NAME, "TextBox5")
        #     TextBox5.clear()
        #     TextBox5.send_keys(int(num1)+int(num2))

        # # Click Booking S/N
        # Button1 = driver.find_element(By.NAME, "Button1")
        # Button1.click()

        # # 响应文本
        # res_txt = ""

        # # 是否有alert告警弹窗
        # alert = EC.alert_is_present()(driver)
        # if alert:
        #     res_txt = alert.text
        #     alert.accept()

        # # Model Number/ Message
        # res_txt = res_txt + driver.find_element(By.ID, "Label1").text

        # with open('./booking_log.txt', "a") as f:
        #     f.write(datetime.datetime.now().strftime(
        #         "%Y-%m-%d %H:%M:%S")+"\t"+res_txt+"\n")

        # # 判断是否预定成功,成功则从sn文件中删除
        # if res_txt.find("MES Checking Succeed") >= 0:
        #     with open('./sn.txt', "w") as f:
        #         f.writelines(sn_list)
        #     with open('./succeed_sn.txt', "a") as f:
        #         f.write(datetime.datetime.now().strftime(
        #             "%Y-%m-%d %H:%M:%S")+"\t"+sn+"\n")

        # # 判断是否已被预定,是则从sn文件中删除
        # if res_txt.find("has been booked  already") >= 0:
        #     with open('./sn.txt', "w") as f:
        #         f.writelines(sn_list)
        #     with open('./booked_sn.txt', "a") as f:
        #         f.write(datetime.datetime.now().strftime(
        #             "%Y-%m-%d %H:%M:%S")+"\t"+sn+"\n")

        # return False
 


if __name__ == '__main__':
    web=webdriver_chrome()
    web.init()
    web.open_web()
    
  
    