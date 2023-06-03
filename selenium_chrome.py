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
        # self.option.add_argument('--disable-gpu')
        
        # 不加载图片, 提升速度
        # self.option.add_argument('blink-settings=imagesEnabled=false')
        
        # 以最高权限运行
        #self.option.add_argument('--no-sandbox')

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
    
    def wait_until_element_found(self,Element1="</html>",timeout=180):
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

    def check_element_enabled(self,value="value",timeout=180):
        ''' 检查 Element 是否存在，并且是 Enabled
        '''
        el=None
        t=1
        while el==None and t<timeout:
            if el==None:
                try:
                    el=self.driver.find_element(by=By.NAME,value=value)
                    if not el.is_enabled():
                        el=None
                except:
                    el=None
                finally:
                    pass

            if el==None:
                try:
                    el=self.driver.find_element(by=By.ID,value=value)
                    if not el.is_enabled():
                        el=None
                except:
                    el=None
                finally:
                    pass

            if el==None:
                try:
                    el=self.driver.find_element(by=By.CLASS_NAME,value=value)
                    if not el.is_enabled():
                        el=None
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
        #self.driver.implicitly_wait(60)

        #self.driver_Org=self.driver

        # self.driver.refresh()

        print('Switch_to Frame')
        # iframe = self.driver.find_element(By.CSS_SELECTOR, "#modal > iframe")
        #iframe = self.driver.find_element(By.CLASS_NAME, "view")
        # switch to selected iframe
        self.wait_until_element_found(Element1='viewer')
        self.driver.switch_to.frame(0)
        time.sleep(5)
 
        self.wait_until_element_found(Element1='ReportViewerControl$ctl04$ctl03$ddValue')
        time.sleep(2)

        print('Select Customer')
        customer_select= Select(self.driver.find_element(By.ID, 'ReportViewerControl_ctl04_ctl03_ddValue'))
        # customer_select.select_by_value("HP/Printer Motherboard")
        customer_select.select_by_index(57)
        self.check_element_enabled(value="ReportViewerControl$ctl04$ctl05$ddValue")
        time.sleep(2)

        print('Input start date')   # input start date
        start_time=self.driver.find_element(By.ID,"ReportViewerControl_ctl04_ctl09_txtValue")
        start_time.clear()
        date=datetime.datetime.now()+datetime.timedelta(days=-1)
        start_time.send_keys(date.strftime("%Y-%m-%d 07:00:00"))
        time.sleep(1)

        print('input end date') # input end date
        end_time=self.driver.find_element(By.ID,"ReportViewerControl_ctl04_ctl11_txtValue")
        end_time.clear()
        end_time.send_keys(datetime.datetime.now().strftime("%Y-%m-%d 07:00:00"))
        time.sleep(1)

        print('click View Report')  # click View Report
        view_report = self.driver.find_element(By.ID, "ReportViewerControl_ctl04_ctl00")
        view_report.click()


        print('wait webload complete')  # click View Report
        self.check_element_enabled(value="ReportViewerControl_ctl04_ctl00",timeout=300)


        print('Click save buttom')  # 
        tag_1= self.driver.find_element(By.ID, 'ReportViewerControl_ctl05_ctl04_ctl00_ButtonImg')
        tag_1.click()
        time.sleep(3)

        print('Select csv data to download')  # 
        tag2= self.driver.find_element(By.XPATH, "//a[@title='CSV (comma delimited)']")
        tag2.click()

        print('wait save complete')  # click View Report
        self.check_element_enabled(value="ReportViewerControl_ctl04_ctl00",timeout=300)

        print('wait save complete')
        time.sleep(10)

        print('close chrome and exit')
        self.driver.close()




if __name__ == '__main__':
    web=webdriver_chrome()
    web.init()
    web.open_web()
    
  
    