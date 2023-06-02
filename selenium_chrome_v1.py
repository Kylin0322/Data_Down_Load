from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
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

        # self.driver = webdriver.Chrome(chrome_options=self.option)  # Chrome浏览器配置
        self.driver = webdriver.Chrome(options=self.option)  # 打开Chrome浏览器UI界面配置

        self.web_address="http://cnhuam0rptq01/reports/report/QM2%20Customized%20Reports/Defect%20Report_mirror"
        #self.web_address ="http://cnhuam0rptq01/ReportServer/Pages/ReportViewer.aspx?%2FQM2%20Customized%20Reports%2FDefect%20Report_mirror&rc:showbackbutton=true"
        

        

    # 使用谷歌浏览器模拟执行
    def open_web(self):

        print('打开网页')
        self.driver.get(self.web_address)  # 打开url网页
        time.sleep(60)

        iframe=self.driver.find_element(By.CLASS_NAME, 'viewer')
        self.driver.switch_to.frame(iframe)
        # select customer
        customer_select= Select(self.driver.find_element(By.ID, 'ReportViewerControl_ctl04_ctl03_ddValue'))
        # customer_select.select_by_value("HP/Printer Motherboard")
        customer_select.select_by_index(57)
        time.sleep(60)

        # input start date
        start_time=self.driver.find_element(By.ID,"ReportViewerControl_ctl04_ctl09_txtValue")
        start_time.clear()
        date=datetime.datetime.now()+datetime.timedelta(days=-1)
        start_time.send_keys(date.strftime("%Y-%m-%d 07:00:00"))

        # input end date
        end_time=self.driver.find_element(By.ID,"ReportViewerControl_ctl04_ctl11_txtValue")
        end_time.clear()
        end_time.send_keys(datetime.datetime.now().strftime("%Y-%m-%d 07:00:00"))

        # click View Report
        view_report = self.driver.find_element(By.ID, "ReportViewerControl_ctl04_ctl00")
        view_report.click()


if __name__ == '__main__':
    web=webdriver_chrome()
    web.init()
    web.open_web()
    
  
    