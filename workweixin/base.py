from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


class Base():
    def setup(self):
        # browser = os.getenv("browser")
        # if browser =='firefox':
        #     self.driver = webdriver.Firefox()
        # elif browser == 'headless':
        #     self.driver = webdriver.PhantomJS()
        # else:
        #     self.driver == webdriver.Chrome()
        #获取cookie
        # options = Options()
        # options.debugger_address = "127.0.0.1:9222"
        # self.driver = webdriver.Chrome(options=options)
        self.driver = webdriver.Chrome()
        self.driver.implicitly_wait(3)
        self.driver.maximize_window()

    def teardown(self):
        self.driver.quit()


