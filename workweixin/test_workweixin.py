import shelve
from time import sleep
import allure
import pytest
from workweixin.base import Base

def add_cookies(self):
    self.driver.get("https://work.weixin.qq.com/wework_admin/frame#index")
    db = shelve.open("cookies")
    cookies=db['cookie']
    db.close()
    for cookie in cookies:
        self.driver.add_cookie(cookie)
        print(cookie)
    self.driver.refresh()
    sleep(3)
"""
获取上传文件按钮
"""
def up_files(self):
    btn =self.driver.find_element_by_xpath('//*[@class="ww_fileImporter_fileContainer_uploadInputMask"]')
    return btn
"""
获取上传的文件名
"""
def get_filename(self):
    filename = self.driver.find_element_by_css_selector(".ww_fileImporter_fileContainer_fileNames").text
    return filename

@allure.feature("导入联系人")
class TestImport(Base):
    @pytest.mark.skip
    @allure.story("获取cookie")
    def test_get_cookies(self):
        self.driver.get('https://work.weixin.qq.com/wework_admin/frame#index')
        cookies = self.driver.get_cookies()
        print(cookies)
        sleep(3)
        db=shelve.open("cookies")
        db['cookie'] = cookies
        db.close()

    @allure.story("成功case")
    @pytest.mark.run(order=1)
    def testimport_success(self):
        add_cookies(self)
        sleep(3)
        btn_import = self.driver.execute_script('return document.getElementsByClassName("index_service_cnt_itemWrap")[1]')
        btn_import.click()
        up_files(self).send_keys('E:\Selenium_HomeWork\workweixin\qqq.xlsx')
        sleep(3)
        filename = get_filename(self)
        assert "qqq.xlsx" == filename
        sleep(3)
        btn_submit=self.driver.find_element_by_xpath('//*[@class="qui_btn ww_btn ww_btn_Large ww_btn_Blue ww_fileImporter_submit"]')
        btn_submit.click()
        sleep(10)

        text = self.driver.find_element_by_xpath('//*[@class="ww_fileImporter_successImportText"]').text
        assert "导入成功1人"==text


    @allure.story("批量导入模板错误")
    @pytest.mark.run(order=2)
    def test_import_fail1(self):
        add_cookies(self)
        btn_import = self.driver.execute_script('return document.getElementsByClassName("index_service_cnt_itemWrap")[1]')
        btn_import.click()
        up_files(self).send_keys('E:\Selenium_HomeWork\workweixin\qq.xls')
        sleep(3)
        filename = get_filename(self)
        assert "qq.xls" == filename
        sleep(3)
        btn_submit=self.driver.find_element_by_xpath('//*[@class="qui_btn ww_btn ww_btn_Large ww_btn_Blue ww_fileImporter_submit"]')
        btn_submit.click()
        sleep(10)
        text = self.driver.find_element_by_xpath('//*[@class="ww_fileImporter_errorTitle"]').text
        assert "批量导入模板错误" == text

    @allure.story("无变化1人")
    @pytest.mark.run(order=3)
    def test_import_fail2(self):
        add_cookies(self)
        btn_import = self.driver.execute_script('return document.getElementsByClassName("index_service_cnt_itemWrap")[1]')
        btn_import.click()
        up_files(self).send_keys('E:\Selenium_HomeWork\workweixin\qqq.xlsx')
        sleep(3)
        filename = get_filename(self)
        assert "qqq.xlsx" == filename
        sleep(3)
        btn_submit=self.driver.find_element_by_xpath('//*[@class="qui_btn ww_btn ww_btn_Large ww_btn_Blue ww_fileImporter_submit"]')
        btn_submit.click()
        sleep(10)
        text = self.driver.find_element_by_xpath('//*[@class="ww_fileImporter_successImportText"]').text
        assert "无变化1人" == text