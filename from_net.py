import re
import time
from urllib.request import Request, urlopen
import ddddocr
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from extract_color import extract_red, extract_color
from selenium.common.exceptions import NoSuchElementException


class InvoiceAuth:
    def __init__(self):
        self.certification = None
        # 加载配置
        chrome_options = Options()
        # 不打开浏览器
        chrome_options.add_argument('--headless')
        chrome_options.add_argument(
            'user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_0) AppleWebKit/537.36 (KHTML, like Gecko)'
            ' Chrome/84.0.4147.89 Safari/537.36')
        # 取消证书验证
        chrome_options.add_argument('--ignore-certificate-errors')
        self.driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)

        # 最大化窗口
        self.driver.maximize_window()

        # ocr对象
        self.ocr = ddddocr.DdddOcr(show_ad=False)

        # 打开目标网页
        self.driver.get("https://inv-veri.chinatax.gov.cn/")

    def auth_from_net(self, invoice_code, invoice_number, invoice_time, check_code, file_name):

        # 输入发票号码
        res = None
        self.driver.find_element(By.XPATH, '//table[@class="comm_table2 fr"]/tbody/tr[2]/td/input').send_keys(
            invoice_code)

        # 输入发票代码
        self.driver.find_element(By.XPATH, '//table[@class="comm_table2 fr"]/tbody/tr[3]/td/input').send_keys(
            invoice_number)

        # 输入开票日期
        self.driver.find_element(By.XPATH, '//table[@class="comm_table2 fr"]/tbody/tr[4]/td//input').send_keys(
            invoice_time)

        # 输入校验码
        self.driver.find_element(By.XPATH, '//table[@class="comm_table2 fr"]/tbody/tr[5]/td/input').send_keys(
            check_code)
        try:
            if self.driver.find_element(By.XPATH, '//div[@id="popup_message"]').text == '验证码请求次数过于频繁，请1分钟后再试！':
                time.sleep(100)
            self.refresh_certification()
        except NoSuchElementException as e:
            # print(e)
            pass
        self.check_send()
        # 如果验证码错误
        while True:
            try:
                if self.driver.find_element(By.XPATH, '//div[@id="popup_message"]').text in ['验证码失效!', '验证码错误!']:
                    time.sleep(1)
                    self.refresh_certification()
                else:
                    break
            except NoSuchElementException as e:
                break
            try:
                if self.driver.find_element(By.XPATH, '//div[@id="popup_message"]').text == '验证码请求次数过于频繁，请1分钟后再试！':
                    time.sleep(100)
                self.refresh_certification()
            except NoSuchElementException as e:
                # print(e)
                pass
            self.check_send()

        try:
            if self.driver.find_element(By.XPATH, '//div[@id="popup_message"]').text == '超过该张发票当日查验次数(请于次日再次查验)!':
                # 点击确定
                self.driver.find_element(By.XPATH, '//div[@id="popup_panel"]/input').click()
        except NoSuchElementException:
            # 获取发票结果
            res = self.get_data()
            # print(res)
            # 截图
            self.driver.save_screenshot('./result/{}.png'.format(file_name))
        self.driver.refresh()
        return res

    def check_send(self):
        # 提取问题，蓝色、红色、全部验证码
        time.sleep(3)
        question = self.driver.find_element(By.XPATH,
                                            '//table[@class="comm_table2 fr"]/tbody/tr[6]/td[@id="yzminfo"]').text
        # 获取验证码图片
        img_url = self.driver.find_element(By.XPATH,
                                           '//table[@class="comm_table2 fr"]/tbody/tr[7]//a//img').get_attribute('src')
        # print(img_url)
        # 存储验证码图片
        request = Request(img_url, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                          '(KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36'})

        data = urlopen(request).read()
        with open('1.jpg', 'wb') as f:
            f.write(data)

        if "红色" in question:
            extract_red("1.jpg", "1.jpg")
        elif "蓝色" in question:
            extract_color("1.jpg", "1.jpg", 'blue')
        elif "黄色" in question:
            extract_color("1.jpg", "1.jpg", 'yellow')
        elif "绿色" in question:
            extract_color("1.jpg", "1.jpg", 'green')

        # 识别验证码
        with open("1.jpg", 'rb') as f:
            image = f.read()
        self.certification = self.ocr.classification(image)
        # self.certification = input()
        # 清空输入栏
        self.driver.find_element(By.XPATH, '//table[@class="comm_table2 fr"]/tbody/tr[6]/td/input').clear()
        time.sleep(1)
        # 输入验证码
        self.driver.find_element(By.XPATH, '//table[@class="comm_table2 fr"]/tbody/tr[6]/td/input').send_keys(
            self.certification if self.certification else '123')
        time.sleep(1)
        # 点击查验
        try:
            self.driver.find_element(By.XPATH, '//table[@class="comm_table2 fr"]/tbody/tr[8]/td/div/button[2]').click()
        except NoSuchElementException:
            pass
        time.sleep(5)

    # 点击刷新验证码
    def refresh_certification(self):
        # 点击确定
        self.driver.find_element(By.XPATH, '//div[@id="popup_panel"]/input').click()
        time.sleep(2)
        # 点击验证码刷新
        self.driver.find_element(By.XPATH, '//td[@id="imgarea"]/div/a/img').click()
        time.sleep(3)

    # 获取json信息
    def get_data(self):
        self.driver.switch_to.frame(self.driver.find_elements(By.TAG_NAME, "iframe")[0])
        # print(self.driver.page_source)
        try:
            if self.driver.find_element(By.XPATH, '//tr[@class="result_title"]//strong').text == '不一致':
                res = {
                    '查验时间': [self.driver.find_element(By.XPATH, '//tr[@class="result_title"]//span[@id="cysj"]').text[
                             5:]],
                    '查验结果': [self.driver.find_element(By.XPATH, '//tr[@class="result_title"]//strong').text],
                    '发票名称': [self.driver.find_element(By.XPATH, '//div[@id="tabPage2"]/h1').text],
                    '发票代码': [self.driver.find_element(By.XPATH, '//div[@id="tabPage2"]/table/tbody/tr[1]/td[2]').text],
                    '发票号码': [self.driver.find_element(By.XPATH, '//div[@id="tabPage2"]/table/tbody/tr[2]/td[2]').text],
                    '开票日期': [self.driver.find_element(By.XPATH, '//div[@id="tabPage2"]/table/tbody/tr[3]/td[2]').text],
                    '校验码': [self.driver.find_element(By.XPATH, '//div[@id="tabPage2"]/table/tbody/tr[4]/td[2]').text]}
                return res
        except NoSuchElementException:
            res = {
                '查验时间': [
                    self.driver.find_element(By.XPATH, '//table[@class="comm_table2"]//span[@id="cysj"]').text[5:]],
                '查验次数': [
                    self.driver.find_element(By.XPATH, '//table[@class="comm_table2"]//span[@id="cycs"]').text[5:]],
                '发票名称': [self.driver.find_element(By.XPATH, '//div[@class="tab-page"]//h1').text],
                '发票代码': [self.driver.find_element(By.XPATH,
                                                  '//div[@class="tab-page"]/table/tbody/tr/td[1]/span').text],
                '发票号码': [self.driver.find_element(By.XPATH,
                                                  '//div[@class="tab-page"]/table/tbody/tr/td[3]/span').text],
                '开票日期': [self.driver.find_element(By.XPATH,
                                                  '//div[@class="tab-page"]/table/tbody/tr/td[5]/span').text],
                '校验码': [self.driver.find_element(By.XPATH,
                                                 '//div[@class="tab-page"]/table/tbody/tr/td[7]/span').text],
                '机器编码': [self.driver.find_element(By.XPATH,
                                                  '//div[@class="tab-page"]/table/tbody/tr/td[9]/span').text],
                '合计金额': [self.driver.find_element(By.XPATH,
                                                  '//span[@id = "je_dzfp"]').text],
                '合计税额': [self.driver.find_element(By.XPATH,
                                                  '//span[@id = "se_dzfp"]').text],
                '价税合计': [(self.driver.find_element(By.XPATH,
                                                   '//span[@id = "jshjdx_dzfp"]').text
                          + '（小写）'
                          + self.driver.find_element(By.XPATH,
                                                     '//span[2][@id = "jshjxx_dzfp"]').text)],
                '购买方名称': [self.driver.find_element(By.XPATH,
                                                   '//span[@id = "gfmc_dzfp"]').text],
                '购买方纳税人识别号': [self.driver.find_element(By.XPATH,
                                                       '//span[@id = "gfsbh_dzfp"]').text],
                '购买方地址、电话': [self.driver.find_element(By.XPATH,
                                                      '//div[@id = "tabPage-dzfp"]/table[2]/tbody/tr[3]/td[2]').text],
                '购买方开户行及账号': [self.driver.find_element(By.XPATH,
                                                       '//div[@id = "tabPage-dzfp"]/table[2]/tbody/tr[4]/td[2]').text],
                '销售方名称': [self.driver.find_element(By.XPATH,
                                                   '//span[@id = "xfmc_dzfp"]').text],
                '销售方纳税人识别号': [self.driver.find_element(By.XPATH,
                                                       '//span[@id = "xfsbh_dzfp"]').text],
                '销售方地址、电话': [self.driver.find_element(By.XPATH,
                                                      '//span[@id = "xfdzdh_dzfp"]').text],
                '销售方开户行及账号': [self.driver.find_element(By.XPATH,
                                                       '//span[@id = "xfyhzh_dzfp"]').text],
                '货物或应税劳务、服务名称': [i.text for i in self.driver.find_elements(
                    By.XPATH, '//td[1][@class="align_left borderRight"]/span')],
                '规格型号': [i.text for i in self.driver.find_elements(By.XPATH,
                                                                   '//td[2][@class="align_left borderRight"]/span')],
                '单位': [i.text for i in self.driver.find_elements(By.XPATH,
                                                                 '//td[3][@class="align_left borderRight"]/span')],
                '数量': [i.text for i in self.driver.find_elements(By.XPATH,
                                                                 '//td[4][@class="align_right borderRight"]/span')],
                '单价': [i.text for i in self.driver.find_elements(By.XPATH,
                                                                 '//td[5][@class="align_right borderRight"]/span')],
                '金额': [i.text for i in self.driver.find_elements(
                    By.XPATH, '//td[6][@class="align_right borderRight"]/span')[:-1]],
                '税率': [i.text for i in self.driver.find_elements(By.XPATH,
                                                                 '//td[7][@class="align_right borderRight"]/span')],
                '税额': [i.text for i in self.driver.find_elements(
                    By.XPATH, '//div[@id = "tabPage-dzfp"]//td[8][@class="align_right"]/span')[:-1]]}
            return res


def excel2png(file_path):
    driver = InvoiceAuth()
    res_l = []
    df = pd.read_excel(file_path, sheet_name='Sheet1', dtype=str)
    for i in range(2):
        res = driver.auth_from_net(invoice_code=df.iloc[i, 3], invoice_number=df.iloc[i, 4],
                                   invoice_time=''.join(re.findall('\d', df.iloc[i, 5])),
                                   check_code=df.iloc[i, 6][-6:], file_name=df.iloc[i, 1])
        # 全为空则值为空，反之用,连接成字符串
        for k in res.keys():
            flag = 0
            for v in res[k]:
                if v:
                    flag = 1
                    break
            if flag:
                res[k] = ','.join(res[k])
            else:
                res[k] = ''
        res_df = pd.DataFrame([res])
        res_l.append(res_df)
    driver.driver.quit()
    res_l_df = pd.concat(res_l, ignore_index=True)
    res_l_df.to_excel('result.xlsx', index=False)



