import json
import logging
import requests
import time
from selenium import webdriver
from selenium.common.exceptions import *
from selenium.webdriver.common.keys import Keys


class _Params:
    url = "https://mail.qq.com/"
    attention_mail = [
        "小米科技",
        "Xiaomi Corporation",
        "vivo商业账户系统",
        "OPPO官方",
        "dsmmssqladmin",  # 邓白氏
        "huanjuly",  # 邓白氏
        "百度移动开放平台",
        "【华为开发者联盟】",
        "腾讯开放平台",
        "360box",
        "App Store Connect",
        "Xiaomi Corporation",
        "PP助手安卓开放平台",
        "vivo开放平台",
        "华为开发者联盟",
        "阿里应用分发开放平台",
        "vivo商业账户系统",
        "dev",  # 魅族
    ]
    attention_inform = [
        "成功上线！",  # 360
        "对接百度渠道成功",  # 百度
        'is now "Ready for Sale".',  # ios
        "已上架",  # xiaomi huawei
        "已成功上架华为应用市场",
        "应用审核通过",  # PP
        "审核不通过",  # pp,huawei
        "【华为开发者联盟】",
        "【百度移动开放平台】",
        "New Message from App Store Review Regarding",  # ios
        "在vivo平台已发布",
        "审核结果【未通过】",  # vivo
        "审核未通过",  # xiaomi
        "【更新邀请】您的应用在架版本较低",
        "在vivo平台被覆盖",
        "更新到最新版本",  # vivo更新
        "开发者资格审核不通过",
        "百度开发者帐号校验未通过！",
        "您的开发者身份审核通过",
        "开发者资格审核通过",
        "百度开发者帐号校验成功！",
        "您的帐号已审核通过。"
        "OPPO开放平台开发者账号审核结果通知",
        "【魅族应用商店】应用审核不通过通知",
        "【魅族应用商店】应用审核通过通知"
    ]


class _Mail:
    def __init__(self, account, password, Webhook, aite):
        """
            account: mail account
            password： mail password
            attention_mail: attention mail
        """
        self.account = account
        self.password = password
        self.Webhook = Webhook
        self.aite = aite
        _Params()
        chrome_options = webdriver.ChromeOptions()
        # chrome_options.add_argument('--headless')  # 无界面
        # self.br = webdriver.Chrome(options=chrome_options)
        self.br = webdriver.Firefox()
        self.LOGIN_STATE = False  # 登录状态

    def _login(self):
        """login reuse"""
        self.br.find_element_by_id("u").send_keys(self.account)
        self.br.find_element_by_id("p").send_keys(self.password)
        self.br.find_element_by_id("login_button").click()

    def Login(self):
        """login method"""
        try:
            self.br.switch_to.frame("login_frame")
            self._login()
            self.br.switch_to.default_content()  # 切回默认框架
            self.br.find_element_by_id("folder_1")  # 验证是否登录成功
            logging.log(logging.INFO, "登录成功")
            self.LOGIN_STATE = True
        except ElementNotInteractableException:
            self.br.find_element_by_id("switcher_plogin").click()
            self._login()
        except NoSuchElementException:
            logging.log(logging.INFO, "登录失败，请检查账号密码")
            self.LOGIN_STATE = False
            return
        finally:
            self.br.switch_to.default_content()

    def Polling_Mail(self):
        """polling mail"""
        if not self.LOGIN_STATE:  # 判断当前是否登录
            self.br.get(url=_Params.url)
            self.br.implicitly_wait(20)
            self.Login()  # 登录并返回登录状态
        try:
            self.br.find_element_by_id("folder_1").click()  # 是否需要重新登录
        except NoSuchElementException:  # 是否需要重新登录
            self.LOGIN_STATE = False
            logging.log(logging.INFO, "重新打开浏览器登陆")
            return
        else:
            try:
                time.sleep(3)
                self.br.switch_to.frame("mainFrame")
                self.br.find_element_by_id("_ua").find_element_by_link_text("未读邮件").click()  # 如果找到说明有未读文件
                logging.log(logging.INFO, "存在未读文件，开始查看是否存在关注邮件")
            except NoSuchElementException:
                logging.log(logging.INFO, "没有新邮件")
                return
            except StaleElementReferenceException:  # 元素未附加到文档
                self.Polling_Mail()
            else:
                for i in range(5):  # 加载所有邮件
                    self.br.execute_script("var q=document.documentElement.scrollTop=100000")
                    time.sleep(1)
                unread_mails = self.br.find_elements_by_xpath('//div[@class="ur_l_item_inner"]')
                attention_unread_mails = []  # 存储关注且未读邮件列表
                for i in unread_mails:
                    if i.find_element_by_css_selector(
                            '.ur_l_sender.grn.b_size.bold').text.strip() in _Params.attention_mail:  # 是否为关注邮件
                        attention_unread_mails.append(i)
                if len(attention_unread_mails) > 0:  # 重要邮件大于0
                    logging.log(logging.INFO, "存在重要邮件")
                    # 获取消息内容，是否包含指定消息
                    for i in attention_unread_mails:
                        text = i.find_element_by_css_selector(".ur_l_subject.b_size.bold").text.strip()

                        for j in _Params.attention_inform:  # 判断content 中是否包含指定文本
                            if text.find(j) != -1:
                                logging.log(logging.INFO, "\n" + "*" * 100 + "\n" + text)
                                text = "阿里分发：" + text if j == "应用审核通过" else text
                                text = "IOS：审核通过" if j == 'is now "Ready for Sale".' else text
                                text = "IOS：审核不通过" if j == "New Message from App Store Review Regarding" else text
                                self.Send_DingDing(info=text)  # 发送钉钉消息
                                break
                        else:
                            # self.Send_DingDing(info="收到一封未收录的重要邮件")  # 发送钉钉消息
                            logging.info(text)
                    self.br.find_element_by_css_selector(".btn_gray.right").send_keys(Keys.ENTER)
                else:
                    logging.log(logging.INFO, "不是重要邮件")
                    self.br.find_element_by_css_selector(".btn_gray.right").send_keys(Keys.ENTER)
                    return
        finally:
            self.br.switch_to.default_content()

    def Send_DingDing(self, info):
        """Send infomation"""
        headers = {
            'Content-Type': 'application/json'
        }
        data = {
            "msgtype": "text",
            "at": {
                "atMobiles": self.aite
            },
            "text": {
                "content": info
            }
        }
        response = requests.post(url=self.Webhook, data=json.dumps(data), headers=headers)  # send info
        if response.json()["errmsg"] == "ok":
            logging.log(logging.INFO, "消息发送成功")
        return


class Start:
    """启动类"""
    def __init__(self, account, password, polling_time):
        """
        :param account: the mail account
        :param password: the mail password
        :param polling_time: the code polling time
        """
        mail = _Mail(account,password)
        while True:
            logging.log(logging.INFO, "开始巡检")
            mail.Polling_Mail()
            logging.log(logging.INFO, "巡检结束，休眠")
            time.sleep(polling_time)
            logging.log(logging.INFO, "休眠结束")
        mail.br.quit()



if __name__ == "__main__":

    Start("754607626","czl58211438",1800)