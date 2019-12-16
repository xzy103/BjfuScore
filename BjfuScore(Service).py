# -*- coding:UTF-8 -*-
# Author: xzy103
# Ver: Service(191216)

import os
import sys
import requests
import win32api
import win32con
from bs4 import BeautifulSoup as bs
from prettytable import PrettyTable
import time

url_start = 'http://newjwxt.bjfu.edu.cn/jsxsd/xsxk/xklc_list?Ves632DSdyV=NEW_XSD_PYGL'  # GET
url_main = 'http://newjwxt.bjfu.edu.cn/jsxsd/framework/xsMain.jsp'  # GET
url_login = 'http://newjwxt.bjfu.edu.cn/jsxsd/xk/LoginToXk'  # POST
url_score = 'http://newjwxt.bjfu.edu.cn/jsxsd/kscj/cjcx_list'


# 发送消息到微信
def sendmail():
    url = 'https://sc.ftqq.com/{}.send'.format(key)
    payload = {'text': "成绩有更新！", 'desp': "RT"}
    requests.post(url, params=payload, timeout=20)
    print('已发送 ')


# 登录教务系统
class Newjwxt(object):
    def __init__(self, username, password):
        self.headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'keep-alive',
            'Content-Type': 'application/x-www-form-urlencoded',
            'Host': 'newjwxt.bjfu.edu.cn',
            'Referer': 'http://newjwxt.bjfu.edu.cn/Logon.do?method=logon',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.186 Safari/537.36'
        }
        self.session = requests.Session()
        self.session.headers.update(self.headers)
        self.session.get(url_start)
        self.username = username
        self.password = password

    def login(self):
        postdata = {
            'USERNAME': self.username,
            'PASSWORD': self.password
        }
        self.session.post(url_login, data=postdata)
        if self.session.get(url_main).status_code == 200:
            print(">>>>>登录成功!")
            return self.session
        else:
            print(">>>>>登录失败!")
            time.sleep(3)
            sys.exit()

    def info(self):
        r = self.session.get(url_main)
        soup = bs(r.content, "html.parser")
        basic_info = soup.find_all("div", {"class": "Nsb_top_menu_nc"})
        info = basic_info[0].text.strip()
        print('账号信息：'+info)

    def write(self):
        with open("score.txt", "w") as fw:
            fw.write(str(self.score))

    def saveclasses(self):
        r = self.session.get(url_score)
        soup = bs(r.content, "html.parser")

        score, title = [], []
        self.score = score
        trs = soup.find('table', {'id': 'dataList'}).find_all('tr')
        for th in trs[0].find_all('th'):
            title.append(th.string.strip())
        score.append(title)
        for i in range(1, len(trs)):
            ls = []
            for td in trs[i].find_all('td'):
                txt = td.text.strip().replace('Ⅰ', '1').replace('Ⅱ', '2')
                ls.append(txt)
            score.append(ls)
        table = PrettyTable(score[0])
        for i in range(1, len(score)):
            table.add_row(score[i])

        grade = soup.find_all('div', {'class': 'Nsb_pw'})
        grade = grade[2].text
        grade = grade.split('在本查询时间段，')[-1].split('所得学分详情')[0]
        grade = grade.replace("\t", '').replace('\r', '').replace('、\n', '').strip()
        
        t1, t2, t3 = [], [], []
        trs = soup.find_all('table', {'id': 'dataList'})[1].find_all('tr')
        for th in trs[0].find_all('th'):
            t1.append(th.string.strip())
        for th in trs[1].find_all('th'):
            t2.append(th.string.strip())
        for td in trs[-1].find_all('td'):
            t3.append(td.string.strip())
        t2[0], t2[3] = '专选总计', '公选总计'
        table2 = PrettyTable(t1[:3]+t2)
        table2.add_row(t3)

        rank, title2 = [], []
        trs = soup.find_all('table', {'id': 'dataList'})[2].find_all('tr')
        for th in trs[0].find_all('th'):
            title2.append(th.string.strip())
        rank.append(title2)
        for i in range(1, len(trs)):
            ls = []
            for td in trs[i].find_all('td'):
                txt = td.text.strip()
                ls.append(txt)
            rank.append(ls)
        table3 = PrettyTable(rank[0])
        for i in range(1, len(rank)):
            table3.add_row(rank[i])

        if os.path.isfile('score.txt') is False:
            self.write()
            print('>>>>>成绩信息已保存')
        with open('score.txt', "r") as fr:
            content = eval(fr.read())
            win32api.SetFileAttributes('score.txt', win32con.FILE_ATTRIBUTE_NORMAL)
        if content == score:
            print('进行了一次查询未更新...')
        else:
            print('成绩有更新！')
            sendmail()
        self.write()
        win32api.SetFileAttributes('score.txt', win32con.FILE_ATTRIBUTE_HIDDEN)


# 获取账户信息
def get_user_info():
    if os.path.isfile('jwxtinfo_service') is True:
        with open('jwxtinfo_service', 'r') as f:
            Usn = f.readline()[:-1]
            Psw = f.readline()[:-1]
            key = f.readline()[:-1]
    else:
        with open('jwxtinfo_service', 'w') as f:
            print(">>>>>初次使用需初始化账号和密码")
            Usn = input('>>>>>请输入教务系统账号：')
            Psw = input('>>>>>请输入教务系统密码：')
            key = input('>>>>>请输入方糖密钥：')
            f.write(Usn + '\n' + Psw + '\n' + key + '\n')
            win32api.SetFileAttributes('jwxtinfo_service', win32con.FILE_ATTRIBUTE_HIDDEN)
    return Usn, Psw, key


if __name__ == '__main__':
    print("北林全家桶 | 教务系统成绩查询 | Service")
    print(">>>>>开始运行...")
    # UserName, PassWord = "160405211", "123456"
    UserName, PassWord, key = get_user_info()
    new = Newjwxt(UserName, PassWord)
    new.login()
    new.info()
    while True:
        new.saveclasses()
        time.sleep(1200)  # 每十二分钟更新一次
