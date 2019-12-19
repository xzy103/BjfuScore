# -*- coding: utf-8 -*-
# author: xzy103
# ver: 2.1(191217)

import requests
import time
import sys
import os
import win32api
import win32con
import datetime
import PySimpleGUI as sg
from bs4 import BeautifulSoup as bs
from prettytable import PrettyTable


layout = [
    [sg.Output(size=(150, 50), font=['黑体', '8'], text_color='green')],
    [sg.Exit('关闭窗口', font=['黑体', '10'], bind_return_key=True, button_color=('white', 'green'))],
    [sg.Exit('后台监测', font=['黑体', '10'], bind_return_key=True, button_color=('white', 'green'))]]
window = sg.Window('BjfuJwxt', layout, background_color='white', no_titlebar=True, resizable=True)


# 发送消息到微信
def sendmail(Key, title, msg):
    url = 'https://sc.ftqq.com/{}.send'.format(Key)
    payload = {'text': title, 'desp': msg}
    requests.post(url, params=payload, timeout=20)


# 屏幕输出
def guiprint(text):
    window.Read(timeout=0)
    print(text)


url_start = 'http://newjwxt.bjfu.edu.cn/jsxsd/xsxk/xklc_list?Ves632DSdyV=NEW_XSD_PYGL'  # GET
url_main = 'http://newjwxt.bjfu.edu.cn/jsxsd/framework/xsMain.jsp'  # GET
url_login = 'http://newjwxt.bjfu.edu.cn/jsxsd/xk/LoginToXk'  # POST
url_score = 'http://newjwxt.bjfu.edu.cn/jsxsd/kscj/cjcx_list'


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
            guiprint(">>>>>登录成功!")
            return self.session
        else:
            guiprint(">>>>>登录失败!")
            time.sleep(3)
            sys.exit()

    def info(self):
        r = self.session.get(url_main)
        soup = bs(r.content, "html.parser")
        basic_info = soup.find_all("div", {"class": "Nsb_top_menu_nc"})
        info = basic_info[0].text.strip()
        guiprint('账号信息：' + info)

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
        guiprint('>>>>>成绩详情')
        guiprint(str(table))
        guiprint('\n')

        grade = soup.find_all('div', {'class': 'Nsb_pw'})
        grade = grade[2].text
        grade = grade.split('在本查询时间段，')[-1].split('所得学分详情')[0]
        grade = grade.replace("\t", '').replace('\r', '').replace('、\n', '').strip()
        guiprint(grade)
        guiprint('\n')

        t1, t2, t3 = [], [], []
        trs = soup.find_all('table', {'id': 'dataList'})[1].find_all('tr')
        for th in trs[0].find_all('th'):
            t1.append(th.string.strip())
        for th in trs[1].find_all('th'):
            t2.append(th.string.strip())
        for td in trs[-1].find_all('td'):
            t3.append(td.string.strip())
        t2[0], t2[3] = '专选总计', '公选总计'
        table2 = PrettyTable(t1[:3] + t2)
        table2.add_row(t3)
        guiprint('>>>>>学分详情')
        guiprint(str(table2))
        guiprint('\n')

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
        guiprint('>>>>>排名详情')
        guiprint(str(table3))
        guiprint('\n')

        if os.path.isfile('score.txt') is False:
            self.write()
            guiprint('>>>>>成绩信息已保存')
        with open('score.txt', "r") as fr:
            content = eval(fr.read())
            win32api.SetFileAttributes('score.txt', win32con.FILE_ATTRIBUTE_NORMAL)
        if content == score:
            guiprint('当前是最新成绩！')
        else:
            table_new = PrettyTable(score[0])
            for sc in score:
                if sc not in content:
                    table_new.add_row(sc)
            guiprint('>>>>>更新的成绩为：')
            guiprint(str(table_new))
            guiprint('>>>>>成绩信息已更新')
            sg.PopupOK('成绩有更新！', no_titlebar=True, text_color='green', font=['黑体', '12'], button_color=('white', 'green'))
        self.write()
        win32api.SetFileAttributes('score.txt', win32con.FILE_ATTRIBUTE_HIDDEN)

    def refeshclasses(self):
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
        table2 = PrettyTable(t1[:3] + t2)
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
        with open('score.txt', "r") as fr:
            content = eval(fr.read())
            fr.close()
            win32api.SetFileAttributes('score.txt', win32con.FILE_ATTRIBUTE_NORMAL)
        if content == score:
            pass
        else:
            table_new = PrettyTable(score[0])
            message = ''
            for sc in score:
                if sc not in content:
                    table_new.add_row(sc)
                    message += str(sc[3]) + '\n' +str(sc[4]) + '\n'
            # FormatTime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.time()))
            sendmail(Key, title="成绩有更新！", msg=message)
            sg.PopupOK('成绩有更新！\n'+str(table_new), auto_close=True, auto_close_duration=120,
                       no_titlebar=True, text_color='green', font=['黑体', '10'], button_color=('white', 'green'))
        self.write()
        win32api.SetFileAttributes('score.txt', win32con.FILE_ATTRIBUTE_HIDDEN)


# 获取账户信息
def get_user_web_info():
    if os.path.isfile('jwxtinfo') is True:
        with open('jwxtinfo', 'r') as f:
            Usn = f.readline()[:-1]
            Psw = f.readline()[:-1]
            Key = f.readline()[:-1]
            f.close()
    else:
        with open('jwxtinfo', 'w') as f:
            layout_input = [[sg.Text('请输入教务系统账号和密码！')],
                            [sg.Text('账号'), sg.Input(key='_UN_', focus=True)],
                            [sg.Text('密码'), sg.Input(key='_PW_', password_char='*')],
                            [sg.Text('密钥'), sg.Input(key='_KEY_')],
                            [sg.Button('验证密钥', bind_return_key=True)]]
            window_input = sg.Window('BjfuJwxt', layout_input, keep_on_top=True)
            event, result = window_input.Read()
            Usn, Psw, Key = result['_UN_'], result['_PW_'], result['_KEY_']

            if "" in [Usn, Psw, Key]:
                sg.PopupOK('信息不完整！', keep_on_top=True)
                f.close()
                os.remove('jwxtinfo')
                sys.exit()
            elif event == '验证密钥':
                sendmail(Key, title='验证消息', msg='如果您能收到此消息说明验证通过\n'+str(time.time()))
                sg.PopupOK('请留意微信是否收到消息', keep_on_top=True)
                f.write(Usn + '\n' + Psw + '\n' + Key + '\n')
                f.close()
            else:
                f.close()
                os.remove('jwxtinfo')
                sys.exit()
            window_input.Close()
            win32api.SetFileAttributes('jwxtinfo', win32con.FILE_ATTRIBUTE_HIDDEN)
    return Usn, Psw, Key


if __name__ == '__main__':
    # UserName, PassWord = "160405211", "123456"
    UserName, PassWord, Key = get_user_web_info()  # 获取账户信息
    new = Newjwxt(UserName, PassWord)  # 登录教务系统
    new.login()  # 登录
    new.info()  # 账户信息
    new.saveclasses()  # 成绩信息
    event, result = window.Read()  # 屏幕输出
    for i in range(20):
        window.SetAlpha((20-i)/20)
        time.sleep(0.02)
    window.Close()  # 关闭窗口

    if event == '后台监测':
        sg.PopupNoButtons('20min刷新一次\n持续监测中...\n程序将转入后台运行',
                          auto_close_duration=3, no_titlebar=True,
                          auto_close=True, text_color='green', font=['黑体', '10'])
        while True:
            try:
                new.login()
                new.info()
                new.refeshclasses()
                time.sleep(1200)
            except:
                FormatTime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.time()))
                sendmail(Key, title='程序运行错误', msg=FormatTime)
                sg.PopupOK('监测程序出现错误，可能原因是登录超时，请尝试重启程序。',
                           no_titlebar=True, text_color='green', font=['黑体', '10'], button_color=('white', 'green'))
                sys.exit()
    else:
        sys.exit()  # 退出程序
