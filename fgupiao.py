# _*_ coding:utf-8 _*_
"""
股票详情 2018.12.22
爬虫股市数据并发送到邮箱中
"""

import json
import sys
import threading
import time

import numpy as np
import pandas as pd
import requests
from bs4 import BeautifulSoup

import smtplib
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
import os

__author__ = "open_china"


class EmailSender(object):
    def __init__(self, sender, passWord, receivers):
        self.sender = sender
        self.passWord = passWord
        self.receivers = receivers

    def send(self, title, context, path):  # filepath文件夹下的文件作为附件
        msg = MIMEMultipart()
        msg['Subject'] = title
        msg['From'] = self.sender
        msg_content = context
        msg.attach(MIMEText(msg_content, 'plain', 'utf-8'))
        lis = os.listdir(path)
        print(lis)
        for fi in lis:  # 添加附件
            print("文件:" + fi)
            with open(path + "\\" + fi, 'rb') as f:
                print(fi)
                mime = MIMEBase('file', fi.split(".")[1], filename=fi)
                mime.add_header('Content-Disposition', 'attachment', filename=fi)
                mime.add_header('Content-ID', '<0>')
                mime.add_header('X-Attachment-Id', '0')
                mime.set_payload(f.read(), "utf-8")
                encoders.encode_base64(mime)
                msg.attach(mime)
        try:
            s = smtplib.SMTP_SSL("smtp.qq.com", 465)  # QQsmtp服务器的端口号为465或587
            s.set_debuglevel(1)
            s.login(self.sender, self.passWord)
            s.sendmail(self.sender, self.receivers, msg.as_string())
            s.quit()
            print("All emails have been sent over!")
            return True
        except smtplib.SMTPException as e:
            print("error----:", e)
            return False


# 多线程下载获取返回值
class DownloadThread(threading.Thread):
    def __init__(self, func, args=()):
        super(DownloadThread, self).__init__()
        self.func = func
        self.args = args

    def run(self):
        self.result = self.func(*self.args)

    def get_result(self):
        threading.Thread.join(self)
        try:
            return self.result
        except Exception as e:
            print("error:", sys._getframe().f_lineno, e)
            pass


class GupiaoSpider(object):
    def __init__(self):
        self.baseurl = "http://quote.eastmoney.com/stocklist.html"
        self.Data = []
        self.Date = time.strftime('%Y%m%d')
        self.Recordpath = 'E:\\pythonData\\股票数据\\'
        if os.path.exists(self.Recordpath) is False:
            os.makedirs(self.Recordpath)
        self.filename = 'Data' + self.Date
        self.limit = 800  # 设置开启N个线程
        self.session = requests.Session()
        self.timeout = 100

    def getTotalUrl(self):
        try:
            req = self.session.get(self.baseurl, timeout=self.timeout)
            if int(req.status_code) != 200:
                return None
            req.encoding = "gbk"
            lis = BeautifulSoup(req.text, 'lxml').select("div.quotebody li")
            data_lis = []
            for msg in lis:
                cuturl = msg.a["href"].split("/")[-1].replace(".html", "")
                names = msg.text.split("(")
                name = names[0]
                code = names[1].replace(")", "")
                if not (cuturl.startswith("sz300") or cuturl.startswith("sh002")):
                    continue
                add = {"url": cuturl, "name": name, "code": code}
                data_lis.append(add)
            return data_lis
        except Exception as e:
            print(sys._getframe().f_lineno, e)
            return None

    def down(self, url, name, code):
        record_d = {}
        record_d["名称"] = name
        record_d["代码"] = code
        linkurl = "https://gupiao.baidu.com/api/stocks/stockdaybar?from=pc&os_ver=1&cuid=xxx&vv=100&format=json&stock_code=" + \
                  url + "&step=3&start=&count=160&fq_type=no&timestamp=" + str(int(time.time()))
        try:
            resp = self.session.get(linkurl, timeout=self.timeout).content
            js = json.loads(resp)
            lis = js.get("mashData", "-")
            msg = lis[0].get("kline")
            record_d["涨幅"] = str(format(float(msg.get("netChangeRatio", "-")), ".2f")) + "%"
            record_d["开盘"] = msg.get("open", "-")
            record_d["最高"] = msg.get("high", "-")
            record_d["最低"] = msg.get("low", "-")
            record_d["收盘"] = msg.get("close", "-")
            record_d["成交量"] = msg.get("volume", "-")
            record_d["昨收"] = msg.get("preClose", "-")
            record_d["收盘"] = msg.get("close", "-")
            print("完成数据:  " + name, code)
            return record_d
        except Exception as e:
            print(sys._getframe().f_lineno, e, name, code)
            return None

    def download(self, tups):
        lis = list(tups)
        col = int(np.floor(len(lis) / self.limit))
        downlis = np.array(lis[0:col * self.limit]).reshape(col, self.limit).tolist()
        if col * self.limit < len(lis):
            downlis.append(lis[col * self.limit:])
        for urls in downlis:
            threads = []
            for parms in urls:
                task = DownloadThread(self.down, (parms["url"], parms["name"], parms["code"]))
                task.start()
                threads.append(task)
            for t in threads:  # 守护线程
                result = t.get_result()
                if result is not None:
                    self.Data.append(result)
        print("批量下载结束...")
        self.save()

    def save(self):
        df = pd.DataFrame(self.Data)
        df.to_excel(self.Recordpath + self.filename + '.xls', index=False)  # 未排名
        df["涨幅"] = df["涨幅"].apply(lambda x: float(str(x).replace("%", "")))
        df = df.sort_values(by=["涨幅"], ascending=[False])
        df["涨幅"] = df["涨幅"].apply(lambda x: str(x) + "%")
        df.to_excel(self.Recordpath + self.filename + '排名.xls', index=False)
        print("保存文件成功：", self.Recordpath)


if __name__ == "__main__":
    context = "股市分析数据如下\n清悉知!!!\n 来自python客户端"
    sender = 'xxxxx@qq.com'
    passWord = 'xxxxxxxxxxxxxxxxxx'
    receivers = ['xxxxxx@qq.com', "xxxxxx@qq.com"]
    stime = time.time()
    spider = GupiaoSpider()
    urllis = spider.getTotalUrl()
    if urllis is not None:
        spider.download(urllis)
    path = spider.Recordpath
    se = EmailSender(sender, passWord, receivers)
    title = spider.Date + "股市分析数据"
    se.send(title, context, path)
    etime = time.time()
    print("used: ", str(etime - stime))
