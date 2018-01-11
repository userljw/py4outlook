# coding:utf-8
import sys
reload(sys)
sys.setdefaultencoding('utf8')

import win32com.client as win32
import win32api
import datetime


def outlook():
    today=datetime.date.today()
    print today
    app = 'Outlook'
    olook = win32.gencache.EnsureDispatch("%s.Application" % app)
    mail = olook.CreateItem(win32.constants.olMailItem)
    mail.Recipients.Add('1@spdb.com.cn')
    mail.CC='1@spdb.com.cn;2@spdb.com.cn'
    mail.Subject = unicode('新员工日小结-陆嘉炜%s ' % today)
    body = unicode("程老师您好!""\r\n 附件是我今天的日小结，烦请查收，谢谢！")
    mail.Body = body
    attachment = unicode(r"D:\资料\新员工入职\每日小结 -陆嘉炜.xlsx")
    mail.Attachments.Add(attachment)
    mail.Send()
    print "send ok"



if __name__ == "__main__":
    win32api.ShellExecute(0, 'open', r'C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE', '', '', 1)
    outlook()
