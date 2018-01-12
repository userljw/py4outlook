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
    mail.Recipients.Add('1@.com.cn')
    mail.CC='1@.com.cn;2.com.cn'
    mail.Subject = unicode(' ' % today)
    body = unicode("！")
    mail.Body = body
    attachment = unicode(r"")
    mail.Attachments.Add(attachment)
    mail.Send()
    print "send ok"



if __name__ == "__main__":
    win32api.ShellExecute(0, 'open', r'C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE', '', '', 1)
    outlook()
