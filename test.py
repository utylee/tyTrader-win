import win32com
from PyQt5.QtCore import QUrl
from win32com import client 
test =win32com.client.Dispatch("KHOPENAPI.KHOpenAPICtrl.1")
test.CommConnect()
print('테스트 입니당 데헷')
