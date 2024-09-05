import requests
import pandas as pd
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import time
import execjs
session=requests.session()
def get_list_1():
    headers = {
        "Accept": "*/*",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Cache-Control": "no-cache",
        "Connection": "keep-alive",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Origin": "http://www.cfcpn.com",
        "Pragma": "no-cache",
        "Referer": "http://www.cfcpn.com/jcw/sys/index/goUrl?url=modules/sys/login/list&column=cggg",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        "X-Requested-With": "XMLHttpRequest"
    }
    cookies = {
        "pageNo": "0",
        "wolfking.jeeplus.session.id": "6f309402-5894-4ada-baeb-e6b880c24aa9",
        "pageSize": "10"
    }
    url = "http://www.cfcpn.com/jcw/noticeinfo/noticeInfo/dataNoticeList"
    list_info = []
    for i in range(1,2):
        data = {
            "noticeType": "1",
            "pageSize": "10",
            "pageNo": str(i),
            "noticeState": "1",
            "isValid": "1",
            "orderBy": "publish_time desc"
        }
        response = requests.post(url, headers=headers, cookies=cookies, data=data, verify=False)
        data=response.json()
        rows=data.get('rows',[])

        for row in rows:
            dict_row = {}.fromkeys(['id', 'noticeTitle', 'area'],None)
            dict_row["id"]=row.get("id","")
            dict_row["noticeTitle"] = row.get("noticeTitle", "")
            dict_row["area"] = row.get("area", "")
            if row.get("id",""):
                get_detail(row.get("id",""))
            list_info.append(dict_row)
        df=pd.DataFrame(list_info)
        df.to_excel('金融投标.xlsx')

def get_list():
    headers = {
        "Accept": "*/*",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Cache-Control": "no-cache",
        "Connection": "keep-alive",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Origin": "http://www.cfcpn.com",
        "Pragma": "no-cache",
        "Referer": "http://www.cfcpn.com/jcw/sys/index/goUrl?url=modules/sys/login/list&column=cggg",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        "X-Requested-With": "XMLHttpRequest"
    }

    url = "http://www.cfcpn.com/jcw/noticeinfo/noticeInfo/dataNoticeList"

    for i in range(4, 5):
        list_info = []
        data = {
            "^noticeType": "1",
            "pageSize": "10",
            "pageNo": str(i),
            "noticeState": "1",
            "isValid": "1",
            "orderBy": "publish_time desc",
            "noticeContent": "",
            "briefContent": "",
            "noticeTitle": "",
            "purchaseName": "",
            "purchaseId": "",
            "categoryLabName": "",
            "beginPublishTime": "",
            "endPublishTime": "",
            "areaProvince": "",
            "labelAllId": "263dbb969f9a480f898a84f101914901"
        }
        cookies = {
            "wolfking.jeeplus.session.id": "6f309402-5894-4ada-baeb-e6b880c24aa9",
            "pageSize": "10",
            "pageNo": str(i-1)
        }
        response = session.post(url, headers=headers, cookies=cookies, data=data, verify=False)
        data = response.json()
        rows = data.get('rows', [])
        try:
            for row in rows:
                dict_row = {}.fromkeys(['id', '标题','发布时间' '地区'], None)
                dict_row["id"] = row.get("id", "")
                dict_row["标题"] = row.get("noticeTitle", "")
                dict_row["发布时间"] = row.get("publishTime", "")
                dict_row["area"] = row.get("area", "")
                if row.get("id", ""):
                    time.sleep(10)
                    # get_detail(row.get("id", ""),str(i-1))

                list_info.append(dict_row)
        except Exception as ex:
            print(row.get("noticeTitle", ""))
        df = pd.DataFrame(list_info)
        df.to_excel('金融投标'+str(i)+'.xlsx')
if __name__ =="__main__":
    get_list()