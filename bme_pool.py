#!/usr/local/bin/python
# -*- coding: utf-8 -*-
import sys,os
from os.path import dirname, join, abspath
sys.path.insert(0, abspath(join(dirname(__file__), '..')))
import requests
import time
import csv
import xlsxwriter 
import xlrd 
import xlwt
import re
from datetime import datetime
from lxml import html,etree
from xlutils.copy import copy
from config import *

def doLogin():
    try:
        payload = {
                'user': USERNAME,
                'pass': PASSWORD,
                'Submit': "Submit"
            }
        with requests.Session() as sess:
            result = sess.post(LOGIN_URL,data=payload,headers=dict(referer=LOGIN_URL))
            if result.status_code == 200:
                return 200,sess
            else:
                return result.status_code,"Can not login"
    except Exception as ex:
        print('[LOGIN]- {}'.format(ex))
        return None

def doPoolProcess(sess,data,nextDate):
    try:
        doGetFormPool = sess.get(url=URL_POOL,headers=dict(referer=URL_POOL_REFERER))
        if doGetFormPool.status_code == 200:
            doPost = sess.post(url=URL_FIND,headers=dict(referer=URL_POOL),data=data)

            if doPost.status_code == 200:
                tree = html.fromstring(doPost.content)

                reg = re.compile('((?!\r\n)\s+)')
                ls = str(reg.sub(' ',str(doPost.content))).split(" ")
                myText = "".join(ls)

                ss = re.findall(r'<td>&nbsp;(\d[0-9]+)</td>',str(myText))

                itemno_hd_1 = re.findall(r'itemno_hd=(\d[0-9]+)&',str(myText))

                reserve_no = re.findall(r'reserve_no;(\d[0-9]+)"',str(myText))

                url = "http://nsmart.nhealth-asia.com/MTDPDB01/reserve/reserve_back1.php?itemno_hd={}&ccsForm=tbl_reserve1".format(itemno_hd_1[0])
                referer = "http://nsmart.nhealth-asia.com/MTDPDB01/reserve/reserve_back1.php?itemno_hd={}".format(itemno_hd_1)

                packet = {
                    "FormState": "1;0;reserve_no;{}".format(reserve_no[0]),
                    "itemno_hd_1": itemno_hd_1[0],
                    "hn_1": 1,
                    "reserve_dept_1": DEPARTMENT_ID,
                    "reserve_branchid_1":BRANCH_ID,
                    "inform_date_1": data['requestdate'],
                    "inform_date1_1": nextDate,
                    "Button_Submit": "บันทึก"
                }

                result = sess.post(url=url,headers=dict(referer=referer),data=packet)

                if result.status_code == 200:
                    print("[{}] - Create Pool Successfully.".format(data['sap_code']))
                    return True
    
    except Exception as ex:
        print('[POOL]- {}'.format(ex))
        return None

def doChangeDepartment(dept):
    return DEPT[str(dept)]

def doMainProcess(sess,loc):
    try:
        i = 1
        ck = True
        while ck != False:
            rb = xlrd.open_workbook(loc) 
            sheet = rb.sheet_by_index(0) 
  
            sheet.cell_value(0, 0)
            date_tuple = xlrd.xldate_as_tuple(sheet.row_values(i)[2], rb.datemode)
            _date = datetime(*date_tuple).strftime('%d/%m/%Y').split('/')
            newDate = _date[0]+'/'+_date[1]+'/'+str(int(_date[2]))

            ndate_tuple = xlrd.xldate_as_tuple(sheet.row_values(i)[3], rb.datemode)
            _ndate = datetime(*ndate_tuple).strftime('%d/%m/%Y').split('/')
            nextDate = _ndate[0]+'/'+_ndate[1]+'/'+str(int(_ndate[2]))

            now = datetime.now()
            s2 = now.strftime("%d/%m/%Y")

            data = {
                "sap_code": sheet.row_values(i)[1],
                "requestdate": newDate,
                "equip_unit": 1,
                "old_reserve": 1,
                "groupid": "",
                "catagory": "",
                "reserve_branchid": BRANCH_ID,
                "reserve_dept": DEPARTMENT_ID,
                "reserve_sub_dept": doChangeDepartment(str(sheet.row_values(i)[4])),
                "borrow_name": str(sheet.row_values(i)[5]).encode('tis_620'),
                "borrow_tele": str(int(sheet.row_values(i)[6])),
                "enterby": "Engineer",
                "enterdate": str(s2),
                "Button_Insert.x": 18,
                "Button_Insert.y": 9
            }

            if sheet.row_values(i)[7] not in ["True",True]:
                result = doPoolProcess(sess,data,nextDate)
                if result == True:
                    wb = copy(rb)
                    w_sheet = wb.get_sheet(0)
                    w_sheet.write(i,7,'True')
                    wb.save(loc)
                else:
                    wb = copy(rb)
                    w_sheet = wb.get_sheet(0)
                    w_sheet.write(i,7,'False')
                    wb.save(loc)

            if int(sheet.row_values(i)[0]) is None:
                ck = False
                break
            else:
                i+=1
                time.sleep(1)
            
    except Exception as ex:
        print('[MAIN]- {}'.format(ex))

if __name__ == "__main__":
    try:
        status,sess = doLogin()
        _start = time.time()
        if status == 200:
            filename = input("Enter Excel file name : ")
            loc = '{}.xls'.format(filename)
            doMainProcess(sess,loc)
            _end = time.time()
            print('[END] - Total time %.2f min.'%((_end-_start)/60))
    except KeyboardInterrupt:
        print('[MAIN]- Pool Service stop!..')
        sys.exit()
    
