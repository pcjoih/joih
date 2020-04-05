import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
import time
import sqlite3
import pandas as pd

TR_REQ_TIME_INTERVAL = 0.5

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self._create_kiwoom_instance()
        self._set_signal_slots()

    def _create_kiwoom_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def _set_signal_slots(self):
        self.OnEventConnect.connect(self._event_connect)
        self.OnReceiveTrData.connect(self._receive_tr_data)

    def comm_connect(self):
        self.dynamicCall("CommConnect()")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def _event_connect(self, err_code):
        if err_code == 0:
            print("connected")
        else:
            print("disconnected")

        self.login_event_loop.exit()

    def set_input_value(self, id, value):
        self.dynamicCall("SetInputValue(QString, QString)", id, value)

    def comm_rq_data(self, rqname, trcode, next, screen_no):
        self.dynamicCall("CommRqData(QString, QString, int, QString", rqname, trcode, next, screen_no)
        self.tr_event_loop = QEventLoop()
        self.tr_event_loop.exec_()

    def _comm_get_data(self, code, real_type, field_name, index, item_name):
        ret = self.dynamicCall("CommGetData(QString, QString, QString, int, QString", code,
                               real_type, field_name, index, item_name)
        return ret.strip()

    def _get_repeat_cnt(self, trcode, rqname):
        ret = self.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        return ret

    def _receive_tr_data(self, screen_no, rqname, trcode, record_name, next, unused1, unused2, unused3, unused4):
        if next == '2':
            self.remained_data = True
        else:
            self.remained_data = False

        if rqname == "req_1":
            self._opt50067(rqname, trcode)

        try:
            self.tr_event_loop.exit()
        except AttributeError:
            pass

    def _opt50067(self, rqname, trcode):
        data_cnt = self._get_repeat_cnt(trcode, rqname)
        for i in range(data_cnt):
            price = self._comm_get_data(trcode, "", rqname, i, "현재가")
            date = self._comm_get_data(trcode, "", rqname, i, "체결시간")
            volume = self._comm_get_data(trcode, "", rqname, i, "거래량")
            open = self._comm_get_data(trcode, "", rqname, i, "시가")
            high = self._comm_get_data(trcode, "", rqname, i, "고가")
            low = self._comm_get_data(trcode, "", rqname, i, "저가")
            self.data['date'].append(date)
            self.data['price'].append(price)
            self.data['volume'].append(volume)
            self.data['open'].append(open)
            self.data['high'].append(high)
            self.data['low'].append(low)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    kiwoom.comm_connect()

    kiwoom.data = {'date': [], 'price': [], 'volume': [], 'open': [], 'high': [], 'low': []}
    kiwoom.callcodelist = {
        'code': ['201Q4235', '201Q4237', '201Q4240', '201Q4242', '201Q4245', '201Q4247', '201Q4250', '201Q4252',
                 '201Q4255', '201Q4257', '201Q4260', '201Q4262', '201Q4265', '201Q4267', '201Q4270'], 'volume': []}
    kiwoom.putcodelist = {
        'code': ['301Q4235', '301Q4232', '3201Q4230', '301Q4227', '301Q4225', '301Q4222', '301Q4220', '301Q4217',
                 '301Q4215', '301Q4212', '301Q4210', '301Q4207', '301Q4205', '301Q4202', '301Q4200'], 'volume': []}
    req_count = 0
    con = sqlite3.connect("option 4월물.db")
    cur = con.cursor()

    for idx in kiwoom.callcodelist["code"]:
        cur.execute('SELECT "date" FROM "' + idx + '"')
        dbline = cur.fetchone()
        print("now processing call code : ", idx)
        kiwoom.set_input_value("종목코드", idx)
        kiwoom.set_input_value("시간단위", "1")
        kiwoom.comm_rq_data("req_1", "opt50067", 0, "0101")
        req_count += 1
        print("Request count : ", req_count)
        if req_count == 99:
            time.sleep(30)
        while kiwoom.remained_data == True:
            time.sleep(TR_REQ_TIME_INTERVAL)
            kiwoom.set_input_value("종목코드", idx)
            kiwoom.set_input_value("시간단위", "1")
            kiwoom.comm_rq_data("req_1", "opt50067", 2, "0101")
            req_count += 1
            print("Request count : ", req_count)
            if req_count == 99:
                time.sleep(30)
            if int(kiwoom.data['date'][-1]) <= int(dbline[0]):
                line_num = kiwoom.data['date'].index(dbline[0])
                del kiwoom.data['date'][line_num:]
                del kiwoom.data['price'][line_num:]
                del kiwoom.data['volume'][line_num:]
                del kiwoom.data['open'][line_num:]
                del kiwoom.data['high'][line_num:]
                del kiwoom.data['low'][line_num:]
                break
        df = pd.DataFrame(kiwoom.data, columns=['date', 'open', 'high', 'low', 'price', 'volume'])
        df.to_sql(idx, con, if_exists='append', index=False)
        con.commit()
        con.close()
        conn = sqlite3.connect("option 4월물.db")
        s = 'select * from "' + idx + '" ORDER BY "date" desc'
        df2 = pd.read_sql(s, con=conn)
        df2.to_sql(idx, conn, if_exists='replace', index=False)
        conn.commit()
        conn.close()

    for idx in kiwoom.putcodelist["code"]:
        cur.execute('SELECT "date" FROM "' + idx + '"')
        dbline = cur.fetchone()
        print("now processing put code : ", idx)
        kiwoom.set_input_value("종목코드", idx)
        kiwoom.set_input_value("시간단위", "1")
        kiwoom.comm_rq_data("req_1", "opt50067", 0, "0101")
        req_count += 1
        print("Request count : ", req_count)
        if req_count == 99:
            time.sleep(30)
        while kiwoom.remained_data == True:
            time.sleep(TR_REQ_TIME_INTERVAL)
            kiwoom.set_input_value("종목코드", idx)
            kiwoom.set_input_value("시간단위", "1")
            kiwoom.comm_rq_data("req_1", "opt50067", 2, "0101")
            req_count += 1
            print("Request count : ", req_count)
            if req_count == 99:
                time.sleep(30)
            if int(kiwoom.data['date'][-1]) <= int(dbline[0]):
                line_num = kiwoom.data['date'].index(dbline[0])
                del kiwoom.data['date'][line_num:]
                del kiwoom.data['price'][line_num:]
                del kiwoom.data['volume'][line_num:]
                del kiwoom.data['open'][line_num:]
                del kiwoom.data['high'][line_num:]
                del kiwoom.data['low'][line_num:]
                break
        df = pd.DataFrame(kiwoom.data, columns=['date', 'open', 'high', 'low', 'price', 'volume'])
        df.to_sql(idx, con, if_exists='append', index=False)
        con.commit()
        con.close()
        conn = sqlite3.connect("option 4월물.db")
        s = 'select * from "' + idx + '" ORDER BY "date" desc'
        df2 = pd.read_sql(s, con=conn)
        df2.to_sql(idx, conn, if_exists='replace', index=False)
        conn.commit()
        conn.close()
