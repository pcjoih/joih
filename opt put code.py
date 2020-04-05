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
    kiwoom.putcodelist = {
        'code': ['301Q4235', '301Q4232', '3201Q4230', '301Q4227', '301Q4225', '301Q4222', '301Q4220', '301Q4217',
                 '301Q4215', '301Q4212', '301Q4210', '301Q4207', '301Q4205', '301Q4202', '301Q4200'], 'volume': []}
    req_count = 0
    con = sqlite3.connect("option 4월물.db")

    for idx in kiwoom.putcodelist["code"]:
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
            if int(kiwoom.data['date'][-1]) < 20200313090000:
                    break
        if int(kiwoom.data['date'][-1]) < 20200313090000:
            for id1, val in enumerate(kiwoom.data['date']):
                if int(val) < 20200313090000:
                    last_idx = id1
                    break
            del kiwoom.data['date'][last_idx:]
            del kiwoom.data['price'][last_idx:]
            del kiwoom.data['volume'][last_idx:]
            del kiwoom.data['open'][last_idx:]
            del kiwoom.data['high'][last_idx:]
            del kiwoom.data['low'][last_idx:]
        df = pd.DataFrame(kiwoom.data, columns=['date', 'open', 'high', 'low', 'price', 'volume'])
        df.to_sql(idx, con, if_exists='replace', index=False)
