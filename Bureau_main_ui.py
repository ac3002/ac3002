# -*- coding: utf-8 -*-
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6 import uic
import time
import afmod
import afmod1

class Window(QWidget):
    chang_cusorSignal = pyqtSignal(bool)
    chang_cusorSignal1 = pyqtSignal(bool)
    def __init__(self):
        super(Window, self).__init__()
#        self.ui = uic.loadUi(r"D:\python_work\QtUi\daan1120131.ui",self)
        self.ui = uic.loadUi(r".\daan1120131.ui", self)
        self.chang_cusorSignal.connect(self.hide_stat)
        self.chang_cusorSignal1.connect(self.hide1_stat)
        global user, ip, pwd, sid, ora_dir, offname, prtoutfile_path
        self.stat_in = self.findChild(QLineEdit, 'stat_day')
        self.end_in = self.findChild(QLineEdit, 'end_day')
        self.stat_btn = self.findChild(QPushButton, 'exe_btn')
        self.show_msg = self.findChild(QLabel, 'show_message')
        self.cle_btn = self.findChild(QLineEdit, 'closebtn')
        self.stat_btn.clicked.connect(self.exe_show)
        self.stat_check()
        self.stat_in.returnPressed.connect(self.on_return_pressed)
        self.end_in.returnPressed.connect(self.on_return_pressed)

    def on_return_pressed(self):
        if self.stat_in.hasFocus():
            self.end_in.setFocus()
        elif self.end_in.hasFocus():
            self.stat_in.setFocus()
            

    def hide_stat(self,flag):
#        self.setCursor(Qt.CursorShape.WaitCursor) 將游標變更為忙碌
        self.stat_btn.setEnabled(flag)
        QMessageBox.information(self,"產製中訊息","資料產製中請稍待----")
    def hide1_stat(self,flag):
#        self.setCursor(Qt.CursorShape.ArrowCursor) 將游標變更為正常忙碌
        self.stat_in.clear()
        self.end_in.clear()
        self.stat_btn.setEnabled(flag)

    def stat_check(self):
        import os
        if os.path.exists('Bureau_reportV3.ini') is False:  # 1110210 add
            self.show_msg.setText('Bureau_reportV3.ini檔案不存在,請檢查!!!!!')
            self.stat_in.setFocus()
            return False
        else:
            filename = afmod.read_init('Bureau_reportV3.ini')
            offname, user, pwd, ip, sid, ora_dir, prtoutfile_path, sample_xls = afmod.read_ini(filename)
            if os.path.exists(sample_xls) is False:
                self.show_msg.setText("xls範例檔案不存在,請檢查!!!!!")
                self.stat_in.setFocus()
                return False
            else:
                return True

    def exe_show(self):
        if len(self.stat_in.text()) != 7 :
            self.show_msg.setText('啟始日期不符格式!!!!')
            self.stat_in.setFocus()
        elif len(self.end_in.text()) != 7 :
            self.show_msg.setText('終止日期不符格式!!!!')
            self.end_in.setFocus()
        else:
            if self.stat_check() is True:
                self.show_msg.setText('報表產製中，請稍後!!!!')
                QApplication.processEvents()   # 強制更新畫面1120322
#                self.chang_cusorSignal.emit(False)
#                today = ''; tomorrow =''; tm_year=''; tm_mon=''; tm_mday=''; tm_hour=''; tm_min=''; tm_sec=''
                today, tomorrow, tm_year, tm_mon, tm_mday, tm_hour, tm_min, tm_sec = afmod.get_today()
                st_day = self.stat_in.text()
                end_day = self.end_in.text()
                out_xls = st_day + end_day + str(tm_hour) + str(tm_min) + str(tm_sec)
                try:
                    filename = afmod.read_init('Bureau_reportV3.ini')
                    offname, user, pwd, ip, sid, ora_dir, prtoutfile_path, sample_xls = afmod.read_ini(filename)
                    ap_str = str(ord('l') + ord('a'))  # 加解密用內含指定字元
                    user = afmod1.dectry(user, ap_str)  # oracle user
                    pwd = afmod1.dectry(pwd, ap_str)  # oracle user的pwd
                    ip = afmod1.dectry(ip, ap_str)  # oracle 的連結ip
                    sid = afmod1.dectry(sid, ap_str)  # oracle 的連結名稱
                    out_xls = prtoutfile_path + out_xls + offname + '.xlsx'
                    try:
                        afmod.con_ora_path(ora_dir, 'A')  # 讀取連結oracle lib的目錄
                    except:
                        self.show_msg.setText('讀取連結oracle lib的目錄錯誤!!!!')
                except:
                    self.show_msg.setText('解密錯誤!!!!')
                    exit()
#                print(st_day, end_day, out_xls, user, pwd, ip, sid, sample_xls) #----以上ok

                import Bureau_sheet3  # 他縣市承辦本所案件資料
                Bureau_sheet3.Bureau_s3(st_day, end_day, out_xls, user, pwd, ip, sid, sample_xls)
                del Bureau_sheet3
                import Bureau_sheet1
                Bureau_sheet1.Bureau_s1(st_day, end_day, out_xls, user, pwd, ip, sid, offname)
                del Bureau_sheet1
                import Bureau_sheet2  # 局務報告-只供本所使用
                Bureau_sheet2.Bureau_s2(st_day, end_day, out_xls, user, pwd, ip, sid, offname)
                del Bureau_sheet2
                import Bureau_sheet4
                Bureau_sheet4.Bureau_s4(st_day, end_day, out_xls, user, pwd, ip, sid)
                del Bureau_sheet4
                import Bureau_sheet5
                Bureau_sheet5.Bureau_s5(st_day, end_day, out_xls, user, pwd, ip, sid)
                del Bureau_sheet5
                import Bureau_sheet6
                Bureau_sheet6.bureau_s6(st_day, end_day, out_xls, user, pwd, ip, sid, offname)
                self.chang_cusorSignal1.emit(True)
                self.show_msg.setText('報表產製結束!!!!')
                QApplication.processEvents()  # 強制更新畫面1120322
                time.sleep(10)
                self.show_msg.setText('請輸入起始日期及終止日期，再按產製！！')
                QApplication.processEvents()  # 強制更新畫面1120322

def main():
    import sys, os
    basedir = os.path.dirname(__file__)
    app = QApplication(sys.argv)
#    app.setWindowIcon(QIcon('daan.ico'))
    app.setWindowIcon(QIcon(os.path.join(basedir,'daan.ico')))
    Form = Window()
    Form.ui.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()