"""
本程式包含了Bureau_sheet3,Bureau_sheet1,Bureau_sheet2,Bureau_sheet5,Bureau_sheet6,6個子模組,
版本為3.0 修正日期1110210 1110406
"""
import afmod
import afmod1
import pyinputplus as pyip
import os
st_day = str(pyip.inputInt('請輸入起始日期 (XXXXXXX):', max=9999999))
end_day = str(pyip.inputInt('請輸入終止日期(XXXXXXX):', max=9999999))
del pyip
global user, ip, pwd, sid, ora_dir, offname, prtoutfile_path

today, tomorrow, tm_year, tm_mon, tm_mday, tm_hour, tm_min, tm_sec = afmod.get_today()
out_xls = st_day + end_day + str(tm_hour) + str(tm_min) + str(tm_sec)
if os.path.exists('Bureau_reportV3.ini') is False:  # 1110210 add
    print("Bureau_reportV3.ini檔案不存在,請檢查!!!!!")
    exit()
filename = afmod.read_init('Bureau_reportV3.ini')
offname, user, pwd, ip, sid, ora_dir, prtoutfile_path, sample_xls = afmod.read_ini(filename)
if os.path.exists(sample_xls) is False:
    print(filename + "xls範例檔案不存在,請檢查!!!!!")
    exit()
del os
# --------------------------------------------------
ap_str = str(ord('l') + ord('a'))  # 加解密用內含指定字元
user = afmod1.dectry(user, ap_str)  # oracle user
pwd = afmod1.dectry(pwd, ap_str)  # oracle user的pwd
ip = afmod1.dectry(ip, ap_str)  # oracle 的連結ip
sid = afmod1.dectry(sid, ap_str)  # oracle 的連結名稱
afmod.con_ora_path(ora_dir, 'A')  # 讀取連結oracle lib的目錄
# ---------------------------------------------------
out_xls = prtoutfile_path + out_xls + offname + '.xlsx'
import Bureau_sheet3  # 他縣市承辦本所案件資料
Bureau_sheet3.Bureau_s3(st_day, end_day, out_xls, user, pwd, ip, sid, sample_xls)
del Bureau_sheet3
import Bureau_sheet1
Bureau_sheet1.Bureau_s1(st_day, end_day, out_xls, user, pwd, ip, sid, offname)
del Bureau_sheet1
import Bureau_sheet2 #局務報告-只供本所使用
Bureau_sheet2.Bureau_s2(st_day, end_day, out_xls, user, pwd,ip,sid)
del Bureau_sheet2
import Bureau_sheet4
Bureau_sheet4.Bureau_s4(st_day, end_day, out_xls, user, pwd, ip, sid)
del Bureau_sheet4
import Bureau_sheet5
Bureau_sheet5.Bureau_s5(st_day, end_day, out_xls, user, pwd, ip, sid)
del Bureau_sheet5
import Bureau_sheet6
Bureau_sheet6.bureau_s6(st_day, end_day, out_xls, user, pwd, ip, sid, offname)
input(" -------輸入任意鍵以結束--------- ")