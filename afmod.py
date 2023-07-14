import time
import datetime
import cx_Oracle
from sqlalchemy import create_engine
import os
import binascii
import configparser
import openpyxl
# 取得今天日期及明天日期
def get_today():
    tm_year, tm_mon, tm_mday, tm_hour, tm_min, tm_sec, tm_wday, tm_yday, tm_isdst = time.localtime()
    tday1 = datetime.date.today()
    tday = tday1.strftime("%Y%m%d")
    tomorrow = tday1 + datetime.timedelta(days=1)
    tomorrow = tomorrow.strftime("%Y%m%d")
    today = int(tday) - 19110000
    tomorrow = int(tomorrow) - 19110000
    return today, tomorrow, tm_year, tm_mon, tm_mday, tm_hour, tm_min, tm_sec
# 設定oracle lib目錄
def con_ora_path(ora_dir,flg):
    if flg =='A':
        if os.path.exists(ora_dir) is False:
            print(ora_dir +"目錄不存在,請檢查!!!!!")
            exit()
        else:
            cx_Oracle.init_oracle_client(lib_dir=ora_dir)
    elif flg == 'M':
        if os.path.exists(ora_dir) is False:
            mkdir(ora_dir)
    elif flg == 'C':
        if os.path.exists(ora_dir) is False:
            return False
    # cx_Oracle.init_oracle_client(lib_dir=r"d:\oracle11\instantclient_11_2")

#  連結資料庫及取資料
def con_ora(str_sql,user,pwd,ip,sid):
    dsn = cx_Oracle.makedsn(ip, 1521, service_name=sid)
    connection = cx_Oracle.connect(user=user, password=pwd,dsn=dsn)
    # connection = cx_Oracle.connect("moicad", "moicadcoco", "new_afweb")
    cursorobj = connection.cursor()
    cur48 = cursorobj.execute(str_sql)
    data = cursorobj.fetchall()
    cursorobj.close()
    return data
# 用cx-oracle-sqlalchemy連結資料庫
def con_ora_sqlalchemy(user, pwd, ip, sid):
    con = user + ':' + pwd +'@' + ip + ':'+'1521' + '/' + sid
    engine = create_engine("oracle+cx_oracle://" + con, max_identifier_length=128)
    return engine
# 轉換中文
def chg_big(data1):
    for k1, k2 in enumerate(data1):
        tmp = ' '
        if k2[1] is None:
            tmp = '空白'
        else:
            try:
                tmp = binascii.unhexlify(k2[1]).decode('BIG5')
                # print(tmp)
            except UnicodeDecodeError:
                tmp = "轉碼錯誤"
                continue
        return tmp
# 讀取ini設定檔資料
def read_ini(filename):
    curpath = os.path.dirname(os.path.realpath(__file__))
    cfgpath = os.path.join(curpath, filename)
    config = configparser.ConfigParser()
    config.read(cfgpath)
    offname = config['office_name']['officname']
    ora_dir = config['ora_dir']['ora_dir']
    user = config['ora_user']['usr']
    pwd = config['ora_user']['pwd']
    ip = config['ora_user']['ip']
    sid = config['ora_user']['sid']
    prtoutfile_path = config['prtoutfile_path']['out_path']
    sample_file = config['sample_file']['sample_xlxs']
    return offname, user, pwd, ip, sid, ora_dir, prtoutfile_path, sample_file
# 檢查ini檔存不存在
def read_init(filename):
    filepath = os.getcwd()
    filename = filepath + "\\" +filename
    if os.path.exists(filename) is False:
        print(filname + "設定檔不存在,請檢查!!!!!")
        exit()
    return filename
# ----------------------將中文(-8)轉為16進位
def utf8_to_big5(utf_str):
    utf_str1 = utf_str.encode('BIG5')
    utf_str1 = str(utf_str1.hex()).upper()  # 轉為大寫
    return utf_str1
# -----------------------
def create_xlsx(xlsfile):
    wb = openpyxl.Workbook()
    wb.save(xlsfile)
