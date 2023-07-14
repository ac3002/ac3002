def Bureau_s5(st_day,end_day,out_xls,user,pwd,ip,sid):
    import afmod
    from openpyxl import load_workbook
#--------------------------------------------------------
    str_sql99N = """
          SELECT DISTINCT aa108,ac20,Sum(ac30) as sum30
           FROM sexpaa,sexpac
             WHERE aa108 <>'A' AND aa25 = ac25  AND aa08 = '1'  AND aa04 = ac04  AND ac20 in ('33','34')  
             AND aa01 >= """
    str_sql = str_sql99N + "\'" + st_day +"\'"+" AND aa01<= "+ "\'" + end_day +"\'"+" group  by aa108,ac20 "
  #  print(str_sql)
    data = afmod.con_ora(str_sql,user,pwd,ip,sid) #連結oracle取得資料
    citys = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'M', 'N', 'O', 'P', 'Q', 'T', 'U', 'V', 'W', 'X', 'Z']
 #   out_xls = prtoutfile_path+out_xls+offname +'.xlsx'
    wb = load_workbook(filename=out_xls)
    on_sheet5 = wb['跨縣市書狀、登記費統計表']
    for i in range(0,len(data)):
        ac108, pe21,sum_ac30 = data[i]
        if pe21 == '34':
            pe21_name = ac108 + '-34'
            cellRange = on_sheet5['C2':'Y25']
            for row in cellRange:
                for c in row:
                    if c.value == "report_time":
                        c.value = st_day + '起至' + end_day + '止'
                    if c.value == pe21_name:
                        c.value = sum_ac30
        if pe21 == '33':
            pe21_name = ac108 + '-33'
            cellRange = on_sheet5['B27':'Y55']
            for row in cellRange:
                for c in row:
                    if c.value == "report_time":
                        c.value = st_day + '起至' + end_day + '止'
                    if c.value == pe21_name:
                        c.value = sum_ac30
    for city in citys:
        tmp_str = city + '-34'
        tmp_str1 = city + '-33'
        cellRange = on_sheet5['C4':'Y60']
        for row in cellRange:
            for c in row:
                if (c.value == tmp_str) or (c.value == tmp_str1):
                    c.value = 0
#------------------------------------------------------
    wb.save(filename=out_xls)

