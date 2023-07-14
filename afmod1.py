#加密
def enctry(s,A):#加密
    import random
    k = '1AX2LH5aX6lA7vq04z-%l>l;a2=BY$MF3bX4mS%xw^5x8^q9<u&bw*CZ(ND)cXnDyeq6cw&aeXrcs@EFP3aeXspFPYr{8v}(zdCyjex'
    oth_str =str(random.randrange(1,9))
    for i in range(1,5):
        oth_str = oth_str + str(random.randrange(1000,9999,3))
    encry_str = ""
    for i,j in zip(s,k):
        # i為字元，j為祕鑰字元
        temp = str(ord(i)+ord(j))+A # 加密字元 = 字元的Unicode碼 + 祕鑰的Unicode碼
        encry_str = encry_str + temp
    lens = str(9999 -len(encry_str))
    encry_str = lens + encry_str + oth_str
    return encry_str
# 解密
def dectry(p,A):# 解密
    k = '1AX2LH5aX6lA7vq04z-%l>l;a2=BY$MF3bX4mS%xw^5x8^q9<u&bw*CZ(ND)cXnDyeq6cw&aeXrcs@EFP3aeXspFPYr{8v}(zdCyjex'
    dec_str = ""
    tmp_p = p[:4]
    tmp_p1 = 9999 - int(tmp_p)
    p = p[4:tmp_p1 +4]
    for i,j in zip(p.split(A)[:-1],k):
        # i 為加密字元，j為祕鑰字元
        temp = chr(int(i) - ord(j)) # 解密字元 = (加密Unicode碼字元 - 祕鑰字元的Unicode碼)的單位元組字元
        dec_str = dec_str+temp
    return dec_str
#--------
def set_rm13(srm13,srm16):
    if srm13 is None:
        srm13 = 0
    if srm16 is None:
       srm16 = 0
    return srm13,srm16
