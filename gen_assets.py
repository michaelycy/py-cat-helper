# -*- coding: UTF-8 -*-
# ! /usr/bin/python3

# 这段程序可将图标gen.ico转换成icon.py文件里的base64数据
import base64

open_icon = open('assets/logo32x32.ico', "rb")
b64str = base64.b64encode(open_icon.read())
open_icon.close()
write_data = "img = %s" % b64str
f = open("logo_icon.py", "w+")
f.write(write_data)
f.close()

open_icon = open('assets/bg.jpg', "rb")
b64str = base64.b64encode(open_icon.read())
open_icon.close()
write_data = "img = %s" % b64str
f = open("background.py", "w+")
f.write(write_data)
f.close()
