# -*- coding: UTF-8 -*-
# ! /usr/bin/python3

# 这段程序可将图标gen.ico转换成icon.py文件里的base64数据
import ntpath
from base64 import b64encode
from os.path import splitext


def image_to_base64_code(filepath):
    """图片转 Base64 字符串

    Args:
        filepath (str): 图片路径

    Returns:
        str: base64 字符串
    """
    print(filepath)
    file = open(filepath, 'rb')
    container_text = file.read()
    code = b64encode(container_text)
    file.close()

    return code


# 图片资源列表
# 'assets/success16x16.png', 'assets/error16x16.png'
images = ['assets/favicon.ico', 'assets/bg.jpg']

text = ''
for img in images:
    base64_code = image_to_base64_code(img)
    filename = splitext(ntpath.basename(img))[0]
    text = text + "%s = %s \n\n" % (filename, base64_code)

file = open('assets.py', 'w+')
file.write(text)
file.close()
