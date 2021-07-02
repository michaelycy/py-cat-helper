import os.path as path
import sys

""" 工作表中必须包含的列名 """
HEAD = ('救助人真实姓名 | Real Name', '电话 | Mobile', '医院 | Infirmary', '使用有效期', '券号')

""" PPT 模板绝对路径 """
if getattr(sys, 'frozen', None):
    basedir = sys._MEIPASS
else:
    basedir = path.dirname(__file__)
PPTX_TEMPLATE = path.join(basedir, 'template/coupon_ppt.pptx')
