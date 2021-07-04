# -*- coding: UTF-8 -*-
# ! /usr/bin/python3

import os
from pathlib import Path
from PIL import Image, ImageTk
from tkinter import Tk, Label, PhotoImage

import setting as config
from helpers.pptx import PPTX
from helpers.excel import Excel
# from helpers.image import gen_tmp
from assets import favicon as logo, bg as background_img
# body_background = gen_tmp(background_img, 'assets/bg.jpg')


class CatHelper(Tk):
    def __init__(self):
        super().__init__()
        self.title("猫咪绝育券生成助手")

        # setting window size
        width = 798
        height = 448
        screenwidth = self.winfo_screenwidth()
        screenheight = self.winfo_screenheight()
        align_str = '%dx%d+%d+%d' % (width,
                                     height,
                                     (screenwidth - width) / 2,
                                     (screenheight - height) / 2)
        self.geometry(align_str)
        self.resizable(width=False, height=False)

        # 设置窗口图标
        # self.iconbitmap(favicon)

        try:
            # # 设置背景图片
            # bg_pil = Image.open(body_background)
            bg_pil = Image.open(os.path.abspath('assets/bg.jpg'))
            image = ImageTk.PhotoImage(bg_pil)

            image_label = Label(self, image=image)
            image_label.image = image
            image_label.place(x=0, y=0, relx=0, rely=0)
        except:
            print('异常')
            pass


if __name__ == "__main__":

    cat_helper = CatHelper()
    cat_helper.iconbitmap(os.path.abspath('favicon.ico'))
    cat_helper.mainloop()
