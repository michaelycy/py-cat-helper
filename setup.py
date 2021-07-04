# -*- coding: UTF-8 -*-
# ! /usr/bin/python3

import os
from pathlib import Path
from PIL import Image, ImageTk
from tkinter.ttk import Button, Style
from tkinter.messagebox import showinfo
from tkinter.filedialog import askopenfilename
from tkinter import DISABLED, NORMAL, NONE, SUNKEN
from tkinter import StringVar, Tk, Label,  Listbox, Entry

import setting as config
from helpers.pptx import PPTX
from helpers.excel import Excel
from helpers.image import gen_tmp
from assets import favicon as logo, bg as background_img

text_color = '#000'
white = '#fff'
primary_color = '#1890ff'
form_item_height = 32
font_family = '微软雅黑'
# favicon = gen_tmp(logo, 'assets/favicon.ico')
# body_background = gen_tmp(background_img, 'assets/bg.jpg')


class CatHelper(Tk):
    def __init__(self):
        super().__init__()
        self.excel = None
        self.active_sheet_name = None
        # 文件名称
        self.filename = None
        self.head_name_column = None
        self.list_box = None
        self.build_btn = None
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

        # # 设置全局样式
        style = Style()
        style.configure(
            "ku.TButton",
            font=(font_family, 14),
            anchor="center",
            borderwidth='4',
            activebackground='#F3F4F5',
            # background='#F3F4F5'
            bg=primary_color,
        )
        style.map(
            'ku.TButton',
            foreground=[('active', '!disabled', primary_color)])

        # 设置背景图片
        # bg_pil = Image.open(body_background)
        # image = ImageTk.PhotoImage(bg_pil)
        # image_label = Label(self, image=image)
        # image_label.image = image
        # image_label.place(x=0, y=0, relx=0, rely=0)

        self.upload_filepath_input_value = StringVar()
        self.upload_filepath_input_value.set('请上传申请表格')
        # 上传模板路径地址
        self.upload_filepath_input = Entry(
            self,
            state=DISABLED,
            relief=SUNKEN,
            textvariable=self.upload_filepath_input_value)
        self.upload_filepath_input.config(
            fg='gray',
            bg='#F3F4F5',
            font=(font_family, 12),
            borderwidth=1,
            insertbackground='#fff',

            highlightbackground='#F3F4F5')

        # 上传按钮
        self.upload_button = Button(
            self,
            text="上传 Excel",
            command=self.start,
            style='ku.TButton'
        )

        # 装载
        self.upload_filepath_input.place(x=285, y=180, width=300, height=32)
        self.upload_button.place(x=590, y=180, height=32)

    def start(self):
        filepath = askopenfilename(
            filetypes=[('Excel files', '*.xlsx')])

        # 若用户选择 excel 时，点击取消按钮
        if len(filepath) == 0:
            print('用户取消操作')
            return

        self.upload_filepath_input_value.set(filepath)
        self.excel = Excel(filepath)
        self.filename = os.path.splitext(Path(filepath).name)[0]
        self.active_sheet_name = sheet_name = self.excel.sheet_names[0]
        thead = self.excel.get_thead(sheet_name)
        valid = self.check_excel_format(thead)
        self.build_btn = Button(
            self,
            text='生成 PPT',
            command=self.build_pptx,
            state=DISABLED,
            style='ku.TButton'
        )

        if valid:
            self.build_btn['state'] = NORMAL

        self.build_btn.place(x=450, y=335, height=32)

    def check_excel_format(self, head):
        f"""
        校验工作表列名
        :param head: {dict[str, int]}
        :return: {bool}
        """
        valid = True
        head_names = list(head.keys())
        self.list_box = Listbox(
            self,
            borderwidth=0,
            activestyle=NONE,
            bg='#F3F4F5')
        # 工作表中 {列名：索引} map
        head_column_map = {}

        for preset_name in config.HEAD:
            for i in range(0, len(head_names)):
                # name = head_names[i] if  type(head_names[i])  is None else ''
                head_name = head_names[i] if isinstance(
                    head_names[i], str) else ''

                if head_name.find(preset_name) == 0:
                    # 因为工作表中 column 是从 1 开始的，但是 for 循环是从 0 开始的
                    head_column_map[preset_name] = i + 1
                    continue
        # 记录
        self.head_name_column = head_column_map
        head_keys = list(head_column_map.keys())
        preset_head_count = len(list(config.HEAD))
        for i in range(0, preset_head_count):
            preset_name = config.HEAD[i]

            if preset_name in head_keys:
                self.list_box.insert(i, preset_name + ' ✅')
            else:
                self.list_box.insert(i, preset_name + ' ❌')
                valid = False

        self.list_box.place(x=285, y=225, width=405, height=95)
        return valid

    def build_pptx(self):
        pptx = PPTX()
        rows = list(self.excel.get_rows(self.active_sheet_name))
        print('工作表中数据条数：', len(rows))
        pptx.copy_slide(len(rows))

        # 遍历表格数据，并更新幻灯片
        for index in range(2, len(rows) + 1):
            # 券号
            case_code = self.excel.get_cell(
                self.active_sheet_name, index, self.head_name_column[config.HEAD[4]])
            # 申请人：救助人真实姓名 + ' ' + 电话
            rescuer = self.excel.get_cell(
                self.active_sheet_name, index, self.head_name_column[config.HEAD[0]])
            rescuer_phone = self.excel.get_cell(
                self.active_sheet_name, index, self.head_name_column[config.HEAD[1]])

            # 申请医院
            rescuer_place = self.excel.get_cell(
                self.active_sheet_name, index, self.head_name_column[config.HEAD[2]])
            # 使用有效期
            use_exp_date = self.excel.get_cell(
                self.active_sheet_name, index, self.head_name_column[config.HEAD[3]])
            # 设置幻灯片内容
            pptx.set_data_with_slide(
                index=index - 1,
                case_code=case_code,
                rescuer=rescuer,
                rescuer_phone=rescuer_phone,
                rescuer_place=rescuer_place,
                use_exp_date=use_exp_date)

        # 关闭 Excel
        self.excel.close()
        # 保存 PPT 到桌面
        pptx.save(
            os.path.join(
                os.path.join(
                    os.path.expanduser("~"),
                    'Desktop'),
                self.filename +
                '.ppt'))
        showinfo('提示', 'PPT 已生成！')

        # 重置状态
        self.list_box.destroy()
        self.build_btn.destroy()
        self.upload_filepath_input_value.set('请上传申请表格')


if __name__ == "__main__":

    cat_helper = CatHelper()
    cat_helper.mainloop()
