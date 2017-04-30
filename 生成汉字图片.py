# -*- coding: utf-8 -*-

import sys
import xlrd
import pygame
from pypinyin import pinyin
import math
import os
import wx

reload(sys)
sys.setdefaultencoding('utf-8')

content_width = 48 * 5
content_height = 80 * 5

card_size_height = 488
card_size_width = 300
big_gezi_size = 60 * 3
big_gezi_x = (card_size_width - big_gezi_size) / 2
big_gezi_y = (card_size_width - big_gezi_size) / 3

perline_number = 4
line_number = 3
small_gezi_size = 60
rect_x = (card_size_width - small_gezi_size * 4) / 2
rect_y = big_gezi_y + big_gezi_size + (content_height - big_gezi_size - line_number * small_gezi_size) / 2


def open_excel(file=r'C:/Users/123/Desktop/Python/1352-1426.xlsx'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception, e:
        print str(e)


def excel_table_byname(filepath, colnameindex=0, by_name=u'Sheet1'):
    data = open_excel(filepath)
    table = data.sheet_by_name(by_name)
    nrows = table.nrows  # 行数
    ncols = table.ncols  # 列数
    colnames = table.row_values(colnameindex)  # 某一行数据
    list = []
    for rownum in range(0, nrows):
        row = table.row_values(rownum)
        if row:
            app = []
            # for i in range(len(colnames)):
            # app = (row[2], row[3], row[4], row[5], row[6], row[7], row[8])
            for column in range(0, ncols):  # 把所有列都显示出来
                value = str(row[column])
                if value.strip():  # 去掉空格
                    app.append(row[column])
            list.append(app)
    return list


def draw_xuxian(root, start_x, start_y, width, height):
    xiebian = math.sqrt(width * width + height * height)
    xuxian_part = int(xiebian / 20)
    xielv = height / width
    points = []
    # 画\虚线
    for x in range(start_x, start_x + width + xuxian_part, xuxian_part):
        y = -start_y / 2 + xielv * x
        points.append([x, y])
        if len(points) == 2:
            pygame.draw.lines(root, [144, 144, 144], False, points, 1)
            points = []
    points = []
    # 画/虚线
    for y in range(start_y, start_y + height + xuxian_part, xuxian_part):
        # x = (y - ((width*height/y + x+width)))/(-xielv) 有问题
        x = start_x + width - (y - start_y)
        points.append([x, y])
        if len(points) == 2:
            pygame.draw.lines(root, [144, 144, 144], False, points, 1)
            points = []
    points = []
    # 画|虚线
    for y in range(start_y, start_y + height + xuxian_part, xuxian_part):
        points.append([start_x + width / 2, y])
        if len(points) == 2:
            pygame.draw.lines(root, [144, 144, 144], False, points, 1)
            points = []
    points = []
    # 画-虚线
    for x in range(start_x, start_x + width + xuxian_part, xuxian_part):
        points.append([x, start_y + height / 2])
        if len(points) == 2:
            pygame.draw.lines(root, [144, 144, 144], False, points, 1)
            points = []


# 画拼音_序号
def draw_pingyin_number(root, word, number, fontpath):
    font = pygame.font.Font(fontpath, 20)
    text = pinyin(word.decode('utf-8'))[0][0]
    rtext = font.render(text, True, (0, 0, 0))
    rtext2 = font.render(number, True, (0, 0, 0))
    root.blit(rtext, ((card_size_width - rtext.get_width()) / 2, big_gezi_y - rtext.get_height()))
    root.blit(rtext2, ((card_size_width - rtext2.get_width()) / 2, big_gezi_y + content_height - 20))


def draw_gezi_and_word(root, outline_width, start_x, start_y, gezi_width, getzi_height, word, font_path, font_size,
                       outline_color):
    pygame.draw.rect(root, outline_color, (start_x, start_y, gezi_width, getzi_height), outline_width)
    draw_word(root, word, start_x, start_y, gezi_width, getzi_height, font_path, font_size)


def draw_word(root, word, start_x, start_y, gezi_width, getzi_height, font_path, font_size):
    font = pygame.font.Font(font_path, font_size)
    text = font.render(word.decode("gbk", "ignore"), True, (0, 0, 0))
    x = (gezi_width - text.get_width()) / 2 + start_x
    y = (getzi_height - text.get_height()) / 2 + start_y
    root.blit(text, (x, y))


def draw_init():
    pygame.init()
    root = pygame.Surface((card_size_width, card_size_height))
    # 背景填充
    root.fill((255, 255, 255))
    # 画虚线
    draw_xuxian(root, (card_size_width - big_gezi_size) / 2, (card_size_width - big_gezi_size) / 3, big_gezi_size,
                big_gezi_size)
    # 画部件大矩形，可以省略此矩形
    # draw_bigRect(root)
    return root


def save_to_png(root, file):
    pygame.image.save(root, file)


def draw_bigRect(root):
    pygame.draw.lines(root, (0, 0, 0), True, [(rect_x, rect_y), (rect_x + small_gezi_size * 4, rect_y),
                                              (rect_x + small_gezi_size * 4, rect_y + small_gezi_size * 3),
                                              (rect_x, rect_y + small_gezi_size * 3)])
    for i in range(0, 4):
        pygame.draw.lines(root, (0, 0, 0), False, [(rect_x + small_gezi_size * i, rect_y),
                                                   (rect_x + small_gezi_size * i, rect_y + small_gezi_size * 3)], 1)

    for i in range(0, 3):
        pygame.draw.lines(root, (0, 0, 0), False, [(rect_x, rect_y + small_gezi_size * i),
                                                   (rect_x + small_gezi_size * 4, rect_y + small_gezi_size * i)], 1)


# 把部件,循环调用draw_gezi_and_word方法，x y坐标要设计好,大小每个60像素，每行4个，card_width为300
def draw_bushou(root, file, row):
    busou_num = 2
    root.set_alpha(255)
    for t in range(0, 3):
        for i in range(0, 4):
            if busou_num >= len(row):
                busou = ""
            else:
                busou = row[busou_num].strip().lstrip()
            draw_gezi_and_word(root, 1, i * small_gezi_size + (card_size_width - small_gezi_size * perline_number) / 2,
                               rect_y + small_gezi_size * t, small_gezi_size, small_gezi_size, busou, file, 40,
                               (250, 250, 250))

            busou_num = busou_num + 1


# 生成所有汉字的图片
def All_create(excel, out):
    # print("𧘇")
    # excel_path=raw_input(r"please input the excel file path:").strip()
    # out_path = raw_input(r"please input output_file path:").strip()
    excel_path = excel
    out_path = out
    if (os.path.exists(out_path) == False):
        os.makedirs(out_path)
    tables = excel_table_byname(str(excel_path))
    for row in tables:

        filename = (str(row[0]))

        filename = filename[0:filename.index('.')]
        # 先初始化画布
        root = draw_init()
        # 画拼音

        draw_pingyin_number(root, row[1], filename, "C:\Windows\Fonts\simsun.ttc")
        # 画格子+画字
        try:
            draw_gezi_and_word(root, 2, big_gezi_x, big_gezi_y, big_gezi_size, big_gezi_size, row[1],
                               "C:\Windows\Fonts\simsun.ttc", 140, (0, 0, 0))
            # 把部件,循环调用draw_gezi_and_word方法，x y坐标要设计好,大小每个50像素，每行4个，card_width为300
            draw_bushou(root, "C:\Windows\Fonts\simsun.ttc", row)
            draw_bigRect(root)
        except Exception, e:
            draw_gezi_and_word(root, 2, big_gezi_x, big_gezi_y, big_gezi_size, big_gezi_size, row[1],
                               "C:\Windows\Fonts\simsun.ttc", 140, (0, 0, 0))
            # 把部件,循环调用draw_gezi_and_word方法，x y坐标要设计好,大小每个50像素，每行4个，card_width为300
            draw_bushou(root, "C:\Windows\Fonts\simsun.ttc", row)
            draw_bigRect(root)
        save_to_png(root, out_path + filename + ".png")


# 生成单个汉字的图片
def One_Create(excel, out, word):
    return


# UI类，
def One_create(param, param1, param2):
    pass


class MainWindow(wx.Frame):
    """We simply derive a new class of Frame."""

    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=title, pos=wx.DefaultPosition, size=wx.Size(700, 425),
                          style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)
        # self.control = wx.TextCtrl(self, style = wx.TE_MULTILINE)#文本框

        self.CreateStatusBar()  # 创建位于窗口的底部的状态栏

        # 设置菜单
        filemenu = wx.Menu()

        # wx.ID_ABOUT和wx.ID_EXIT是wxWidgets提供的标准ID
        menuAbout = filemenu.Append(wx.ID_ABOUT, u"关于", u"关于程序的信息")
        filemenu.AppendSeparator()
        menuExit = filemenu.Append(wx.ID_EXIT, u"退出", u"终止应用程序")

        # 创建菜单栏
        menuBar = wx.MenuBar()
        menuBar.Append(filemenu, u"菜单")
        self.SetMenuBar(menuBar)

        # 设置events
        self.Bind(wx.EVT_MENU, self.OnAbout, menuAbout)
        self.Bind(wx.EVT_MENU, self.OnExit, menuExit)

        # 设置sizers
        self.SetSizeHintsSz(wx.DefaultSize, wx.DefaultSize)

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        sbSizer1 = wx.StaticBoxSizer(wx.StaticBox(self, wx.ID_ANY, wx.EmptyString), wx.VERTICAL)

        sbSizer1.SetMinSize(wx.Size(-1, 280))
        bSizer3 = wx.BoxSizer(wx.HORIZONTAL)

        self.excelfile_tv = wx.StaticText(self, wx.ID_ANY, u"Excel文件路径：", wx.DefaultPosition, wx.DefaultSize, 0)
        self.excelfile_tv.Wrap(-1)
        bSizer3.Add(self.excelfile_tv, 0, wx.ALL, 5)

        self.excelfile_et = wx.FilePickerCtrl(self, wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.*",
                                              wx.DefaultPosition, wx.Size(1000, -1), wx.FLP_DEFAULT_STYLE)
        bSizer3.Add(self.excelfile_et, 0, wx.ALL, 5)

        sbSizer1.Add(bSizer3, 1, wx.EXPAND, 5)

        bSizer5 = wx.BoxSizer(wx.HORIZONTAL)

        self.outpath_tv = wx.StaticText(self, wx.ID_ANY, u"输出文件路径： ", wx.DefaultPosition, wx.DefaultSize, 0)
        self.outpath_tv.Wrap(-1)
        bSizer5.Add(self.outpath_tv, 0, wx.ALL, 5)

        self.outpath_et = wx.DirPickerCtrl(self, wx.ID_ANY, wx.EmptyString, u"Select a folder", wx.DefaultPosition,
                                           wx.Size(1000, -1), wx.DIRP_DEFAULT_STYLE)
        bSizer5.Add(self.outpath_et, 0, wx.ALL, 5)

        sbSizer1.Add(bSizer5, 1, wx.EXPAND, 5)

        bSizer6 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_checkBox1 = wx.CheckBox(self, wx.ID_ANY, u"是否只生成特定字的图片", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer6.Add(self.m_checkBox1, 0, wx.ALL, 5)

        self.word_tv = wx.StaticText(self, wx.ID_ANY, u"若是则输入单个汉字：", wx.DefaultPosition, wx.DefaultSize, 0)
        self.word_tv.Wrap(-1)
        bSizer6.Add(self.word_tv, 0, wx.ALL, 5)

        self.word_et = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer6.Add(self.word_et, 0, wx.ALL, 5)

        sbSizer1.Add(bSizer6, 1, wx.EXPAND, 5)

        bSizer7 = wx.BoxSizer(wx.VERTICAL)

        sbSizer1.Add(bSizer7, 1, wx.EXPAND, 5)

        bSizer1.Add(sbSizer1, 1, wx.EXPAND, 5)

        sbSizer2 = wx.StaticBoxSizer(wx.StaticBox(self, wx.ID_ANY, wx.EmptyString), wx.VERTICAL)

        sbSizer2.SetMinSize(wx.Size(-1, 50))
        self.begin_bt = wx.Button(self, wx.ID_ANY, u"开始生成", wx.DefaultPosition, wx.DefaultSize, 0)
        sbSizer2.Add(self.begin_bt, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        bSizer1.Add(sbSizer2, 1, wx.EXPAND, 5)

        self.SetSizer(bSizer1)
        self.Layout()

        self.Centre(wx.BOTH)

        # 设置按钮event
        self.begin_bt.Bind(wx.EVT_BUTTON, self.OnBeginClick)

        self.Show(True)  # 显示这个frame

    def OnAbout(self, e):
        # 创建一个带"OK"按钮的对话框。wx.OK是wxWidgets提供的标准ID
        dlg = wx.MessageDialog(self, "transform word to image.", "about", wx.OK)  # 语法是(self, 内容, 标题, ID)
        dlg.ShowModal()  # 显示对话框
        dlg.Destroy()  # 当结束之后关闭对话框

    def OnExit(self, e):
        self.Close(True)  # 关闭整个frame

    # 按钮回调函数
    def OnBeginClick(self, event):
        '''
        print(self.outpath_et.GetPath()+"\\")
        print(self.excelfile_et.GetPath())
        print(self.m_checkBox1.GetValue())
        print(self.word_et.GetValue())
        '''

        if (self.m_checkBox1.GetValue() == False):
            All_create(self.excelfile_et.GetPath(), self.outpath_et.GetPath() + "\\")

        else:
            One_create(self.excelfile_et.GetPath(), self.outpath_et.GetPath() + "\\", self.word_et.GetValue())  # 生成单个汉字
        dlg = wx.MessageDialog(self, "Success!", "message", wx.OK)  # 语法是(self, 内容, 标题, ID)
        dlg.ShowModal()  # 显示对话框
        dlg.Destroy()  # 当结束之后关闭对话框


def main():
    app = wx.App(False)  # 创建1个APP，禁用stdout/stderr重定向
    frame = MainWindow(None, 'WordToImage')  # 这是一个顶层的window
    frame.Show(True)  # 显示这个frame
    app.MainLoop()


main()