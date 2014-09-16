__author__ = "anjing"
#coding=utf-8

from xlrd import *
from xlwt import *

from Tkinter import *
from tkFileDialog import *
from tkMessageBox import *
import os
from ConfigParser import *
import sys

reload(sys)
sys.setdefaultencoding("utf-8")

def readConfig():
    if not os.path.exists("./config.ini"):
        cfgfile = open("./config.ini", "w")

        cf = ConfigParser()
        cf.add_section("ExcelFilePath")
        cf.set("ExcelFilePath", "path", "")
        cf.add_section("Field")

        cf.set("Field", "Field_0", "")

        cf.add_section("Query")
        cf.set("Query", "key", "")

        cf.write(cfgfile)
        cfgfile.close()

        return

    try:
        cf = ConfigParser()
        cf.read("./config.ini")

        filepath.set(cf.get("ExcelFilePath", "path"))

        for k, v in cf.items("Field"):
            if v.strip():
                field.insert(END, v)

        query.set(cf.get("Query", "key"))
    except:
        pass

def writeConfig():

    cf = ConfigParser()
    cf.add_section("ExcelFilePath")
    cf.set("ExcelFilePath", "path", unicode(filepath.get()))
    cf.add_section("Field")

    ft = field.get(0, END)

    for i in range(len(ft)):
        cf.set("Field", "Field_"+str(i), unicode(ft[i]))

    cf.add_section("Query")
    cf.set("Query", "key", unicode(query.get()))

    cf.write(open("./config.ini", "w"))

    sys.exit()

root = Tk()

root.title(u"课表核对")
root.resizable(width=False, height=False)

pathframe = LabelFrame(root, text=u"路径设置")

pathtip = Label(pathframe,text=u"文件路径: ")
pathtip.pack(side=LEFT)

filepath = StringVar()
pathinput = Entry(pathframe, width=50, textvariable=filepath)
pathinput.pack(side=LEFT, fill=X, expand=1, padx=2)

def filebrowse(target):
    target.set("") #清空entry里面的内容
    #调用filedialog模块的askdirectory()函数去打开文件夹
    fp = askdirectory()
    if fp:
        target.set(fp+"/") #将选择好的路径加入到entry里面

pathbtn = Button(pathframe, text=u"浏览", command=lambda: filebrowse(filepath))
pathbtn.pack(side=LEFT, padx=2)

operateframe = Frame(root)

filedframe = LabelFrame(operateframe, text=u"内容选择")

rowframe = Frame(filedframe)
rowframe.pack(side=TOP)
rowlabel = Label(rowframe, text=u"核对信息: ")
rowlabel.pack(side=LEFT)
title = StringVar()
rowinput = Entry(rowframe, textvariable=title)
rowinput.pack(side=LEFT, fill=X, expand=1)

def insertfield(event):

    if not title.get().strip():
        return

    for x in field.get(0, END):
        if title.get().strip() == x:
            return

    field.insert(END, title.get())

rowinput.bind("<Return>", insertfield)

fieldname = StringVar()
field = Listbox(filedframe, listvariable=fieldname, selectmode=SINGLE)
field.pack(side=TOP, fill=BOTH, expand=1)

def deleteitems(event):
    field.delete(field.curselection()[0])

field.bind("<Delete>", deleteitems)

rowbtn = Button(rowframe, text=u"添加", command=lambda: insertfield(""))
rowbtn.pack(side=LEFT)

queryframe = LabelFrame(operateframe, text=u"查询结果")

inputframe = Frame(queryframe)
inputframe.pack(side=TOP, fill=X, expand=1)
querylabel = Label(inputframe, text=u"查询对象: ")
querylabel.pack(side=LEFT)
query = StringVar()
queryinput = Entry(inputframe, textvariable=query)
queryinput.pack(side=LEFT, fill=X, expand=1)

queryresult = Text(queryframe)
queryresult.pack(side=TOP, fill=X, expand=1)

def scan(key):
    queryresult.insert(END, key+",")

    if not filepath:
        showwarning(u"错误", u"请输入路径！")
        return

    if not fieldname.get():
        showwarning(u"错误", u"请输入查询信息！")
        return

    result = dict()

    for item in field.get(0, END):
        result[item] = ""

    def findrc(sh, i):
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                v = unicode(sh.cell_value(r, c)).replace(" ","")
                if i in v:
                    return (r,c)

    for f in os.listdir(filepath.get()):
        if not (".xls" in f or ".xlsx" in f):
            continue

        p = filepath.get() + f

        try:
            xls = open_workbook(p)
        except:
            continue

        for x in range(xls.nsheets):
            sh = xls.sheets()[x]

            rtemp = findrc(sh, key)
            if not rtemp:
                continue

            row = rtemp[0]

            for info in result.keys():
                result[info] = findrc(sh, info)

            for r in result.keys():
                if not result[r]:
                    continue

                print p

                result[r] = unicode( sh.cell_value(row, result[r][1]) )

                queryresult.insert(END, result[r])
                queryresult.insert(END, ",")

    queryresult.insert(END, "\n")

    return (key, result)

def generatexls(name, result):

    wb = Workbook()

    ws = wb.add_sheet("transcripts")

    titlefont = Font() # Create the Font
    titlefont.name = "SimHei"
    titlefont.bold = True
    titlefont.height = 360 # 18px * 20
    titlealignment = Alignment()
    titlealignment.horz = Alignment.HORZ_CENTER
    titlealignment.vert = Alignment.VERT_CENTER
    titlestyle = XFStyle() # Create the Style
    titlestyle.font = titlefont # Apply the Font to the Style
    titlestyle.alignment = titlealignment
    ws.row(0).height_mismatch = 1
    ws.row(0).height = 27*20
    ws.write_merge(0,0,0,6,u"中国传媒大学硕士研究生成绩单", titlestyle)

    subtitlefont = Font() # Create the Font
    subtitlefont.name = "SimSum"
    subtitlefont.height = 220 # 11px * 20
    subtitlealignment = Alignment()
    subtitlealignment.horz = Alignment.HORZ_CENTER
    subtitlealignment.vert = Alignment.VERT_CENTER
    subtitlestyle = XFStyle() # Create the Style
    subtitlestyle.font = subtitlefont # Apply the Font to the Style
    subtitlestyle.alignment = subtitlealignment
    ws.row(1).height_mismatch = 1
    ws.row(1).height = 25*20
    ws.write_merge(1,1,0,6,u"（本表一式两份，一份存研究生学位档案，一份存个人档案）", subtitlestyle)

    personalinfofont = Font() # Create the Font
    personalinfofont.name = "SimSum"
    personalinfofont.height = 220 # 11px * 20
    personalinfoalignment = Alignment()
    personalinfoalignment.horz = Alignment.HORZ_LEFT
    personalinfoalignment.vert = Alignment.VERT_CENTER
    personalinfostyle = XFStyle() # Create the Style
    personalinfostyle.font = personalinfofont # Apply the Font to the Style
    personalinfostyle.alignment = personalinfoalignment
    ws.row(2).height_mismatch = 1
    ws.row(2).height = 18*20
    ws.write_merge(2,2,0,6,u"   学习期限：自 2011年 9 月至2014年 6月", personalinfostyle)
    ws.row(3).height_mismatch = 1
    ws.row(3).height = 18*20
    ws.write_merge(3,3,0,6,u"   姓名：韩卯辉    学号：113520081002001    院、系：计算机学院", personalinfostyle)
    ws.row(4).height_mismatch = 1
    ws.row(4).height = 18*20
    ws.write_merge(4,4,0,6,u"   专业：信号与信息处理    方向：信号处理技术    导  师：黄祥林", personalinfostyle)

    tabletitlefont = Font() # Create the Font
    tabletitlefont.name = "SimSum"
    tabletitlefont.height = 240 # 12px * 20
    tabletitlefont.bold = True
    tabletitlealignment = Alignment()
    tabletitlealignment.horz = Alignment.HORZ_CENTER
    tabletitlealignment.vert = Alignment.VERT_CENTER
    tabletitlealignment.wrap = True
    borders = Borders()
    borders.top = Borders.THIN
    borders.left = Borders.THIN
    borders.right = Borders.THIN
    borders.bottom = Borders.THIN
    tabletitlestyle = XFStyle() # Create the Style
    tabletitlestyle.font = tabletitlefont # Apply the Font to the Style
    tabletitlestyle.alignment = tabletitlealignment
    tabletitlestyle.borders = borders
    ws.row(5).height_mismatch = 1
    ws.row(5).height = 19*20
    ws.row(6).height_mismatch = 1
    ws.row(6).height = 29*20
    ws.write_merge(5,6,0,0,u"序号",tabletitlestyle)
    ws.write_merge(5,6,1,1,u"课程名称",tabletitlestyle)
    ws.write_merge(5,6,2,2,u"课程类别",tabletitlestyle)
    ws.write_merge(5,6,3,3,u"学分",tabletitlestyle)
    ws.write_merge(5,5,4,6,u"开课学期及成绩",tabletitlestyle)
    ws.write(6,4,u"第一学期",tabletitlestyle)
    ws.write(6,5,u"第二学期",tabletitlestyle)
    ws.write(6,6,u"第三学期",tabletitlestyle)

    ws.col(0).width_mismatch = 1
    ws.col(0).width = 4*256
    ws.col(1).width_mismatch = 1
    ws.col(1).width = 29*256
    ws.col(2).width_mismatch = 1
    ws.col(2).width = 13*256
    ws.col(3).width_mismatch = 1
    ws.col(3).width = 5*256
    ws.col(4).width_mismatch = 1
    ws.col(4).width = 8*256
    ws.col(5).width_mismatch = 1
    ws.col(5).width = 8*256
    ws.col(6).width_mismatch = 1
    ws.col(6).width = 8*256

    for i in range(1,20):
        ws.row(6+i).height_mismatch = 1
        ws.row(6+i).height = 23*20

        ws.write(6+i,0,i,easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
        ws.write(6+i,1,u"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
        ws.write(6+i,2,u"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
        ws.write(6+i,3,2,easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
        ws.write(6+i,4,u"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
        ws.write(6+i,5,u"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
        ws.write(6+i,6,u"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))

    wb.save("transcripts.xls")


def queryprocess():
    if not query.get().strip():
        return

    querylist = query.get().strip().split(",")

    for q in querylist:
        generatexls(q, scan(q))

querybtn = Button(inputframe, text=u"查询", command=queryprocess)
querybtn.pack(side=LEFT)

filedframe.pack(side=LEFT, fill=BOTH, expand=1, anchor=W)
queryframe.pack(side=RIGHT, anchor=E)

pathframe.pack(side=TOP, fill=X, expand=1, padx=5, pady=2)
operateframe.pack(side=TOP, fill=X, expand=1, padx=5, pady=2)

root.wm_protocol("WM_DELETE_WINDOW", writeConfig)

readConfig()

root.mainloop()
