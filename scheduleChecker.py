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
import string
import re

reload(sys)
sys.setdefaultencoding("utf-8")

def readConfig():
    if not os.path.exists("./config.ini"):
        cfgfile = open("./config.ini", "w")

        cf = ConfigParser()
        cf.add_section("ExcelFilePath")
        cf.set("ExcelFilePath", "score", "")
        cf.set("ExcelFilePath", "name", "")
        cf.set("ExcelFilePath", "result", "")
        # cf.add_section("Field")
        # cf.set("Field", "Field_0", "")
        cf.add_section("Query")
        cf.set("Query", "key", "")

        cf.write(cfgfile)
        cfgfile.close()

        return

    try:
        cf = ConfigParser()
        cf.read("./config.ini")

        transcriptpath.set(cf.get("ExcelFilePath", "score"))
        infopath.set(cf.get("ExcelFilePath", "name"))
        resultpath.set(cf.get("ExcelFilePath", "result"))
        # for k, v in cf.items("Field"):
        #     if v.strip():
        #         field.insert(END, v)

        query.set(cf.get("Query", "key"))
    except:
        pass

def writeConfig():

    cf = ConfigParser()
    cf.add_section("ExcelFilePath")
    cf.set("ExcelFilePath", "score", unicode(transcriptpath.get()))
    cf.set("ExcelFilePath", "name", unicode(infopath.get()))
    cf.set("ExcelFilePath", "result", unicode(resultpath.get()))
    # cf.add_section("Field")
    # ft = field.get(0, END)
    # for i in range(len(ft)):
    #     cf.set("Field", "Field_"+str(i), unicode(ft[i]))

    cf.add_section("Query")
    cf.set("Query", "key", unicode(query.get()))

    cf.write(open("./config.ini", "w"))

    sys.exit()

root = Tk()

root.title(u"课表核对")
root.resizable(width=False, height=False)

pathframe = LabelFrame(root, text=u"路径设置")

def filebrowse(target):
    target.set("") #清空entry里面的内容
    #调用filedialog模块的askdirectory()函数去打开文件夹
    fp = askdirectory()
    if fp:
        target.set(fp+"/") #将选择好的路径加入到entry里面

transcriptframe = Frame(pathframe)
transcripttip = Label(transcriptframe,text=u"成绩路径: ")
transcripttip.pack(side=LEFT)
transcriptpath = StringVar()
transcriptinput = Entry(transcriptframe, width=50, textvariable=transcriptpath)
transcriptinput.pack(side=LEFT, fill=X, expand=1, padx=2)
transcriptbtn = Button(transcriptframe, text=u"浏览", command=lambda: filebrowse(transcriptpath))
transcriptbtn.pack(side=LEFT, padx=2)
transcriptframe.pack(side=TOP, fill=X, expand=1)

infoframe = Frame(pathframe)
infotip = Label(infoframe,text=u"名单路径: ")
infotip.pack(side=LEFT)
infopath = StringVar()
infoinput = Entry(infoframe, width=50, textvariable=infopath)
infoinput.pack(side=LEFT, fill=X, expand=1, padx=2)
infobtn = Button(infoframe, text=u"浏览", command=lambda: filebrowse(infopath))
infobtn.pack(side=LEFT, padx=2)
infoframe.pack(side=TOP, fill=X, expand=1)

resultframe = Frame(pathframe)
resulttip = Label(resultframe,text=u"名单路径: ")
resulttip.pack(side=LEFT)
resultpath = StringVar()
resultinput = Entry(resultframe, width=50, textvariable=resultpath)
resultinput.pack(side=LEFT, fill=X, expand=1, padx=2)
resultbtn = Button(resultframe, text=u"浏览", command=lambda: filebrowse(resultpath))
resultbtn.pack(side=LEFT, padx=2)
resultframe.pack(side=TOP, fill=X, expand=1)

queryframe = LabelFrame(root, text=u"查询结果")

inputframe = Frame(queryframe)
inputframe.pack(side=TOP, fill=X, expand=1)
querylabel = Label(inputframe, text=u"查询对象: ")
querylabel.pack(side=LEFT)
query = StringVar()
queryinput = Entry(inputframe, textvariable=query)
queryinput.pack(side=LEFT, fill=X, expand=1)

queryresult = Text(queryframe)
queryresult.pack(side=TOP, fill=X, expand=1)

def scan(valuelist, path, key):
    if not path:
        showwarning(u"错误", u"请输入路径！")
        return

    if not key:
        showwarning(u"错误", u"请输入查询信息！")
        return

    result = list()

    def findrc(sh, i):
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                v = unicode(sh.cell_value(r, c)).replace(" ","")
                if cmp(i,v)==0:
                    return (r,c)

    for f in os.listdir(path):

        p = path + f

        if os.path.isdir(p):
            result.extend(scan(valuelist, p+"/", key))
            continue

        if not (".xls" in f or ".xlsx" in f):
            continue

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

            tmp = dict()
            result.append(tmp)

            for vl in valuelist:
                tmp[vl] = findrc(sh, vl)

            for r in tmp.keys():
                if not tmp[r]:
                    continue

                tmp[r] = unicode( sh.cell_value(row, tmp[r][1]) )

                queryresult.insert(END, tmp[r])
                queryresult.insert(END, ",")

            queryresult.insert(END, "\n")

    return result

def generatexls(name, score):

    wb = Workbook()

    ws = wb.add_sheet("transcripts")

    ####################title begin########################
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

    info = "   姓名：%s    学号：%s    院、系：%s" % (name[0][u"学生姓名"].encode("utf-8"),name[0][u"学号"].encode("utf-8"),name[0][u"院系"].encode("utf-8"))
    ws.write_merge(3,3,0,6, unicode(info), personalinfostyle)
    ws.row(4).height_mismatch = 1
    ws.row(4).height = 18*20
    info = "   专业：%s    方向：%s    导  师：%s" % (name[0][u"专业"].encode("utf-8"),name[0][u"研究方向"].encode("utf-8"),name[0][u"导师姓名"].encode("utf-8"))
    ws.write_merge(4,4,0,6,unicode(info), personalinfostyle)

    ####################title end########################

    ####################table begin########################
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

        if i>len(score):
            ws.write(6+i,1,u"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            ws.write(6+i,2,u"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            ws.write(6+i,3,u"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            ws.write(6+i,4,u"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            ws.write(6+i,5,u"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            ws.write(6+i,6,u"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            continue
        else:
            ws.write(6+i,1,score[i-1][u"课程名称"],easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            ws.write(6+i,2,u"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            ws.write(6+i,3,2,easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))

            timedict = {u"上":4,u"下":5,u"1":4,u"2":5,}
            t = score[i-1][u"学年学期"][len(score[i-1][u"学年学期"])-1]

            sc = score[i-1][u"考试成绩"]


            if timedict[t]==4:
                ws.write(6+i, 4, round(float(sc)) if "." in sc else sc,easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
                ws.write(6+i, 5, u"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            else:
                ws.write(6+i, 5, round(float(sc)) if "." in sc else sc,easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
                ws.write(6+i, 4, u"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))

            ws.write(6+i, 6, u"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))

    ####################table end########################

    ####################tail begin########################
    ws.row(26).height_mismatch = 1
    ws.row(26).height = 18*20

    tailfont = Font() # Create the Font
    tailfont.name = "SimSum"
    tailfont.height = 220 # 11px * 20
    tailalignment = Alignment()
    tailalignment.horz = Alignment.HORZ_RIGHT
    tailalignment.vert = Alignment.VERT_CENTER
    tailstyle = XFStyle() # Create the Style
    tailstyle.font = tailfont # Apply the Font to the Style
    tailstyle.alignment = tailalignment
    ws.row(27).height_mismatch = 1
    ws.row(27).height = 23*20
    ws.write_merge(27,27,0,6,u"研究生教学秘书审核签字：                   ", tailstyle)
    ws.row(28).height_mismatch = 1
    ws.row(28).height = 23*20
    ws.write_merge(28,28,0,6,u"院(系)公章：                   ", tailstyle)
    ws.row(29).height_mismatch = 1
    ws.row(29).height = 42*20
    ws.write_merge(29,29,0,6,u"年    月    日", tailstyle)

    ####################tail end########################

    xlsname = "%s%s.xls" % (resultpath.get(),name[0][u"学号"])

    wb.save(unicode(xlsname))


def queryprocess():
    if not query.get().strip():
        return

    querylist = query.get().strip().split(",")

    nameinfolist = [u"院系", u"学号", u"学生姓名", u"专业", u"研究方向", u"导师姓名"]
    scoreinfolist = [u"学年学期", u"课程名称", u"考试成绩"]

    for q in querylist:
        name = scan(nameinfolist, infopath.get(), q)
        score = scan(scoreinfolist, transcriptpath.get(), q)

        if not len(name):
            continue

        for x in score:
            if not x[u"课程名称"]:
                score.remove(x)
                continue

            if not x[u"课程名称"].strip():
                score.remove(x)
                continue

        generatexls(name, score)

querybtn = Button(inputframe, text=u"查询", command=queryprocess)
querybtn.pack(side=LEFT)

pathframe.pack(side=TOP, fill=X, expand=1, padx=5, pady=2)
queryframe.pack(side=TOP, fill=X, expand=1, padx=5, pady=2)

root.wm_protocol("WM_DELETE_WINDOW", writeConfig)

readConfig()

root.mainloop()
