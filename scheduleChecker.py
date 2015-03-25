from tkinter import *
from tkinter import filedialog
import os
from configparser import *
import sys
import string
import re

import xlsscan
import xlswriter

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
    cf.set("ExcelFilePath", "score", transcriptpath.get())
    cf.set("ExcelFilePath", "name", infopath.get())
    cf.set("ExcelFilePath", "result", resultpath.get())
    # cf.add_section("Field")
    # ft = field.get(0, END)
    # for i in range(len(ft)):
    #     cf.set("Field", "Field_"+str(i), unicode(ft[i]))

    cf.add_section("Query")
    cf.set("Query", "key", query.get())

    cf.write(open("./config.ini", "w"))

    sys.exit()

root = Tk()

root.title("检查")
root.resizable(width=False, height=False)

pathframe = LabelFrame(root, text="路径设置")

def filebrowse(target):
    target.set("") #清空entry里面的内容
    #调用filedialog模块的askdirectory()函数去打开文件夹
    fp = filedialog.askdirectory()
    if fp:
        target.set(fp+"/") #将选择好的路径加入到entry里面

transcriptframe = Frame(pathframe)
transcripttip = Label(transcriptframe,text="成绩目录: ")
transcripttip.pack(side=LEFT)
transcriptpath = StringVar()
transcriptinput = Entry(transcriptframe, width=50, textvariable=transcriptpath)
transcriptinput.pack(side=LEFT, fill=X, expand=1, padx=2)
transcriptbtn = Button(transcriptframe, text="浏览", command=lambda: filebrowse(transcriptpath))
transcriptbtn.pack(side=LEFT, padx=2)
transcriptframe.pack(side=TOP, fill=X, expand=1)

infoframe = Frame(pathframe)
infotip = Label(infoframe,text="名单目录: ")
infotip.pack(side=LEFT)
infopath = StringVar()
infoinput = Entry(infoframe, width=50, textvariable=infopath)
infoinput.pack(side=LEFT, fill=X, expand=1, padx=2)
infobtn = Button(infoframe, text="浏览", command=lambda: filebrowse(infopath))
infobtn.pack(side=LEFT, padx=2)
infoframe.pack(side=TOP, fill=X, expand=1)

resultframe = Frame(pathframe)
resulttip = Label(resultframe,text="结果目录: ")
resulttip.pack(side=LEFT)
resultpath = StringVar()
resultinput = Entry(resultframe, width=50, textvariable=resultpath)
resultinput.pack(side=LEFT, fill=X, expand=1, padx=2)
resultbtn = Button(resultframe, text="浏览", command=lambda: filebrowse(resultpath))
resultbtn.pack(side=LEFT, padx=2)
resultframe.pack(side=TOP, fill=X, expand=1)

queryframe = LabelFrame(root, text="结果")

inputframe = Frame(queryframe)
inputframe.pack(side=TOP, fill=X, expand=1)
querylabel = Label(inputframe, text="查询学号: ")
querylabel.pack(side=LEFT)
query = StringVar()
queryinput = Entry(inputframe, textvariable=query)
queryinput.pack(side=LEFT, fill=X, expand=1)

sl = Scrollbar(queryframe)
sl.pack(side=RIGHT, fill=Y)

queryresult = Text(queryframe, yscrollcommand=sl.set)
queryresult.pack(side=TOP, fill=X, expand=1)

queryresult.config(state=DISABLED)
sl.config(command=queryresult.yview)

def checkpath():
    queryresult.config(state=NORMAL)
    queryresult.insert(END, "===========开始查询%s=============\n" % transcriptpath.get())
    queryresult.config(state=DISABLED)

    xlsscan.check(transcriptpath.get(), queryresult)

    queryresult.config(state=NORMAL)
    queryresult.insert(END, "===========查询完成%s=============\n" % transcriptpath.get())
    queryresult.config(state=DISABLED)

    return

checkbtn = Button(transcriptframe, text="检查", command=checkpath)
checkbtn.pack(side=LEFT, padx=2)

def queryprocess():
    if not query.get().strip():
        return

    querylist = query.get().strip().split(",")

    nameinfolist = ["院系", "学号", "学生姓名", "专业", "研究方向", "导师姓名"]
    scoreinfolist = ["学年学期", "课程名称", "考试成绩"]

    for q in querylist:
        name = xlsscan.scan(nameinfolist, infopath.get(), q, queryresult)
        score = xlsscan.scan(scoreinfolist, transcriptpath.get(), q, queryresult)

        if not len(name):
            continue

        xlswriter.generatexls(name, score, resultpath.get())

querybtn = Button(inputframe, text="查询", command=queryprocess)
querybtn.pack(side=LEFT)

pathframe.pack(side=TOP, fill=X, expand=1, padx=5, pady=2)
queryframe.pack(side=TOP, fill=X, expand=1, padx=5, pady=2)

root.wm_protocol("WM_DELETE_WINDOW", writeConfig)

readConfig()

root.mainloop()
