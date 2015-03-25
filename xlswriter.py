# -*- coding: utf-8 -*-

from xlwt import *
from tkinter import *

course = {
    "电磁场与微波技术":{"外语语言基础":("学位课",4),"中国特色社会主义理论与实践研究":("学位课",2),"马克思主义与社会科学方法论":("学位课",1),"自然辩证法":("学位课",1),"应用泛函分析":("学位课",2),"高等电磁理论":("学位课",2),"计算电磁学":("学位课",2),"微波测量":("学位课",2),"微波EDA":("学位课",2),"近代微波技术":("非学位课",2),"现代通信原理":("非学位课",2),"日语":("非学位课",2),"专业外语":("非学位课",2),"光纤通信":("非学位课",2),"电磁兼容":("非学位课",2),"时域有限差分方法":("非学位课",2),"近代天线理论与技术":("非学位课",2),"电波传播":("非学位课",2),"现代微波器件":("非学位课",2),"微波集成电路":("非学位课",2),"电磁成像及应用":("非学位课",2),"电磁场与电磁波":("非学位课",0),"数理方程":("非学位课",0),"微波技术":("非学位课",0)},
    "电路与系统":{"外语语言基础":("学位课",4),"中国特色社会主义理论与实践研究":("学位课",2),"马克思主义与社会科学方法论":("学位课",1),"自然辩证法":("学位课",1),"随机过程":("学位课",2),"VLSI系统设计导论":("学位课",2),"现代信号处理":("学位课",2),"专用集成电路设计":("学位课",2),"硬件系统设计方法":("学位课",2),"现代通信原理":("非学位课",2),"信号检测与估值":("非学位课",2),"通信系统仿真":("非学位课",2),"应用泛函分析":("非学位课",2),"日语":("非学位课",2),"专业外语":("非学位课",2),"信道编码技术":("非学位课",2),"信源编码技术":("非学位课",2),"SOC原理与设计":("非学位课",2),"多采样率信号处理":("非学位课",2),"数字电视广播":("非学位课",2),"嵌入式系统与结构":("非学位课",2),"实时信号处理":("非学位课",2),"智能控制技术":("非学位课",2),"FPGA嵌入式系统设计":("非学位课",2),"软件工程":("非学位课",2),"移动通信":("非学位课",2),"电磁场与电磁波":("非学位课",0),"数理方程":("非学位课",0),"微波技术":("非学位课",0)},
    "通信与信息系统":{"外语语言基础":("学位课",4),"中国特色社会主义理论与实践研究":("学位课",2),"马克思主义与社会科学方法论":("学位课",1),"自然辩证法":("学位课",1),"随机过程":("学位课",2),"现代通信原理":("学位课",2),"现代信号处理":("学位课",2),"通信系统仿真":("学位课",2),"硬件系统设计方法":("学位课",2),"理论声学":("非学位课",2),"应用泛函分析":("非学位课",2),"应用软件设计":("非学位课",2),"数字图像处理":("非学位课",2),"密码学":("非学位课",2),"高速电路EDA设计":("非学位课",2),"日语":("非学位课",2),"信源编码技术":("非学位课",2),"数字声音广播":("非学位课",2),"专业外语":("非学位课",2),"数字电视广播":("非学位课",2),"移动通信":("非学位课",2),"宽带通信网络":("非学位课",2),"现代计算机网络技术":("非学位课",2),"有线电视综合信息网技术":("非学位课",2),"电波传播":("非学位课",2),"广播电视测量与监测":("非学位课",2),"信道编码技术":("非学位课",2),"实时信号处理":("非学位课",2),"现代音频测量":("非学位课",2),"室内环境声学与声场模拟":("非学位课",2),"心理声学":("非学位课",2),"听觉心理实验方法":("非学位课",2),"汉语语音信号处理":("非学位课",2),"嵌入式系统与结构":("非学位课",2),"数字水印与信息伪装技术":("非学位课",2),"数据挖掘与智能信息处理":("非学位课",2),"软件工程":("非学位课",2),"通信电路算法与结构":("非学位课",2),"电磁场与电磁波":("非学位课",0),"数理方程":("非学位课",0),"微波技术":("非学位课",0)},
    "信号与信息处理":{"外语语言基础":("学位课",4),"中国特色社会主义理论与实践研究":("学位课",2),"马克思主义与社会科学方法论":("学位课",1),"自然辩证法":("学位课",1),"随机过程":("学位课",2),"现代通信原理":("学位课",2),"现代信号处理":("学位课",2),"通信系统仿真":("学位课",2),"硬件系统设计方法":("学位课",2),"数字图像处理":("非学位课",2),"应用泛函分析":("非学位课",2),"应用软件设计":("非学位课",2),"剧场工程":("非学位课",2),"计算机控制技术":("非学位课",2),"人工智能与神经网络":("非学位课",2),"高速电路EDA设计":("非学位课",2),"日语":("非学位课",2),"信源编码技术":("非学位课",2),"数字声音广播":("非学位课",2),"专业外语":("非学位课",2),"智能控制技术":("非学位课",2),"数字电视广播":("非学位课",2),"移动通信":("非学位课",2),"宽带通信网络":("非学位课",2),"实时信号处理":("非学位课",2),"广播电视测量与监测":("非学位课",2),"信道编码技术":("非学位课",2),"SOC原理与设计":("非学位课",2),"舞台设备控制技术":("非学位课",2),"网络控制系统":("非学位课",2),"模式识别":("非学位课",2),"嵌入式系统与结构":("非学位课",2),"通信电路算法与结构":("非学位课",2),"软件工程":("非学位课",2),"传感器与电子接口技术":("非学位课",2),"机器视觉":("非学位课",2),"电磁场与电磁波":("非学位课",0),"数理方程":("非学位课",0),"微波技术":("非学位课",0)},
    "电子与通信工程":{"外语语言基础":("学位课",4),"科学社会主义理论与实践":("学位课",1),"自然辩证法":("学位课",2),"随机过程":("学位课",2),"现代通信原理":("学位课",2),"现代信号处理":("学位课",2),"通信系统仿真":("学位课",2),"硬件系统设计方法":("学位课",2),"应用泛函分析":("学位课",2),"高等电磁理论":("学位课",2),"计算电磁学":("学位课",2),"微波测量":("学位课",2),"微波EDA":("学位课",2),"数字图像处理":("非学位课",2),"应用软件设计":("非学位课",2),"近代微波技术":("非学位课",2),"时域有限差分":("非学位课",2),"剧场工程":("非学位课",2),"计算机控制技术":("非学位课",2),"日语":("非学位课",2),"数字电视广播":("非学位课",2),"信源编码技术":("非学位课",2),"数字声音广播":("非学位课",2),"专业外语":("非学位课",2),"移动通信":("非学位课",2),"宽带通信网络":("非学位课",2),"信道编码技术":("非学位课",2),"实时信号处理":("非学位课",2),"数据挖掘与智能信息处理":("非学位课",2),"软件工程":("非学位课",2),"现代计算机网络技术":("非学位课",2),"舞台设备控制技术":("非学位课",2),"网络控制系统":("非学位课",2),"近代天线理论与技术":("非学位课",2),"电波传播":("非学位课",2),"现代微波器件":("非学位课",2),"光纤通信":("非学位课",2),"微波集成电路":("非学位课",2),"电磁成像及应用":("非学位课",2)},
    "集成电路工程":{"外语语言基础":("学位课",4),"科学社会主义理论与实践":("学位课",1),"自然辩证法":("学位课",2),"随机过程":("学位课",2),"VLSI系统设计导论":("学位课",2),"现代信号处理":("学位课",2),"专用集成电路设计":("学位课",2),"硬件系统设计方法":("学位课",2),"信号检测与估值":("非学位课",2),"通信系统仿真":("非学位课",2),"现代通信原理":("非学位课",2),"应用泛函分析":("非学位课",2),"日语":("非学位课",2),"专业外语":("非学位课",2),"信道编码技术":("非学位课",2),"信源编码技术":("非学位课",2),"SOC原理与设计":("非学位课",2),"多采样率信号处理":("非学位课",2),"数字电视广播":("非学位课",2),"嵌入式系统与结构":("非学位课",2),"实时信号处理":("非学位课",2),"FPGA嵌入式系统设计":("非学位课",2)},
    "计算机技术":{"外语语言基础":("学位课",4),"中国特色社会主义理论与实践研究":("学位课",2),"自然辩证法":("学位课",1),"算法设计与分析":("学位课",2),"计算机网络与通信":("学位课",2),"现代软件工程":("学位课",2),"随机过程":("学位课",2),"媒体资产数据化":("非学位课",2),"动画技术":("非学位课",2),"信息安全技术":("非学位课",2),"Mobile computing":("非学位课",2),"Information Security: Technology and Application":("非学位课",2),"IPTV技术":("非学位课",2),"Web信息集成与发现":("非学位课",2),"专业外语资料阅读":("非学位课",2),"数据仓库与数据挖掘":("非学位课",2),"数字媒体存储技术":("非学位课",2),"数字音视频":("非学位课",2),"计算机网络程序设计实践":("非学位课",2),"计算机网路工程实践":("非学位课",2),"浏览器技术与开发":("非学位课",2),"工程实践":("非学位课",6),"操作系统":("非学位课",0),"计算机组成原理":("非学位课",0),"软件工程":("非学位课",0),"计算机网络基础":("非学位课",0)},
    "计算机软件与理论":{"外语语言基础":("学位课",4),"中国特色社会主义理论与实践研究":("学位课",2),"马克思主义与社会科学方法论":("学位课",1),"自然辩证法":("学位课",1),"随机过程":("学位课",2),"算法设计与分析":("学位课",2),"计算机网络与通信":("学位课",2),"计算机网络程序设计实践":("学位课",2),"现代软件工程":("学位课",2),"分布式计算原理与应用":("非学位课",2),"信息安全技术":("非学位课",2),"媒体资产数据化":("非学位课",2),"Mobile computing":("非学位课",2),"Information Security: Technology and Application":("非学位课",2),"现代信号处理":("非学位课",2),"人工智能与神经网络":("非学位课",2),"现代密码学":("非学位课",2),"计算机网络工程实践":("非学位课",2),"专业外语资料阅读":("非学位课",2),"数据仓库与数据挖掘":("非学位课",2),"语义网":("非学位课",2),"组合数学及其应用":("非学位课",2),"面向对象的程序设计工具与实践":("非学位课",2),"人机交互技术":("非学位课",2),"数字音视频":("非学位课",2),"数字媒体存储技术":("非学位课",2),"安全认证技术":("非学位课",2),"云计算":("非学位课",2),"模式识别":("非学位课",2),"浏览器技术与开发":("非学位课",2),"操作系统":("非学位课",0),"计算机组成原理":("非学位课",0),"软件工程":("非学位课",0),"计算机网络基础":("非学位课",0)},
    "计算机应用技术":{"外语语言基础":("学位课",4),"中国特色社会主义理论与实践研究":("学位课",2),"马克思主义与社会科学方法论":("学位课",1),"自然辩证法":("学位课",1),"随机过程":("学位课",2),"算法设计与分析":("学位课",2),"计算机网络与通信":("学位课",2),"计算机网络程序设计实践":("学位课",2),"现代软件工程":("学位课",2),"媒体资产数据化":("非学位课",2),"并行计算原理":("非学位课",2),"信息安全技术":("非学位课",2),"并行程序设计":("非学位课",2),"动画技术":("非学位课",2),"Web信息集成与发现":("非学位课",2),"Mobile computing":("非学位课",2),"Information Security: Technology and Application":("非学位课",2),"高级计算机图形学":("非学位课",2),"计算机网络工程实践":("非学位课",2),"现代密码学":("非学位课",2),"专业外语资料阅读":("非学位课",2),"OPENGL编程":("非学位课",2),"数据仓库与数据挖掘":("非学位课",2),"语义网":("非学位课",2),"人工智能与神经网络":("非学位课",2),"组合数学及其应用":("非学位课",2),"面向对象的程序设计工具与实践":("非学位课",2),"人机交互技术":("非学位课",2),"并行计算应用技术":("非学位课",2),"高速DSP与并行计算技术":("非学位课",2),"电子商务（双语）":("非学位课",2),"浏览器技术与开发":("非学位课",2),"云计算":("非学位课",2),"操作系统":("非学位课",0),"计算机组成原理":("非学位课",0),"软件工程":("非学位课",0),"计算机网络基础":("非学位课",0)},
    "软件工程技术":{"外语语言基础":("学位课",4),"中国特色社会主义理论与实践研究":("学位课",2),"马克思主义与社会科学方法论":("学位课",1),"自然辩证法":("学位课",1),"随机过程":("学位课",2),"算法设计与分析":("学位课",2),"计算机网络与通信":("学位课",2),"计算机网络程序设计实践":("学位课",2),"现代软件工程":("学位课",2),"分布式计算原理与应用":("非学位课",2),"信息安全技术":("非学位课",2),"媒体资产数据化":("非学位课",2),"Mobile computing":("非学位课",2),"Information Security: Technology and Application":("非学位课",2),"现代信号处理":("非学位课",2),"人工智能与神经网络":("非学位课",2),"现代密码学":("非学位课",2),"计算机网络工程实践":("非学位课",2),"专业外语资料阅读":("非学位课",2),"数据仓库与数据挖掘":("非学位课",2),"语义网":("非学位课",2),"组合数学及其应用":("非学位课",2),"面向对象的程序设计工具与实践":("非学位课",2),"人机交互技术":("非学位课",2),"数字音视频":("非学位课",2),"数字媒体存储技术":("非学位课",2),"安全认证技术":("非学位课",2),"云计算":("非学位课",2),"模式识别":("非学位课",2),"浏览器技术与开发":("非学位课",2),"操作系统":("非学位课",0),"计算机组成原理":("非学位课",0),"软件工程":("非学位课",0),"计算机网络基础":("非学位课",0)},
    "物理电子学":{"外语语言基础":("学位课",4),"中国特色社会主义理论与实践研究":("学位课",2),"马克思主义与社会科学方法论":("学位课",1),"自然辩证法":("学位课",1),"随机过程":("学位课",2),"应用泛函分析":("学位课",2),"半导体物理学":("学位课",2),"导波光学":("学位课",2),"光电信息处理":("学位课",2),"光电子学导论":("学位课",2),"VLSI系统设计导论":("非学位课",2),"专用集成电路设计":("非学位课",2),"数字图像处理":("非学位课",2),"光电材料与器件":("非学位课",2),"红外物理与技术":("非学位课",2),"现代光学设计":("非学位课",2),"非线性光学":("非学位课",2),"现代通信原理":("非学位课",2),"现代信号处理":("非学位课",2),"日语":("非学位课",2),"硬件系统设计方法":("非学位课",2),"光通信理论":("非学位课",2),"现代微波器件":("非学位课",2),"微波集成电路":("非学位课",2),"传感器技术":("非学位课",2),"激光光谱学":("非学位课",2),"专业外语":("非学位课",2),"非线性光线光学":("非学位课",2),"光网络":("非学位课",2),"近代天线理论与技术":("非学位课",2),"光电子器件原理及设计":("非学位课",2),"半导体激光技术":("非学位课",2),"有限元方法的数学理论":("非学位课",2),"固体物理":("非学位课",0),"数理方程":("非学位课",0),"光电子技术":("非学位课",0)},
    "计算数学":{"外语语言基础":("学位课",4),"中国特色社会主义理论与实践研究":("学位课",2),"自然辩证法":("学位课",1),"泛函分析":("学位课",4),"微分方程数值解":("学位课",4),"非线性方程组迭代解法":("学位课",4),"最优化理论与方法":("学位课",4),"数字信号处理":("学位课",4),"有限元方法的数学理论":("非学位课",4),"FEPG专业软件的应用":("非学位课",4),"无约束最优化计算方法":("非学位课",4),"图像处理中的数学问题":("非学位课",4),"小波分析理论及其应用":("非学位课",4),"智能计算":("非学位课",4),"区域分解算法":("非学位课",4),"非线性数值分析的理论与方法":("非学位课",4),"图像处理中的快速算法":("非学位课",4),"反问题的计算方法":("非学位课",4),"计算机视觉":("非学位课",4),"矩阵论":("非学位课",4),"算法设计与分析":("非学位课",4),"电磁问题数值计算方法":("非学位课",4)},
    "应用数学":{"外语语言基础":("学位课",4),"中国特色社会主义理论与实践研究":("学位课",2),"自然辩证法":("学位课",1),"泛函分析":("学位课",4),"抽象代数":("学位课",4),"最优化理论与方法":("学位课",4),"拓扑学基础":("学位课",4),"高等计量经济学":("学位课",4),"模糊集理论及其应用":("非学位课",4),"计量经济分析与建模":("非学位课",4),"复半单李代数":("非学位课",4),"图像处理中的数学问题":("非学位课",4),"非线性积分理论及其应用":("非学位课",4),"融合算子理论及其应用":("非学位课",4),"李超代数":("非学位课",4),"环与代数":("非学位课",4),"运筹学概率模型及算法":("非学位课",4),"图论与网络科学":("非学位课",4),"群与代数表示引论":("非学位课",4),"灰色系统":("非学位课",4),"时间序列分析":("非学位课",4),"灰色缓冲算子理论":("非学位课",4),"投入产出分析":("非学位课",4),"多元统计分析":("非学位课",4),"数据分析与Eviews应用":("非学位课",4),"博弈论":("非学位课",4),"Kac-Moody代数":("非学位课",4)}
}

def trueround(number, places=0):
    '''
    trueround(number, places)

    example:

        >>> trueround(2.55, 1) == 2.6
        True

    uses standard functions with no import to give "normal" behavior to
    rounding so that trueround(2.5) == 3, trueround(3.5) == 4,
    trueround(4.5) == 5, etc. Use with caution, however. This still has
    the same problem with floating point math. The return object will
    be type int if places=0 or a float if places=>1.

    number is the floating point number needed rounding

    places is the number of decimal places to round to with '0' as the
        default which will actually return our interger. Otherwise, a
        floating point will be returned to the given decimal place.

    Note:   Use trueround_precision() if true precision with
            floats is needed

    GPL 2.0
    copywrite by Narnie Harshoe <signupnarnie@gmail.com>
    '''
    place = 10**(places)
    rounded = (int(number*place + 0.5if number>=0 else -0.5))/place
    if rounded == int(rounded):
        rounded = int(rounded)
    return rounded

def generatexls(name, score, outpath):

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
    ws.write_merge(0,0,0,6,"中国传媒大学硕士研究生成绩单", titlestyle)

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
    ws.write_merge(1,1,0,6,"（本表一式两份，一份存研究生学位档案，一份存个人档案）", subtitlestyle)

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

    major = name[0]["专业"]
    studentid = name[0]["学号"]

    if major=="计算机技术" or major=="电子与通信工程" or major=="集成电路工程":
        years = 2
    else:
        years = 3

    startyear = int(studentid[0:2])
    ws.write_merge(2,2,0,6,"   学习期限：自 20%d年 9 月至20%d年 6月" % (startyear,startyear+years), personalinfostyle)
    ws.row(3).height_mismatch = 1
    ws.row(3).height = 18*20

    info = "   姓名：%s    学号：%s    院、系：%s" % (name[0]["学生姓名"],studentid,name[0]["院系"])
    ws.write_merge(3,3,0,6, info, personalinfostyle)
    ws.row(4).height_mismatch = 1
    ws.row(4).height = 18*20
    info = "   专业：%s    方向：%s    导  师：%s" % (major,name[0]["研究方向"],name[0]["导师姓名"])
    ws.write_merge(4,4,0,6,info, personalinfostyle)

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
    ws.write_merge(5,6,0,0,"序号",tabletitlestyle)
    ws.write_merge(5,6,1,1,"课程名称",tabletitlestyle)
    ws.write_merge(5,6,2,2,"课程类别",tabletitlestyle)
    ws.write_merge(5,6,3,3,"学分",tabletitlestyle)
    ws.write_merge(5,5,4,6,"开课学期及成绩",tabletitlestyle)
    ws.write(6,4,"第一学期",tabletitlestyle)
    ws.write(6,5,"第二学期",tabletitlestyle)
    ws.write(6,6,"第三学期",tabletitlestyle)

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
            ws.write(6+i,1,"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            ws.write(6+i,2,"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            ws.write(6+i,3,"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            ws.write(6+i,4,"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            ws.write(6+i,5,"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            ws.write(6+i,6,"",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            continue
        else:
            coursename = score[i-1]["课程名称"]
            ws.write(6+i,1,coursename,easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))

            try:
                ws.write(6+i,2,course[major][coursename][0],easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
                ws.write(6+i,3,course[major][coursename][1],easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            except KeyError:
                ws.write(6+i,2,"非学位课",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
                ws.write(6+i,3,2,easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))

            timedict = {"上":4,"下":5,"1":4,"2":5,}
            t = score[i-1]["学年学期"][len(score[i-1]["学年学期"])-1]

            sc = score[i-1]["考试成绩"]
            if "." in str(sc):
                sc = round(float(sc))

            if timedict[t]==4:
                ws.write(6+i, 4, sc, easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
                ws.write(6+i, 5, "",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            else:
                ws.write(6+i, 5, sc, easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
                ws.write(6+i, 4, "",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))

            ws.write(6+i, 6, "",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))

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
    ws.write_merge(27,27,0,6,"研究生教学秘书审核签字：                   ", tailstyle)
    ws.row(28).height_mismatch = 1
    ws.row(28).height = 23*20
    ws.write_merge(28,28,0,6,"院(系)公章：                   ", tailstyle)
    ws.row(29).height_mismatch = 1
    ws.row(29).height = 42*20
    ws.write_merge(29,29,0,6,"年    月    日", tailstyle)

    ####################tail end########################

    xlsname = "%s%s.xls" % (outpath,name[0]["学号"])

    wb.save(xlsname)