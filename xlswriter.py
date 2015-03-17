# -*- coding: utf-8 -*-

from xlwt import *
from Tkinter import *

reload(sys)
sys.setdefaultencoding("utf-8")

course = {
    u"电磁场与微波技术":{u"外语语言基础":(u"学位课",4),u"中国特色社会主义理论与实践研究":(u"学位课",2),u"马克思主义与社会科学方法论":(u"学位课",1),u"自然辩证法":(u"学位课",1),u"应用泛函分析":(u"学位课",2),u"高等电磁理论":(u"学位课",2),u"计算电磁学":(u"学位课",2),u"微波测量":(u"学位课",2),u"微波EDA":(u"学位课",2),u"近代微波技术":(u"非学位课",2),u"现代通信原理":(u"非学位课",2),u"日语":(u"非学位课",2),u"专业外语":(u"非学位课",2),u"光纤通信":(u"非学位课",2),u"电磁兼容":(u"非学位课",2),u"时域有限差分方法":(u"非学位课",2),u"近代天线理论与技术":(u"非学位课",2),u"电波传播":(u"非学位课",2),u"现代微波器件":(u"非学位课",2),u"微波集成电路":(u"非学位课",2),u"电磁成像及应用":(u"非学位课",2),u"电磁场与电磁波":(u"非学位课",0),u"数理方程":(u"非学位课",0),u"微波技术":(u"非学位课",0)},
    u"电路与系统":{u"外语语言基础":(u"学位课",4),u"中国特色社会主义理论与实践研究":(u"学位课",2),u"马克思主义与社会科学方法论":(u"学位课",1),u"自然辩证法":(u"学位课",1),u"随机过程":(u"学位课",2),u"VLSI系统设计导论":(u"学位课",2),u"现代信号处理":(u"学位课",2),u"专用集成电路设计":(u"学位课",2),u"硬件系统设计方法":(u"学位课",2),u"现代通信原理":(u"非学位课",2),u"信号检测与估值":(u"非学位课",2),u"通信系统仿真":(u"非学位课",2),u"应用泛函分析":(u"非学位课",2),u"日语":(u"非学位课",2),u"专业外语":(u"非学位课",2),u"信道编码技术":(u"非学位课",2),u"信源编码技术":(u"非学位课",2),u"SOC原理与设计":(u"非学位课",2),u"多采样率信号处理":(u"非学位课",2),u"数字电视广播":(u"非学位课",2),u"嵌入式系统与结构":(u"非学位课",2),u"实时信号处理":(u"非学位课",2),u"智能控制技术":(u"非学位课",2),u"FPGA嵌入式系统设计":(u"非学位课",2),u"软件工程":(u"非学位课",2),u"移动通信":(u"非学位课",2),u"电磁场与电磁波":(u"非学位课",0),u"数理方程":(u"非学位课",0),u"微波技术":(u"非学位课",0)},
    u"通信与信息系统":{u"外语语言基础":(u"学位课",4),u"中国特色社会主义理论与实践研究":(u"学位课",2),u"马克思主义与社会科学方法论":(u"学位课",1),u"自然辩证法":(u"学位课",1),u"随机过程":(u"学位课",2),u"现代通信原理":(u"学位课",2),u"现代信号处理":(u"学位课",2),u"通信系统仿真":(u"学位课",2),u"硬件系统设计方法":(u"学位课",2),u"理论声学":(u"非学位课",2),u"应用泛函分析":(u"非学位课",2),u"应用软件设计":(u"非学位课",2),u"数字图像处理":(u"非学位课",2),u"密码学":(u"非学位课",2),u"高速电路EDA设计":(u"非学位课",2),u"日语":(u"非学位课",2),u"信源编码技术":(u"非学位课",2),u"数字声音广播":(u"非学位课",2),u"专业外语":(u"非学位课",2),u"数字电视广播":(u"非学位课",2),u"移动通信":(u"非学位课",2),u"宽带通信网络":(u"非学位课",2),u"现代计算机网络技术":(u"非学位课",2),u"有线电视综合信息网技术":(u"非学位课",2),u"电波传播":(u"非学位课",2),u"广播电视测量与监测":(u"非学位课",2),u"信道编码技术":(u"非学位课",2),u"实时信号处理":(u"非学位课",2),u"现代音频测量":(u"非学位课",2),u"室内环境声学与声场模拟":(u"非学位课",2),u"心理声学":(u"非学位课",2),u"听觉心理实验方法":(u"非学位课",2),u"汉语语音信号处理":(u"非学位课",2),u"嵌入式系统与结构":(u"非学位课",2),u"数字水印与信息伪装技术":(u"非学位课",2),u"数据挖掘与智能信息处理":(u"非学位课",2),u"软件工程":(u"非学位课",2),u"通信电路算法与结构":(u"非学位课",2),u"电磁场与电磁波":(u"非学位课",0),u"数理方程":(u"非学位课",0),u"微波技术":(u"非学位课",0)},
    u"信号与信息处理":{u"外语语言基础":(u"学位课",4),u"中国特色社会主义理论与实践研究":(u"学位课",2),u"马克思主义与社会科学方法论":(u"学位课",1),u"自然辩证法":(u"学位课",1),u"随机过程":(u"学位课",2),u"现代通信原理":(u"学位课",2),u"现代信号处理":(u"学位课",2),u"通信系统仿真":(u"学位课",2),u"硬件系统设计方法":(u"学位课",2),u"数字图像处理":(u"非学位课",2),u"应用泛函分析":(u"非学位课",2),u"应用软件设计":(u"非学位课",2),u"剧场工程":(u"非学位课",2),u"计算机控制技术":(u"非学位课",2),u"人工智能与神经网络":(u"非学位课",2),u"高速电路EDA设计":(u"非学位课",2),u"日语":(u"非学位课",2),u"信源编码技术":(u"非学位课",2),u"数字声音广播":(u"非学位课",2),u"专业外语":(u"非学位课",2),u"智能控制技术":(u"非学位课",2),u"数字电视广播":(u"非学位课",2),u"移动通信":(u"非学位课",2),u"宽带通信网络":(u"非学位课",2),u"实时信号处理":(u"非学位课",2),u"广播电视测量与监测":(u"非学位课",2),u"信道编码技术":(u"非学位课",2),u"SOC原理与设计":(u"非学位课",2),u"舞台设备控制技术":(u"非学位课",2),u"网络控制系统":(u"非学位课",2),u"模式识别":(u"非学位课",2),u"嵌入式系统与结构":(u"非学位课",2),u"通信电路算法与结构":(u"非学位课",2),u"软件工程":(u"非学位课",2),u"传感器与电子接口技术":(u"非学位课",2),u"机器视觉":(u"非学位课",2),u"电磁场与电磁波":(u"非学位课",0),u"数理方程":(u"非学位课",0),u"微波技术":(u"非学位课",0)},
    u"电子与通信工程":{u"外语语言基础":(u"学位课",4),u"科学社会主义理论与实践":(u"学位课",1),u"自然辩证法":(u"学位课",2),u"随机过程":(u"学位课",2),u"现代通信原理":(u"学位课",2),u"现代信号处理":(u"学位课",2),u"通信系统仿真":(u"学位课",2),u"硬件系统设计方法":(u"学位课",2),u"应用泛函分析":(u"学位课",2),u"高等电磁理论":(u"学位课",2),u"计算电磁学":(u"学位课",2),u"微波测量":(u"学位课",2),u"微波EDA":(u"学位课",2),u"数字图像处理":(u"非学位课",2),u"应用软件设计":(u"非学位课",2),u"近代微波技术":(u"非学位课",2),u"时域有限差分":(u"非学位课",2),u"剧场工程":(u"非学位课",2),u"计算机控制技术":(u"非学位课",2),u"日语":(u"非学位课",2),u"数字电视广播":(u"非学位课",2),u"信源编码技术":(u"非学位课",2),u"数字声音广播":(u"非学位课",2),u"专业外语":(u"非学位课",2),u"移动通信":(u"非学位课",2),u"宽带通信网络":(u"非学位课",2),u"信道编码技术":(u"非学位课",2),u"实时信号处理":(u"非学位课",2),u"数据挖掘与智能信息处理":(u"非学位课",2),u"软件工程":(u"非学位课",2),u"现代计算机网络技术":(u"非学位课",2),u"舞台设备控制技术":(u"非学位课",2),u"网络控制系统":(u"非学位课",2),u"近代天线理论与技术":(u"非学位课",2),u"电波传播":(u"非学位课",2),u"现代微波器件":(u"非学位课",2),u"光纤通信":(u"非学位课",2),u"微波集成电路":(u"非学位课",2),u"电磁成像及应用":(u"非学位课",2)},
    u"集成电路工程":{u"外语语言基础":(u"学位课",4),u"科学社会主义理论与实践":(u"学位课",1),u"自然辩证法":(u"学位课",2),u"随机过程":(u"学位课",2),u"VLSI系统设计导论":(u"学位课",2),u"现代信号处理":(u"学位课",2),u"专用集成电路设计":(u"学位课",2),u"硬件系统设计方法":(u"学位课",2),u"信号检测与估值":(u"非学位课",2),u"通信系统仿真":(u"非学位课",2),u"现代通信原理":(u"非学位课",2),u"应用泛函分析":(u"非学位课",2),u"日语":(u"非学位课",2),u"专业外语":(u"非学位课",2),u"信道编码技术":(u"非学位课",2),u"信源编码技术":(u"非学位课",2),u"SOC原理与设计":(u"非学位课",2),u"多采样率信号处理":(u"非学位课",2),u"数字电视广播":(u"非学位课",2),u"嵌入式系统与结构":(u"非学位课",2),u"实时信号处理":(u"非学位课",2),u"FPGA嵌入式系统设计":(u"非学位课",2)},
    u"计算机技术":{u"外语语言基础":(u"学位课",4),u"中国特色社会主义理论与实践研究":(u"学位课",2),u"自然辩证法":(u"学位课",1),u"算法设计与分析":(u"学位课",2),u"计算机网络与通信":(u"学位课",2),u"现代软件工程":(u"学位课",2),u"随机过程":(u"学位课",2),u"媒体资产数据化":(u"非学位课",2),u"动画技术":(u"非学位课",2),u"信息安全技术":(u"非学位课",2),u"Mobile computing":(u"非学位课",2),u"Information Security: Technology and Application":(u"非学位课",2),u"IPTV技术":(u"非学位课",2),u"Web信息集成与发现":(u"非学位课",2),u"专业外语资料阅读":(u"非学位课",2),u"数据仓库与数据挖掘":(u"非学位课",2),u"数字媒体存储技术":(u"非学位课",2),u"数字音视频":(u"非学位课",2),u"计算机网络程序设计实践":(u"非学位课",2),u"计算机网路工程实践":(u"非学位课",2),u"浏览器技术与开发":(u"非学位课",2),u"工程实践":(u"非学位课",6),u"操作系统":(u"非学位课",0),u"计算机组成原理":(u"非学位课",0),u"软件工程":(u"非学位课",0),u"计算机网络基础":(u"非学位课",0)},
    u"计算机软件与理论":{u"外语语言基础":(u"学位课",4),u"中国特色社会主义理论与实践研究":(u"学位课",2),u"马克思主义与社会科学方法论":(u"学位课",1),u"自然辩证法":(u"学位课",1),u"随机过程":(u"学位课",2),u"算法设计与分析":(u"学位课",2),u"计算机网络与通信":(u"学位课",2),u"计算机网络程序设计实践":(u"学位课",2),u"现代软件工程":(u"学位课",2),u"分布式计算原理与应用":(u"非学位课",2),u"信息安全技术":(u"非学位课",2),u"媒体资产数据化":(u"非学位课",2),u"Mobile computing":(u"非学位课",2),u"Information Security: Technology and Application":(u"非学位课",2),u"现代信号处理":(u"非学位课",2),u"人工智能与神经网络":(u"非学位课",2),u"现代密码学":(u"非学位课",2),u"计算机网络工程实践":(u"非学位课",2),u"专业外语资料阅读":(u"非学位课",2),u"数据仓库与数据挖掘":(u"非学位课",2),u"语义网":(u"非学位课",2),u"组合数学及其应用":(u"非学位课",2),u"面向对象的程序设计工具与实践":(u"非学位课",2),u"人机交互技术":(u"非学位课",2),u"数字音视频":(u"非学位课",2),u"数字媒体存储技术":(u"非学位课",2),u"安全认证技术":(u"非学位课",2),u"云计算":(u"非学位课",2),u"模式识别":(u"非学位课",2),u"浏览器技术与开发":(u"非学位课",2),u"操作系统":(u"非学位课",0),u"计算机组成原理":(u"非学位课",0),u"软件工程":(u"非学位课",0),u"计算机网络基础":(u"非学位课",0)},
    u"计算机应用技术":{u"外语语言基础":(u"学位课",4),u"中国特色社会主义理论与实践研究":(u"学位课",2),u"马克思主义与社会科学方法论":(u"学位课",1),u"自然辩证法":(u"学位课",1),u"随机过程":(u"学位课",2),u"算法设计与分析":(u"学位课",2),u"计算机网络与通信":(u"学位课",2),u"计算机网络程序设计实践":(u"学位课",2),u"现代软件工程":(u"学位课",2),u"媒体资产数据化":(u"非学位课",2),u"并行计算原理":(u"非学位课",2),u"信息安全技术":(u"非学位课",2),u"并行程序设计":(u"非学位课",2),u"动画技术":(u"非学位课",2),u"Web信息集成与发现":(u"非学位课",2),u"Mobile computing":(u"非学位课",2),u"Information Security: Technology and Application":(u"非学位课",2),u"高级计算机图形学":(u"非学位课",2),u"计算机网络工程实践":(u"非学位课",2),u"现代密码学":(u"非学位课",2),u"专业外语资料阅读":(u"非学位课",2),u"OPENGL编程":(u"非学位课",2),u"数据仓库与数据挖掘":(u"非学位课",2),u"语义网":(u"非学位课",2),u"人工智能与神经网络":(u"非学位课",2),u"组合数学及其应用":(u"非学位课",2),u"面向对象的程序设计工具与实践":(u"非学位课",2),u"人机交互技术":(u"非学位课",2),u"并行计算应用技术":(u"非学位课",2),u"高速DSP与并行计算技术":(u"非学位课",2),u"电子商务（双语）":(u"非学位课",2),u"浏览器技术与开发":(u"非学位课",2),u"云计算":(u"非学位课",2),u"操作系统":(u"非学位课",0),u"计算机组成原理":(u"非学位课",0),u"软件工程":(u"非学位课",0),u"计算机网络基础":(u"非学位课",0)},
    u"软件工程技术":{u"外语语言基础":(u"学位课",4),u"中国特色社会主义理论与实践研究":(u"学位课",2),u"马克思主义与社会科学方法论":(u"学位课",1),u"自然辩证法":(u"学位课",1),u"随机过程":(u"学位课",2),u"算法设计与分析":(u"学位课",2),u"计算机网络与通信":(u"学位课",2),u"计算机网络程序设计实践":(u"学位课",2),u"现代软件工程":(u"学位课",2),u"分布式计算原理与应用":(u"非学位课",2),u"信息安全技术":(u"非学位课",2),u"媒体资产数据化":(u"非学位课",2),u"Mobile computing":(u"非学位课",2),u"Information Security: Technology and Application":(u"非学位课",2),u"现代信号处理":(u"非学位课",2),u"人工智能与神经网络":(u"非学位课",2),u"现代密码学":(u"非学位课",2),u"计算机网络工程实践":(u"非学位课",2),u"专业外语资料阅读":(u"非学位课",2),u"数据仓库与数据挖掘":(u"非学位课",2),u"语义网":(u"非学位课",2),u"组合数学及其应用":(u"非学位课",2),u"面向对象的程序设计工具与实践":(u"非学位课",2),u"人机交互技术":(u"非学位课",2),u"数字音视频":(u"非学位课",2),u"数字媒体存储技术":(u"非学位课",2),u"安全认证技术":(u"非学位课",2),u"云计算":(u"非学位课",2),u"模式识别":(u"非学位课",2),u"浏览器技术与开发":(u"非学位课",2),u"操作系统":(u"非学位课",0),u"计算机组成原理":(u"非学位课",0),u"软件工程":(u"非学位课",0),u"计算机网络基础":(u"非学位课",0)},
    u"物理电子学":{u"外语语言基础":(u"学位课",4),u"中国特色社会主义理论与实践研究":(u"学位课",2),u"马克思主义与社会科学方法论":(u"学位课",1),u"自然辩证法":(u"学位课",1),u"随机过程":(u"学位课",2),u"应用泛函分析":(u"学位课",2),u"半导体物理学":(u"学位课",2),u"导波光学":(u"学位课",2),u"光电信息处理":(u"学位课",2),u"光电子学导论":(u"学位课",2),u"VLSI系统设计导论":(u"非学位课",2),u"专用集成电路设计":(u"非学位课",2),u"数字图像处理":(u"非学位课",2),u"光电材料与器件":(u"非学位课",2),u"红外物理与技术":(u"非学位课",2),u"现代光学设计":(u"非学位课",2),u"非线性光学":(u"非学位课",2),u"现代通信原理":(u"非学位课",2),u"现代信号处理":(u"非学位课",2),u"日语":(u"非学位课",2),u"硬件系统设计方法":(u"非学位课",2),u"光通信理论":(u"非学位课",2),u"现代微波器件":(u"非学位课",2),u"微波集成电路":(u"非学位课",2),u"传感器技术":(u"非学位课",2),u"激光光谱学":(u"非学位课",2),u"专业外语":(u"非学位课",2),u"非线性光线光学":(u"非学位课",2),u"光网络":(u"非学位课",2),u"近代天线理论与技术":(u"非学位课",2),u"光电子器件原理及设计":(u"非学位课",2),u"半导体激光技术":(u"非学位课",2),u"有限元方法的数学理论":(u"非学位课",2),u"固体物理":(u"非学位课",0),u"数理方程":(u"非学位课",0),u"光电子技术":(u"非学位课",0)},
    u"计算数学":{u"外语语言基础":(u"学位课",4),u"中国特色社会主义理论与实践研究":(u"学位课",2),u"自然辩证法":(u"学位课",1),u"泛函分析":(u"学位课",4),u"微分方程数值解":(u"学位课",4),u"非线性方程组迭代解法":(u"学位课",4),u"最优化理论与方法":(u"学位课",4),u"数字信号处理":(u"学位课",4),u"有限元方法的数学理论":(u"非学位课",4),u"FEPG专业软件的应用":(u"非学位课",4),u"无约束最优化计算方法":(u"非学位课",4),u"图像处理中的数学问题":(u"非学位课",4),u"小波分析理论及其应用":(u"非学位课",4),u"智能计算":(u"非学位课",4),u"区域分解算法":(u"非学位课",4),u"非线性数值分析的理论与方法":(u"非学位课",4),u"图像处理中的快速算法":(u"非学位课",4),u"反问题的计算方法":(u"非学位课",4),u"计算机视觉":(u"非学位课",4),u"矩阵论":(u"非学位课",4),u"算法设计与分析":(u"非学位课",4),u"电磁问题数值计算方法":(u"非学位课",4)},
    u"应用数学":{u"外语语言基础":(u"学位课",4),u"中国特色社会主义理论与实践研究":(u"学位课",2),u"自然辩证法":(u"学位课",1),u"泛函分析":(u"学位课",4),u"抽象代数":(u"学位课",4),u"最优化理论与方法":(u"学位课",4),u"拓扑学基础":(u"学位课",4),u"高等计量经济学":(u"学位课",4),u"模糊集理论及其应用":(u"非学位课",4),u"计量经济分析与建模":(u"非学位课",4),u"复半单李代数":(u"非学位课",4),u"图像处理中的数学问题":(u"非学位课",4),u"非线性积分理论及其应用":(u"非学位课",4),u"融合算子理论及其应用":(u"非学位课",4),u"李超代数":(u"非学位课",4),u"环与代数":(u"非学位课",4),u"运筹学概率模型及算法":(u"非学位课",4),u"图论与网络科学":(u"非学位课",4),u"群与代数表示引论":(u"非学位课",4),u"灰色系统":(u"非学位课",4),u"时间序列分析":(u"非学位课",4),u"灰色缓冲算子理论":(u"非学位课",4),u"投入产出分析":(u"非学位课",4),u"多元统计分析":(u"非学位课",4),u"数据分析与Eviews应用":(u"非学位课",4),u"博弈论":(u"非学位课",4),u"Kac-Moody代数":(u"非学位课",4)}
}

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

    major = name[0][u"专业"]
    studentid = name[0][u"学号"]

    if major==u"计算机技术" or major==u"电子与通信工程" or major==u"集成电路工程":
        years = 2
    else:
        years = 3

    startyear = int(studentid[0:2])
    ws.write_merge(2,2,0,6,u"   学习期限：自 20%d年 9 月至20%d年 6月" % (startyear,startyear+years), personalinfostyle)
    ws.row(3).height_mismatch = 1
    ws.row(3).height = 18*20

    info = "   姓名：%s    学号：%s    院、系：%s" % (name[0][u"学生姓名"].encode("utf-8"),studentid.encode("utf-8"),name[0][u"院系"].encode("utf-8"))
    ws.write_merge(3,3,0,6, unicode(info), personalinfostyle)
    ws.row(4).height_mismatch = 1
    ws.row(4).height = 18*20
    info = "   专业：%s    方向：%s    导  师：%s" % (major.encode("utf-8"),name[0][u"研究方向"].encode("utf-8"),name[0][u"导师姓名"].encode("utf-8"))
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
            coursename = score[i-1][u"课程名称"]
            ws.write(6+i,1,coursename,easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))

            try:
                ws.write(6+i,2,course[major][coursename][0],easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
                ws.write(6+i,3,course[major][coursename][1],easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
            except KeyError:
                ws.write(6+i,2,u"非学位课",easyxf('font: name SimSum, height 240; borders: left thin, right thin, top thin, bottom thin; alignment: vert center, horz center;'))
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

    xlsname = "%s%s.xls" % (outpath,name[0][u"学号"])

    wb.save(unicode(xlsname))