#前言

还记得不久之前，写过一篇关于团队管理工具的调研文章[传送门](https://blog.csdn.net/zhaoenweiex/article/details/76407275)，当时调研了一大圈发现对于管理层来说最大的痛点就是没有一个现成的工具支持word周报的导出，传统企业还是比较偏向纸质文件的。再加上想练练python说动手就动手，于是就有了这个项目。
项目定位算是练练手+解决自身痛点。现在终于有了个模样，发布出来，有兴趣的人可以一起来搞一搞。[项目链接](https://github.com/zhaoenweiex/weekreport)

#目的

通过脚本来自动合并团队的周报，节省团队管理者的时间，提升工作效率。

#特性

 1. 支持团队周报的自动下载
 2. 支持自动合并为word
 3. 支持自动团队工作情况合并到记录历史情况的excel中
 4. 将excel和word作为附件发送到

#架构

基于python3编写的命令行脚本程序，基于stmplib，poplib，docx等通用依赖库开发，需要本地磁盘的操作权限。
其实作为一个命令行的小程序来说就没啥架构可言，实际上就是切分了几个模块。
功能模块可以分为：
1.mailOperate模块：基于stmplib和poplib依赖库实现邮件收发，email依赖实现数据解析
2.xlsOperate模块：基于xlrd,xlwt进行excel的操作
3.docxOperate模块：基于docx进行word操作
4.excel合并业务模块：封装业务逻辑，基于xlsOperate模块进行excel的操作
5.word合并业务模块：封装业务逻辑，基于docxOperate模块进行word操作

#准备工作

##邮箱准备

团队管理人员申请一个专门的团队邮箱(163需要并设定专门的客户端密码)。

##模板准备

团队成员下载周报模板，在template文件夹下，填写自己工作内容，再将每个人的周报作为附件发送到指定的团队邮箱中。

##Python环境准备

1.安装python3(https://www.python.org/downloads/windows/)
2.安装pip(下载https://bootstrap.pypa.io/get-pip.py，在下载目录下启动命令行执行python get-pip.py)
3.执行pip install docx
4.执行pip install xlrd
5.执行pip install xlwt
6.执行pip install poplib
7.执行pip install stmplib
8.执行pip install python-docx

##调整配置
程序的配置都是在Config.json中
其中

emailaddress，团队邮箱地址
toAddress，成果发送的邮箱地址
password，邮箱的密码
pop3_serve，团队邮箱的pop3地址
smtp_server，团队邮箱的stmp地址
teamNumber，团队成员数量
orgName，团队名称
根据各自情况进行设定

#使用方法

在需要汇总时，启动脚本，自动(startFullyAuto.py)或手动模式。
1.全自动模式
直接执行命令会自动下载团队所有周报(本周+附件包含周报两个字),并合并为一个word并将数据写入到历史excel中，然后将成果作为附件自动发送到团队邮箱中，以待进一步处理。
2.手动模式
1)启动预处理脚本(startManualPre.py)
2)手动调整内容
3)启动后处理脚本(startManualPost.py），从团队邮箱中找到邮件并进行进一步处理。

#总结

该项目是作为微报项目的一个组成部分的，还有后续敬请期待。

