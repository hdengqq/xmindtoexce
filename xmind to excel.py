# -*- coding:utf-8 -*-
'''
 outher:hdeng 2018
'''

import xlwt,time
from mekk.xmind import XMindDocument
import sys,os
import tkFileDialog
import ttk
import sys
from Tkinter import *
import Tkinter as tk
from tkMessageBox import *
import _winreg


class ChangeImageName(object):
    def __init__(self, master):
        self.resultBook = xlwt.Workbook(encoding='gbk')
        #BookSheet = resultBook.add_sheet(str(u'用例').encode('gbk'),cell_overwrite_ok=True)
        self.BookSheet = self.resultBook.add_sheet(str("test").encode('utf-8'),cell_overwrite_ok=True)

        titleList =["用例目录","用例ID","用例名称","用例类型","优先级","用例描述","前提条件","步骤","期待结果","创建人"]
        colNum = len(titleList)
        '''将titleList写入excel中'''
        for colIndex in range(colNum):
            titleList[colIndex]=unicode(titleList[colIndex],'utf-8')  #字符一直提示失败，在这里强行转换一下才可以写入
            self.BookSheet.write(0,colIndex,titleList[colIndex])
        #make_center(master, 800, 530)
        style = ttk.Style()
        style.configure("BW.TLabel", foreground="red")
        menubar = Menu(root)
        filemenu = Menu(menubar, tearoff=0)
        filemenu.add_command(label="操作说明", command = self.help)
        menubar.add_cascade(label="使用帮助",menu=filemenu)
        root.config(menu=menubar)

        # *************用例名称frame********************
        describe_label_frame = ttk.Labelframe(master, text=u'选择用例名称', width=800)
        describe_label_frame.grid(row=0, pady=10, padx=20, sticky=W + N + S + E)
        self.case_name = tk.IntVar()
        self.ruler1_path_lab = tk.Radiobutton(describe_label_frame, text=u"第一级节点 _ 第二级节点",variable = self.case_name,value = 1,comman =self.export_case_name)
        self.ruler1_path_lab.grid(row=1, column=0, sticky=tk.E,padx=0)
        self.ruler1_path_lab = tk.Radiobutton(describe_label_frame, text=u"第二级节点 _ 第三级节点",variable = self.case_name,value = 2,comman =self.export_case_name)
        self.ruler1_path_lab.grid(row=2, column=0, sticky=tk.E,ipadx=0)
        self.ruler1_path_lab = tk.Radiobutton(describe_label_frame, text=u"第三级节点 _ 第四级节点",variable = self.case_name,value = 3,comman =self.export_case_name)
        self.ruler1_path_lab.grid(row=3, column=0, sticky=tk.E,ipadx=0)
        # *************用例步骤frame********************
        self.case_step = tk.IntVar()
        config_label_frame = ttk.LabelFrame(master, text=u'选择用例步骤', width=800)
        self.folder_path_lab = tk.Radiobutton(config_label_frame, text  =u"    最后的两级节点",variable = self.case_step,value = 1,command = self.export_case_step)
        self.folder_path_lab.grid(row=1, column=0, sticky=tk.E,ipadx=0)
        self.folder_path_text = ttk.Entry(config_label_frame)
        self.folder_path_lab2 = tk.Radiobutton(config_label_frame, text=u"    倒数第二级节点",variable = self.case_step,value = 2,command = self.export_case_step)
        self.folder_path_lab2.grid(row=2, column=0,  sticky=tk.E,ipadx=0)
        self.folder_path_lab3 = tk.Radiobutton(config_label_frame, text=u"第三 _ 第四级节点",variable = self.case_step,value = 3,command = self.export_case_step)
        self.folder_path_lab3.grid(row=3, column=0,  sticky=tk.E,ipadx=0)
        #self.commit_button = ttk.Button(config_label_frame, text=u"提交", )
        config_label_frame.grid(row=1, pady=10, padx=20, sticky=W + N + S + E)
        # *************预期结果frame********************
        self.case_expation = tk.IntVar()
        intro_label_frame = ttk.Labelframe(master, text=u"选择预期结果", width=800)
        self.intro_label =  tk.Radiobutton(intro_label_frame, text=u'最后一级节点',variable = self.case_expation,value = 1,command = self.export_case_expation)
        intro_label_frame.grid(row=2, pady=10, padx=20, sticky=W + N + S + E)
        self.intro_label.grid(row=1, column=1, sticky=tk.E,ipadx=0)
        self.intro_label2 =  tk.Radiobutton(intro_label_frame, text=u'最后两级节点',variable = self.case_expation,value = 2,command = self.export_case_expation)
        self.intro_label2.grid(row=2, column=1, sticky=tk.E,ipadx=0)

        intro_label_frame1 = ttk.Labelframe(master, text=u"常用漏测场景检查", width=800)
        self.check_var2 = tk.IntVar()
        self.Ck_2 = tk.Checkbutton(intro_label_frame1, text = "检查用例中是否漏掉常用场景", variable= self.check_var2)
        #self.Ck_2.select()
        self.Ck_2.grid(row=0, column=0)
        intro_label_frame1.grid(row=3, pady=10, padx=20, sticky=W + N + S + E)

        # self.commit_button = ttk.Button(root, text=u"导入xmind用例",command = self.restart_program)
        # self.commit_button.grid(row=10, column=0, sticky=E, pady=1)
        self.commit_button = ttk.Button(root, text=u"自定义转换",command=self.save_excel)
        self.commit_button.grid(row=10, column=0, sticky=W)
        self.commit_button = ttk.Button(root, text=u"默认转换", command=self.save_to_default)
        self.commit_button.grid(row=10, column=0, sticky=E)

        self.file_opt = options = {}
        options['defaultextension'] = '.xmind'
        options['filetypes'] = [('xmind files', '.xmind')]
        options['initialdir'] = self.get_desktop()
        self.file_dir = self.askopenfilename()[0]

        # #定义用例名，步骤，预期选择项
        # name = self.case_name.get()
        # step = self.case_step.get()
        # expation = self.case_expation.get()

    def help(self):
         showinfo(title = "使用说明",message="【默认转换】：\n将xmind用例第1和2级节点作为用例名称，倒数第2级和第3级作为用例步骤，最后一级作为预期结果】\n\n【自定义转换】：\n选择窗口中自己想要的节点作为用例名、步骤或预期进行转换\n\n【操作说明】：\n1.打开工具会自动调起打开xmind用例窗口\n2.选择xmind用例后根据自己编写的xmind用例风格选择适合的节点作为用例名、步骤和结果\n3.常用场景检查选择后，转换成功会弹出提示文档推荐用例中要关注的检查点\n"
                                         "4.点击开始转换会自动转换成功并打开用例（当常用场景检查选择后需要关闭提示文档才会自动打开用例）\n5.用例名称相同会自动增加脚标保证用例名称唯一性\n6.用例文档和场景检查文档会保存在与工具相同目录下\n"
                                         "7.支持最大xmind用例节点为6级（含根节点）\n\n -------------------------研发机动组 ")


    def askopenfilename(self):
        dirnamelist = []
        dirName =tkFileDialog.askopenfilename(**self.file_opt) #打开目录找到文件夹
        dirnamelist.append(dirName)
        return dirnamelist

    def get_desktop(self):#打开桌面

         key = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER,\
                          r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders',)
         return _winreg.QueryValueEx(key, "Desktop")[0]


    # def restart_program(self): #重启程序
    #
    #     python = sys.executable
    #     os.execl(python, python, * sys.argv)
    #     self.askopenfilename()


    # 获取到xmind中所有的用例节点，放入一个二维列表中
    def get_total_list(self):

        xmind = XMindDocument.open(self.file_dir)
        sheet = xmind.get_first_sheet()
        root = sheet.get_root_topic()
        root_title = root.get_title()
        fist_node = root.get_subtopics()
        # print 'fist_node',fist_node

        fist_node_list = []
        sencond_node_list =[]
        third_node_list = []
        total_third_node_list=[]
        four_node_list = []
        total_four_node_list = []
        five_node_list = []
        total_five_node_list = []
        six_node_list =[]
        total_six_node_list =[]
        for sencond_node in  fist_node:
            sencond_node_list.append(sencond_node.get_title())
            fist_node_list.append(sencond_node.get_title())
            for third_node in sencond_node.get_subtopics():
                third_node_list.append(third_node.get_title())
                for four_node in third_node.get_subtopics():
                    four_node_list.append(four_node.get_title())
                    for five_node in four_node.get_subtopics():
                        five_node_list.append(five_node.get_title())
                        for six_node in five_node.get_subtopics():
                            six_node_list.append(six_node.get_title())
                        total_six_node_list.append(six_node_list)
                        six_node_list=[]
                    total_five_node_list.append(five_node_list)
                    five_node_list=[]
                total_four_node_list.append(four_node_list)
                four_node_list = []
            total_third_node_list.append(third_node_list)
            third_node_list = []
        return fist_node_list,total_third_node_list,total_four_node_list,total_five_node_list,total_six_node_list,root_title



    # 将用例例表中的空列表增加空值
    def rebase_list(self,seq):
        for i in range(len(seq)):
            if len(seq[i]) == 0 or seq[i] == None:
                seq[i].append("")
        return  seq

    #将二维列表转成一维列表，用于过渡找到空值列表在用例例中的位置
    def flatten(self,seq):
        node_list = []
        for i in range(len(seq)):
            if len(seq[i]) == 0 or seq[i] == None:
                node_list.append("")
            for j in range(len(seq[i])):
                node_list.append(seq[i][j])
        return node_list

    # 将所有的用例导入一个总的列表中
    def case_list_covert(self):

        fist_node_list = self.get_total_list()[0]
        total_third_node_list = self.get_total_list()[1]
        total_four_node_list = self.get_total_list()[2]
        total_five_node_list = self.get_total_list()[3]
        total_six_node_list = self.get_total_list()[4]

        #各节点列表空值处理，如果上一级节点用例列表中有空值，下一级列表中相对应的位置就增加一个空值
        two_node = self.flatten(total_third_node_list)
        for i in range(len(two_node)):
            if two_node[i] == "" or two_node[i] == None:
                total_four_node_list.insert(i,[""])
        three_node = self.flatten(total_four_node_list)
        for j in range(len(three_node)):
            if three_node[j] == "" or three_node[j] == None:
                total_five_node_list.insert(j,[""])
        four_node = self.flatten(total_five_node_list)
        for k in range(len(four_node)):
            if four_node[k] == "" or four_node[k] == None:
                total_six_node_list.insert(k,[""])

        list0 = self.rebase_list(fist_node_list)
        list1 = self.rebase_list(total_third_node_list)
        list2 = self.rebase_list(total_four_node_list)
        list3 = self.rebase_list(total_five_node_list)
        list4 = self.rebase_list(total_six_node_list)

        # 将列表转成用例
        case_list = []
        total_list = []
        k = 0
        b = 0
        c = 0
        for m in range(len(list0)):
            for i in range(len(list1[m])):
                for j in range(len(list2[k])):
                    for n in range(len(list3[b])):
                        for l in range(len(list4[c])):
                            case_list.append(list0[m])
                            case_list.append(list1[m][i])
                            case_list.append(list2[k][j])
                            case_list.append(list3[b][n])
                            case_list.append(list4[c][l])
                            total_list.append(case_list)
                            case_list=[]
                        c = c+1
                    b = b+1
                k = k+1
        return total_list


    #选择用例名称,一级节点为根节点的下一级节点

    def export_case_name(self,param=None):
        total_list = self.case_list_covert()

        #将第一级节点和第二级节点打印出来
        if param == None:
            option = self.case_name.get()
        else:
            option = param
        if option == 1:
            j = 1
            n = 0
            for i in range(len(total_list)):
                if i < len(total_list)-1:
                    if total_list[i][1] == total_list[j][1]:
                        name = total_list[i][0] + "_" + total_list[i][1] + "_" + str(n)
                        self.BookSheet.write(i+1,2,name)
                        n = n+1
                    else:
                        name = total_list[i][0] + "_" + total_list[i][1]
                        self.BookSheet.write(i+1,2,name)
                        n = 0
                    j = j+1
                else:
                    name = total_list[i][0] + "_" + total_list[i][1]
                    self.BookSheet.write(i+1,2,name)
        #将第二级节点和第三级节点打印出来
        elif option == 2:
            j = 1
            n = 0
            for i in range(len(total_list)):
                if i < len(total_list)-1:
                    if total_list[i][2] != "":
                        if total_list[i][2] == total_list[j][2]:
                            name = total_list[i][1] + "_" + total_list[i][2] + "_" + str(n)
                            self.BookSheet.write(i+1,2,name)
                            n = n + 1
                        else :
                            name = total_list[i][1] + "_" + total_list[i][2]
                            self.BookSheet.write(i+1,2,name)
                            n = 0
                        j = j +1
                    else:
                        name = total_list[i][1]
                        self.BookSheet.write(i+1,2,name)
        elif option == 3:
        #将第三级节点和第四级节点打印出来
            for i in range(len(total_list)):
                if total_list[i][3] != "":
                    name = total_list[i][2] + "_" + total_list[i][3]
                    self.BookSheet.write(i+1,2,name)
                else:
                    name = total_list[i][2]
                    self.BookSheet.write(i+1,2,name)
        else:
             return False


    #选择用例步骤
    def export_case_step(self,param = None):
        total_list = self.case_list_covert()
        BookSheet = self.BookSheet
        if param == None:
            option = self.case_step.get()
        else:
            option = param
        #option = self.case_step.get()
        # 将用例节点的最后一级和最后第二级打印出来
        if option == 1:
            max_node = len(total_list[0])
            for i in range(len(total_list)):
                for j in range(max_node):
                    if total_list[i][max_node-1-j] !="":
                        content = total_list[i][max_node-1-j] + "_" + total_list[i][max_node-2-j]
                        BookSheet.write(i+1,7,content)
                        break
        elif option == 2:
            # 将倒数第二级的用例节点打印出来
            max_node = len(total_list[0])
            for i in range(len(total_list)):
                  for j in range(max_node):
                    if total_list[i][max_node-1-j] !="":
                        BookSheet.write(i+1,7,total_list[i][max_node-2-j])
                        break
                    else:
                        BookSheet.write(i+1,7,total_list[i][max_node-3-j])
        elif option == 3:
            #将第三级节点和第四级节点打印出来
            for i in range(len(total_list)):
                if total_list[i][3] != "":
                    name = total_list[i][2] + "_" + total_list[i][3]
                    BookSheet.write(i+1,7,name)
                else:
                    name = total_list[i][2]
                    BookSheet.write(i+1,7,name)
        elif option == 4:
            # 2018新加将倒数第二级的用例节点和倒数第三级节点打印出来
            max_node = len(total_list[0])
            for i in range(len(total_list)):
                  for j in range(max_node):
                    if total_list[i][max_node-1-j] !="":
                        name = total_list[i][max_node-3-j]+"_"+total_list[i][max_node-2-j]
                        BookSheet.write(i+1,7,name)
                        break
                    else:
                        name = total_list[i][max_node-4-j]+"_"+total_list[i][max_node-3-j]
                        BookSheet.write(i+1,7,name)
        else:
            return False


    # 将预期结果导出
    def export_case_expation(self,param = None):

        # 将用例节点的最后一级打印出来
        total_list = self.case_list_covert()
        BookSheet = self.BookSheet
        if param == None:
            option = self.case_expation.get()
        else:
            option = param
        #option = self.case_expation.get()
        if option == 1:
            max_node = len(total_list[0])
            for i in range(len(total_list)):
                for j in range(max_node):
                    if total_list[i][max_node-1-j] !="":
                        BookSheet.write(i+1,8,total_list[i][max_node-1-j])
                        break
        elif option == 2:
            max_node = len(total_list[0])
            for i in range(len(total_list)):
                  for j in range(max_node):
                    if total_list[i][max_node-2-j] !="":
                        BookSheet.write(i+1,8,total_list[i][max_node-1-j])
                        break
        else:
              return False

    def xminddata(self):
        cases = {
                 #探索性测试异常场景
                 '请确认是否覆盖网络切换检查的场景，如从wif切换到3G':['网络切换','WIFI','wifi' '切换网络','切换','切换成4G','XG','从无网到有网','恢复网络','断网恢复'],
                 '请确认是否覆盖弱网下检查的用例场景':['弱网','网络较弱'],
                 '请确认是否覆盖到2G/3G/4G等网络的场景':['4G','3G','2G'],
                 '请确认是否覆盖视频电话/短信/闹铃等冲突的场景，如操作过程中有来电话，播放过程中退出操作等':['电话','视频','短信','闹铃','播放过程中'],
                 '请确认是否有检查多点触控场景，如同时点击两个控件，同时滑动':['同时点','多点','同时滑','同时拖'],
                 '请确认是否有频繁压力操作场景，如反复频繁点击，或反复进入某个界面':['反复','频繁','多次'],
                 '请确认是否有频覆盖安装操作场景，如覆盖旧版本':['覆盖'],
                 '请确认是否有调节字体大小的场景，如调最大字体':['字体','放大','缩小'],
                 '请确认是否有结合指纹锁，密码锁的场景，如指纹锁，手势密码等':['手势','指纹'],
                 '请确认是否有被踢下线的场景，如操作过程中被踢下线':['被踢','下线'],
                 '请确认是否有第三方跳转相关的操作，如从第三方跳转到待测界面':['跳转','唤起','返回QQ'],
                 #漏测场景
                 '请确认是否有检查系统等兼容性的的用例场景，如ios11/10':['ios9','ios8','ios10','IOS'],
                 '请确认是否有检查系统等兼容性的的用例场景，建议将iphone4/4S等小屏机加入用例检查':['ip4','iphone4','4s','IP4','IPHONE','4S'],
                 '请确认是否有检查退后台和杀进程的一些场景，如删除数据的一些操作':['退后台','杀进程','杀掉进程','重启'],
                 '请确认是否有检查横屏、旋转的场景':['横屏','旋转'],

                 }

        xmindFile = open(self.file_dir)
        File = xmindFile.read()
        remond_list = []
        txt_dir = os.getcwd() + "\\" + "log"  + str(time.time())+ ".txt"
        for i in range(0,len(cases.values())):
            #检查异常场景用例字典库中的所有value值(value为列表list)
            for n in range(0,len(cases.values()[i])):#取用例字典库中的value值list中的每个数据
                if cases.values()[i][n] in File:# 将字典库中的value值list中的每个数据与导入的用例进行对比
                     break   # 当匹配到到键值对中有的关键字时直接跳过
                else:
                    show = cases.keys()[i] # 当没有匹配到键值对中有的关键字时，以“键”来提示测试人员修正
                    remond_list.append(show)
        new_remond = list(set(remond_list))

        file_object = open(txt_dir, 'w+')

        for i in new_remond:
            file_object.write(i + "\n")
        file_object.close()
        os.popen(txt_dir + " " + "\n")


    def save_excel(self):
        case_file_name = "test2018" + str(time.time()) + ".xls"
        if self.file_dir == "":
            showinfo(title = "扎心了老铁",message="请重新打开工具，导入Xmind文件")
        elif self.export_case_name() == False:
           showinfo(title = "扎心了老铁",message="请选择用例名称")
        elif self.export_case_step() == False:
           showinfo(title = "扎心了老铁",message="请选择用例步骤")
        elif self.export_case_expation() == False:
           showinfo(title = "扎心了老铁",message="请选择预期结果")
        else:
            if self.check_var2.get() == 1:
                self.xminddata()

            self.export_root_dir()
            self.resultBook.save(case_file_name)
            filedir = os.getcwd() + "\\" + case_file_name
            os.popen(filedir)

    def save_to_default(self):
        case_file_name = "test2018" + str(time.time()) + ".xls"
        try:
            if self.file_dir == "":
                showinfo(title="扎心了老铁", message="请关闭后重新打开工具，导入Xmind文件")

            self.export_case_name(param=1)
            self.export_case_step(param=4)
            self.export_case_expation(param=1)
            if self.check_var2.get() == 1:
                self.xminddata()
            self.export_root_dir()
            self.resultBook.save(case_file_name)
            filedir = os.getcwd() + "\\" + case_file_name
            os.popen(filedir)

        except Exception as e:
            print e

    def export_root_dir(self):
        total_list = self.case_list_covert()
        name = self.get_total_list()[5]
        BookSheet = self.BookSheet
        for i in range(len(total_list)):
            BookSheet.write(i+1,0,name)


if __name__ == '__main__':

    root = tk.Tk()
    root.title(u"Xmind to Excel")
    root.resizable(0, 0)
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    width = 250
    height = 450
    size = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    root.geometry(size)
    ChangeImageName(root)
    #LoginPage(root)
    root.mainloop()
