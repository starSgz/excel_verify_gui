from tkinter import *
import tkinter as tk
import tkinter.ttk
from tkinter import messagebox
from tkinter import filedialog
from v4 import verify_data
import datetime

class App_verify(Frame):
    def __init__(self,master=None,progressbarOne=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.createWidget()

        self.progressbarOne = progressbarOne
        # self.progressbarOne.pack(side=tkinter.TOP)
    def createWidget(self):
        '''创建页面'''
        # self.label01 = Label(self,text='校验')
        # self.label01.grid(row=0, column=0)

        excel_url = StringVar()
        self.entry1 = tk.Entry(self, width='40',text='上传文件',textvariable=excel_url)
        self.entry1.grid(row=1, column=0, ipadx='30', ipady='10', padx='10', pady='20')


        self.btn1 = tk.Button(self, text='上传文件', command=self.upload_file)
        self.btn1.grid(row=1, column=1, ipadx='30', ipady='10', padx='10', pady='20')

        # download_url = StringVar()
        # self.entry2 = tk.Entry(self, width='40',text='存储地址',textvariable=download_url)
        # self.entry2.grid(row=2, column=0, ipadx='30', ipady='10', padx='10', pady='20')
        #
        # self.btn2 = tk.Button(self, text='存储地址', command=self.downoad_file)
        # self.btn2.grid(row=2, column=1, ipadx='30', ipady='10', padx='10', pady='20')

        self.btn3 =tk.Button(self, text='开始校验', command=self.excute_verify)
        self.btn3.grid(row=3, column=0, ipadx='150', ipady='10', padx='10', pady='20')




    #上传
    def upload_file(self):
        selectFile = tk.filedialog.askopenfilename(filetypes=[("数据表", [".xls", ".xlsx",".csv"])])  # askopenfilename 1次上传1个；askopenfilenames1次上传多个
        self.entry1.insert(0, selectFile)

    #下载
    # def downoad_file(self):
    #     selectFile = tk.filedialog.askdirectory()  # askopenfilename 1次上传1个；askopenfilenames1次上传多个
    #     self.entry2.insert(0, selectFile)

    # def login(self):
    #
    #     messagebox.showinfo("正在处理...请勿点击")
    #     print("用户名:" + self.entry1.get())
    #     time.sleep(5)
    #     messagebox.showinfo("已完成...")
    #     # print("密码:" + self.entry2.get())


    def excute_verify(self):
        if self.entry1.get()=='':
            messagebox.showerror(title='错误', message="请添加目录！")
            return ''
        messagebox.showinfo(title='开始执行',message="开始校验...请耐心等待！")

        start_time = datetime.datetime.now()

        #执行校验函数
        excute = verify_data(excel_url=self.entry1.get(),progressbarOne=self.progressbarOne,master=self.master)
        excute.verify_excel_data()

        end_time = datetime.datetime.now()
        take_time = end_time - start_time
        print(take_time)
        messagebox.showinfo(title='执行结束',message="校验完毕！")
        self.progressbarOne['value'] =   0




if __name__ == '__main__':

    def center_window(root, width, height):
        screenwidth = root.winfo_screenwidth()  # 获取显示屏宽度
        screenheight = root.winfo_screenheight()  # 获取显示屏高度
        size = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)  # 设置窗口居中参数
        root.geometry(size)  # 让窗口居中显示

    root = Tk()
    center_window(root, 800, 200)
    progressbarOne = tkinter.ttk.Progressbar(root)
    progressbarOne.pack(side=tkinter.TOP)
    progressbarOne['length']=800
    root.title('校验文件')
    app = App_verify(master=root,progressbarOne=progressbarOne)

    root.mainloop()




