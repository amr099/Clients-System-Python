from calendar import month
from tkinter import *
from tkinter import Entry, Label, Tk, ttk, messagebox
import tkinter as tk
from tkinter import font
from babel.core import Locale
from babel.dates import format_date    
from openpyxl import load_workbook
from datetime import date, datetime
import re
from openpyxl.workbook import workbook
from openpyxl.styles import NamedStyle, Font, Border, Side
from tkcalendar import Calendar,DateEntry


file = 'xl.xlsx'
try:
  load_workbook(file)
except:
  messagebox.showerror('File Not Found!','غير موجود xl.xlsx  ملف الاكسيل')

wb = load_workbook(file)

services = ['اتعاب لجنه داخلية','اضافة سياره'
        ,'اعداد و مراجعة ميزانيه','اقرار ضرائب عامه'
        ,'اقرار ضرائب قيمه مضافه', 'اقرار ضرائب مرتبات', 'بطاقه ضريبيه'
        ,'تجديد اشتراك البوابه الالكترونيه', 'تحت الحساب','تسوية ملف ضريبى', 'تعديل النشاط',
        'جواب مرور', 'حفظ الملف بالضرائب'
        ,'رسوم استخراج مستوردين', 'رسوم البوابه الالكترونيه','رسوم تجديد مستوردين', 'سجل تجارى',
        'سداد ضرائب عامه', 'سداد ضرائب قيمه مضافه','شطب سجل تجارى', 'شهاده بالموقف الضريبى'
        ,'شهادة دخل 1', 'شهادة دخل 2', 'عمل موقع الكترونى', 'غرفه تجاريه',
        'فحص ضرائب عامه', 'فحص ضريبة قيمه مضافه','لجنة طعن ضرائب عامه', 'مركز مالى',
        'ميزانيه عموميه', 'نماذج 41 ض',]

class Revenue(tk.Toplevel):
      def __init__(self, parent):
        super().__init__(parent)

        self.geometry('1600x900')
        self.title('الايرادات')
        self['bg']='#E5E8C7'


        Label(self, text='من',
        bg = '#43516C', fg = 'white', font = 'fantasy 30 bold',
        borderwidth=5, relief="ridge", padx=20, pady=10).place(relx=.7, rely=.2,anchor= CENTER, width=400)

        startdate = DateEntry(self,width=30,bg="darkblue",fg="white", font = ('fantasy', 25, 'bold'),
        year=datetime.now().year,month=datetime.now().month,day=datetime.now().day
        ,locale='en_US', date_pattern='dd/MM/yyyy')
        startdate.place(relx=.7, rely=.35,anchor= CENTER, width=200, height=50)
        

        Label(self, text='الى',
        bg = '#43516C', fg = 'white', font = 'fantasy 30 bold',
        borderwidth=5, relief="ridge", padx=20, pady=10).place(relx=.3, rely=.2,anchor= CENTER, width=400)

        enddate = DateEntry(self,width=30,bg="darkblue",fg="white", font = ('fantasy', 25, 'bold'),
        year=datetime.now().year,month=datetime.now().month,day=datetime.now().day
        ,locale='en_US', date_pattern='dd/MM/yyyy')
        enddate.place(relx=.3, rely=.35,anchor= CENTER, width=200, height=50)
        

        def revenue():      
                total_revenue = 0
                for sheet in wb.worksheets:
                        ws = sheet
                        for i in range(4, ws.max_row+1):
                                celldate = date(ws.cell(row=i, column=8).value,ws.cell(row=i, column=9).value,1)
                                if celldate >= startdate.get_date() and celldate <= enddate.get_date():
                                        total_revenue += ws.cell(row=i, column=4).value

                Label(self, text=str(total_revenue),
                bg = '#F1B18A', fg = 'white', font = 'fantasy 20 bold',
                borderwidth=5, relief="ridge", padx=20, pady=10).place(relx=.5, rely=.8,anchor= CENTER, width=400)
                                

        Button(self, height = 1, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 20 bold', bd = '8px solid #DBA531', 
                text='بحث',
                command=revenue).place(relx=.5, rely=.65,anchor= CENTER)

        Button(self, height = 1, width = 10, bg = 'grey', fg = 'white',
        activebackground='#43516C', font = 'fantasy 15 bold', bd = '8px solid #DBA531', 
                text='اغلاق',
                command=self.destroy).place(relx=.1, rely=.9,anchor= CENTER)        
        


class Payment(tk.Toplevel):
      def __init__(self, parent):
        super().__init__(parent)

        self.geometry('1600x900')
        self.title('الدفوعات')
        self['bg']='#E5E8C7'


        entries = []
        for sheet in wb.worksheets:
                if sheet.cell(row=1,column=1).value != None :
                        entries.append(f'{str(sheet.cell(row=1,column=1).value)} ({str(sheet.cell(row=1,column=2).value)})')                
        if len(entries) < 1:
                entries.append('لا يوجد عملاء')   

        Label(self, text='الاسم',
        bg = '#43516C', fg = 'white', font = 'fantasy 20 bold', borderwidth=5,
        relief="ridge", padx=20, pady=10).place(relx=.8, rely=.1,anchor= CENTER, width=200)

        payName = StringVar()
        menu1 = ttk.Combobox(self, font=('Calibri', 20,'bold'),values = entries, textvariable = payName )
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu1.place(relx=.5, rely=.1,anchor= CENTER, width=600)
        menu1['state'] = 'readonly'


        Label(self, text='المدفوع',
        bg = '#43516C', fg = 'white', font = 'fantasy 20 bold', borderwidth=5, relief="ridge",
        padx=20, pady=10).place(relx=.8, rely=.25,anchor= CENTER, width=200)

        payAmount = IntVar()
        entry1 = ttk.Entry(self, textvariable = payAmount, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.5, rely=.25,anchor= CENTER, width=600, height=50)
        

        Label(self, text='ملاحظات',
        bg = '#43516C', fg = 'white', font = 'fantasy 20 bold', borderwidth=5, relief="ridge",
        padx=20, pady=10).place(relx=.8, rely=.4,anchor= CENTER, width=200)

        payComment = StringVar()
        entry1 = ttk.Entry(self, textvariable=payComment, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.5, rely=.4,anchor= CENTER, width=600, height=50)

        Label(self, text='التاريخ',
        bg = '#43516C', fg = 'white', font = 'fantasy 20 bold', borderwidth=5, relief="ridge",
        padx=20, pady=10).place(relx=.8, rely=.55,anchor= CENTER, width=200)

        date = DateEntry(self,width=30,bg="darkblue",fg="white",year=datetime.now().year,
        month=datetime.now().month,day=datetime.now().day
        ,locale='en_US', date_pattern='dd/MM/yyyy',font = ('fantasy', 25, 'bold'))
        date.place(relx=.625, rely=.55,anchor= CENTER, width=200, height=50)
        
        
        
        def payment():
                try:
                        payAmount.get()
                except:
                        messagebox.showerror('Invalid!','من فضلك قم بادخال الخانات بطريقه صحيحه ') 
                        self.destroy()
                        return      
                if payAmount.get() == '':
                        messagebox.showerror('Invalid!','من فضلك قم بادخال خانة المدفوع بطريقه صحيحه ')
                        self.destroy()
                        return
                if payName.get() == '':
                        messagebox.showerror('Invalid!','من فضلك قم بادخال خانة الاسم بطريقه صحيحه ')
                        self.destroy()
                        return        

                flag = False
                finalAmount = 0
                for sheet in wb.worksheets:
                        ws = sheet
                        if f'{str(sheet.cell(row=1,column=1).value)} ({str(sheet.cell(row=1,column=2).value)})' == payName.get():
                                for i in range (1, ws.max_row+1):
                                        if ws.cell(column=7,row=i).value == ws.cell(column=3,row=1).value:
                                                finalAmount = ws.cell(column=5,row=i).value
                                ws['C1'] = int(ws.cell(column=3, row=1).value) + 1
                                ws.append([date.get_date().strftime("%d/%m/%Y"),
                                '-', '-',
                                payAmount.get(),
                                finalAmount - payAmount.get(),
                                payComment.get(),
                                ws.cell(column=3, row=1).value,
                                date.get_date().year, date.get_date().month, date.get_date().day])
                                wb.save(file)
                                flag = True
                                self.destroy()
                                messagebox.showinfo('Done','تم التسجيل بنجاح ')
                                break
                if not flag:
                        messagebox.showerror('Not Exists!','الاسم غير موجود') 
                        self.destroy()
                


        Button(self, height = 1, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 20 bold', bd = '8px solid #DBA531', 
                text='تسجيل',
                command=payment).place(relx=.5, rely=.75,anchor= CENTER)


        Button(self, height = 1, width = 10, bg = 'grey', fg = 'white',
        activebackground='#43516C', font = 'fantasy 15 bold', bd = '8px solid #DBA531', 
                text='اغلاق',
                command=self.destroy).place(relx=.1, rely=.9,anchor= CENTER)        
        
# # --------------------------------------------------------------------------------------------------------------------------------------------                

class Search(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self.geometry('1600x900')
        self.title('البحث عن عميل')
        self['bg']='#E5E8C7'

        entries = []
        for sheet in wb.worksheets:
                if sheet.cell(row=1,column=1).value != None :
                        entries.append(f'{str(sheet.cell(row=1,column=1).value)} ({str(sheet.cell(row=1,column=2).value)})')

        if len(entries) < 1:
                entries.append('لا يوجد عملاء')  

        searchName = StringVar()
        menu = ttk.Combobox(self, textvariable= searchName, font=('Calibri', 20,'bold'),values = entries )
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu.place(relx=.5, rely=.25,anchor= CENTER, width=600)
        menu['state'] = 'readonly'

        def search():
                treev = ttk.Treeview(self, selectmode ='browse', style="mystyle.Treeview", height=900)
                treev.pack()
                
                style = ttk.Style()
                
                style.configure("mystyle.Treeview", background = '#E5E8C7' ,rowheight=40,
                highlightthickness=0, bd=0, font=('Calibri', 14)) 
                style.configure("mystyle.Treeview.Heading", font=('Calibri', 20,'bold')) 
                style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})]) 

                verscrlbar = ttk.Scrollbar(self,
                                orient ="vertical",
                                command = treev.yview)

                verscrlbar.pack(side ='left', fill ='x')   
                treev.configure(xscrollcommand = verscrlbar.set)

                treev["columns"] = ("1", "2", "3", "4", "5", "6")

                treev['show'] = 'headings'
        
                treev.column("1", width = 100,anchor ='e')
                treev.column("2", width = 220,anchor ='e')
                treev.column("3", width = 100,anchor ='e')
                treev.column("4", width = 100,anchor ='e')
                treev.column("5", width = 100,anchor ='e')
                treev.column("6", width = 950,anchor ='e')
                
                treev.heading("1", text ="التاريخ")
                treev.heading("2", text ="الخدمه")
                treev.heading("3", text ="التكلفه")
                treev.heading("4", text ="المدفوع")
                treev.heading("5", text ="الرصيد")
                treev.heading("6", text ="الملاحظات")

                for sheet in wb.worksheets:
                        ws = sheet
                        if f'{str(ws.cell(row=1,column=1).value)} ({str(ws.cell(row=1,column=2).value)})' == searchName.get():
                                for i in range (4, ws.max_row+1):
                                        if i%2 == 0:
                                                treev.insert("", 'end', text ="L7",
                                                                values =(
                                                                '-' if (ws.cell(row=i,column=1).value) == None else str(ws.cell(row=i,column=1).value), 
                                                                '-' if (ws.cell(row=i,column=2).value) == None else str(ws.cell(row=i,column=2).value),
                                                                '-' if (ws.cell(row=i,column=3).value) == None else str(ws.cell(row=i,column=3).value),
                                                                '-' if (ws.cell(row=i,column=4).value) == None else str(ws.cell(row=i,column=4).value),
                                                                '-' if (ws.cell(row=i,column=5).value) == None else str(ws.cell(row=i,column=5).value),
                                                                '-' if (ws.cell(row=i,column=6).value) == None else str(ws.cell(row=i,column=6).value)), tags = ('even',))
                                        else:
                                                treev.insert("", 'end', text ="L7",
                                                                values =(
                                                                '-' if (ws.cell(row=i,column=1).value) == None else str(ws.cell(row=i,column=1).value), 
                                                                '-' if (ws.cell(row=i,column=2).value) == None else str(ws.cell(row=i,column=2).value),
                                                                '-' if (ws.cell(row=i,column=3).value) == None else str(ws.cell(row=i,column=3).value),
                                                                '-' if (ws.cell(row=i,column=4).value) == None else str(ws.cell(row=i,column=4).value),
                                                                '-' if (ws.cell(row=i,column=5).value) == None else str(ws.cell(row=i,column=5).value),
                                                                '-' if (ws.cell(row=i,column=6).value) == None else str(ws.cell(row=i,column=6).value)), tags = ('odd',))                        
                                        debt = str(ws.cell(row=i,column=5).value)
                        
                        treev.tag_configure('odd', background='#e1dddd', font=('Calibri', 14)) 
                        treev.tag_configure('even', background= '#f5f3f3', font=('Calibri', 14)) 

                # Label(self, text= debt,
                # bg = '#43516C', fg = 'white', font = 'fantasy 20 bold', borderwidth=5, relief="ridge",
                # padx=20, pady=10 ).place(relx=.1, rely=.9,anchor= CENTER, width=200)
                self.title( searchName.get())

        
        Button(self, height = 1, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 20 bold', bd = '8px solid #DBA531', 
                text='ابحث',
                command= search).place(relx=.5, rely=.5,anchor= CENTER)
        
        Button(self, height = 1, width = 10, bg = 'grey', fg = 'white',
        activebackground='#43516C', font = 'fantasy 15 bold', bd = '8px solid #DBA531', 
                text='اغلاق',
                command=self.destroy).place(relx=.5, rely=.8,anchor= CENTER)

        # --------------------------------------------------------------------------------------------------------------------------------------------

class Insert(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self.geometry('1600x900')
        self.title('اضافة عميل')
        self['bg']='#E5E8C7'

        Label(self, text='الاسم',
        bg = '#43516C', fg = 'white', font = 'fantasy 20 bold', borderwidth=5, relief="ridge", 
        padx=20, pady=10 ).place(relx=.8, rely=.1,anchor= CENTER, width=200)

        insertName = tk.StringVar()
        entry1 = ttk.Entry(self, textvariable=insertName, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.5, rely=.1,anchor= CENTER, width=600)


        Label(self, text='كود',
        bg = '#43516C', fg = 'white', font = 'fantasy 20 bold',
        borderwidth=5, relief="ridge", padx=20, pady=10).place(relx=.8, rely=.2,anchor= CENTER, width=200)

        insertCode = IntVar()
        entry1 = ttk.Entry(self, textvariable=insertCode, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.66, rely=.2,anchor= CENTER, width=100)


        Label(self, text='الخدمه',
        bg = '#43516C', fg = 'white', font = 'fantasy 20 bold',
        borderwidth=5, relief="ridge", padx=20, pady=10).place(relx=.8, rely=.3,anchor= CENTER, width=200)

        insertService = tk.StringVar()
        menu2 = ttk.Combobox(self, textvariable = insertService, font=('Calibri', 20,'bold'), values = services)
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu2.place(relx=.5, rely=.3,anchor= CENTER, width=600)
        

        Label(self, text='التكلفه',
        bg = '#43516C', fg = 'white', font = 'fantasy 20 bold', borderwidth=5, relief="ridge",
        padx=20, pady=10).place(relx=.8, rely=.6,anchor= CENTER, width=200)

        insertCost = IntVar()
        entry1 = ttk.Entry(self, textvariable=insertCost, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.5, rely=.6,anchor= CENTER, width=600, height=50)


        Label(self, text='المدفوع',
        bg = '#43516C', fg = 'white', font = 'fantasy 20 bold', borderwidth=5, relief="ridge",
        padx=20, pady=10).place(relx=.8, rely=.5,anchor= CENTER, width=200)

        insertAmount = IntVar()
        entry1 = ttk.Entry(self, textvariable=insertAmount, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.5, rely=.5,anchor= CENTER, width=600, height=50)
        
        
        Label(self, text='ملاحظات',
        bg = '#43516C', fg = 'white', font = 'fantasy 20 bold', borderwidth=5, relief="ridge",
        padx=20, pady=10).place(relx=.8, rely=.7,anchor= CENTER, width=200)

        insertComment = StringVar()
        entry1 = ttk.Entry(self, textvariable=insertComment, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.5, rely=.7,anchor= CENTER, width=600, height=50)


        Label(self, text='التاريخ',
        bg = '#43516C', fg = 'white', font = 'fantasy 20 bold', borderwidth=5, relief="ridge",
        padx=20, pady=10).place(relx=.8, rely=.4,anchor= CENTER, width=200)

        date = DateEntry(self,width=30,bg="darkblue",fg="white", font = ('fantasy', 25, 'bold'),
        year=datetime.now().year,month=datetime.now().month,day=datetime.now().day
        ,locale='en_US', date_pattern='dd/MM/yyyy')
        date.place(relx=.625, rely=.4,anchor= CENTER, width=200, height=50)
        
        
        entries = []
        for sheet in wb.worksheets:
                if sheet.cell(row=1,column=1).value != None :
                        entries.append(str(sheet.cell(row=1,column=1).value)) 

        codeentries = []
        for sheet in wb.worksheets:
                if sheet.cell(row=1,column=2).value != None :
                        codeentries.append(sheet.cell(row=1,column=2).value)                

        def insert():
                try:
                        insertCost.get()  
                except:
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال خانة التكلفه بطريقه صحيحه ') 
                        return
                try:
                        insertAmount.get()
                except:
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال خانة المدفوع بطريقه صحيحه ') 
                        return              
                if insertCost.get() == "" or insertName.get() == "" or insertService.get() == "" or  insertService.get() == None:
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال الخانات المطلوبه ')
                        return
                if insertName.get() in entries or insertCode.get() in codeentries:
                        self.destroy()
                        messagebox.showerror('Already Exists!','الاسم او الكود موجودين بالفعل ') 
                        return
                
                else:        
                        ws = wb.create_sheet(f'{insertName.get()} ({insertCode.get()})')
                        ws.title = f'{insertName.get()} ( {insertCode.get()} )'
                        ws.append([insertName.get(), insertCode.get(), 1])
                        ws.append([])
                        ws.append(['التاريخ', 'الخدمه','التكلفه','المدفوع','الرصيد', 'الملاحظات', 'مسلسل'])
                        ws.append([date.get_date.strftime("%d/%m/%Y"),
                        insertService.get(),
                        insertCost.get(),
                        insertAmount.get(), 
                        (insertCost.get()) - insertAmount.get() ,
                        insertComment.get(), 1,
                        date.get_date.year,date.get_date.month, date.get_date.day])
                        wb.save(file)
                        messagebox.showinfo('Done','تم التسجيل بنجاح ')
                        self.destroy()

        
        Button(self, height = 1, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 20 bold', bd = '8px solid #DBA531', 
                text='اضافه',
                command=insert).place(relx=.5, rely=.85,anchor= CENTER)
        
        
        Button(self, height = 1, width = 10, bg = 'grey', fg = 'white',
        activebackground='#43516C', font = 'fantasy 15 bold', bd = '8px solid #DBA531', 
                text='اغلاق',
                command=self.destroy).place(relx=.1, rely=.9,anchor= CENTER)

# -------------------------------------------------------------------------------------------------------

class Add(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self.geometry('1600x900')
        self.title('اضافة عميل')
        self['bg']='#E5E8C7'

        entries = []
        for sheet in wb.worksheets:
                if sheet.cell(row=1,column=1).value != None :
                        entries.append(f'{str(sheet.cell(row=1,column=1).value)} ({str(sheet.cell(row=1,column=2).value)})')
                        
        Label(self, text='الاسم',
        bg = '#43516C', fg = 'white', font = 'fantasy 20 bold', borderwidth=5, relief="ridge", 
        padx=20, pady=10 ).place(relx=.8, rely=.1,anchor= CENTER, width=200)


        addName = StringVar()
        menu1 = ttk.Combobox(self, font=('Calibri', 20,'bold'),values = entries, textvariable = addName )
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu1.place(relx=.5, rely=.1,anchor= CENTER, width=600)
        menu1['state'] = 'readonly'


        Label(self, text='الخدمه',
        bg = '#43516C', fg = 'white', font = 'fantasy 20 bold',
        borderwidth=5, relief="ridge", padx=20, pady=10).place(relx=.8, rely=.2,anchor= CENTER, width=200)

        addService = tk.StringVar()
        menu2 = ttk.Combobox(self, textvariable = addService, font=('Calibri', 20,'bold'), values = services)
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu2.place(relx=.5, rely=.2,anchor= CENTER, width=600)
        

        Label(self, text='التكلفه',
        bg = '#43516C', fg = 'white', font = 'fantasy 20 bold', borderwidth=5, relief="ridge",
        padx=20, pady=10).place(relx=.8, rely=.3,anchor= CENTER, width=200)

        addCost = IntVar()
        entry1 = ttk.Entry(self, textvariable=addCost, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.5, rely=.3,anchor= CENTER, width=600, height=50)


        Label(self, text='المدفوع',
        bg = '#43516C', fg = 'white', font = 'fantasy 20 bold', borderwidth=5, relief="ridge",
        padx=20, pady=10).place(relx=.8, rely=.5,anchor= CENTER, width=200)

        addAmount = IntVar()
        entry1 = ttk.Entry(self, textvariable=addAmount, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.5, rely=.5,anchor= CENTER, width=600, height=50)
        

        Label(self, text='ملاحظات',
        bg = '#43516C', fg = 'white', font = 'fantasy 20 bold', borderwidth=5, relief="ridge",
        padx=20, pady=10).place(relx=.8, rely=.6,anchor= CENTER, width=200)

        addComment = StringVar()
        entry1 = ttk.Entry(self, textvariable=addComment, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.5, rely=.6,anchor= CENTER, width=600, height=50)


        Label(self, text='التاريخ',
        bg = '#43516C', fg = 'white', font = 'fantasy 20 bold', borderwidth=5, relief="ridge",
        padx=20, pady=10).place(relx=.8, rely=.4,anchor= CENTER, width=200)

        date = DateEntry(self,width=30,bg="darkblue",fg="white", font = ('fantasy', 25, 'bold'),
        year=datetime.now().year,month=datetime.now().month,day=datetime.now().day
        ,locale='en_US', date_pattern='dd/MM/yyyy')
        date.place(relx=.625, rely=.4,anchor= CENTER, width=200, height=50)
        
        
        def add():

                if addCost.get() == '' or addName.get() == '' or addService.get() == '':
                        messagebox.showerror('Invalid!','من فضلك قم بادخال الخانات المطلوبه ')
                        self.destroy()
                        return
                try:
                        addAmount.get() 
                except:
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال خانة المدفوع بطريقه صحيحه ')
                try:
                        addCost.get() 
                except:
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال خانة التكلفه بطريقه صحيحه ')
                try:
                        addService.get() 
                except:
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال خانة الخدمه بطريقه صحيحه ')                                    
                flag = False
                finalAmount = 0
                for sheet in wb.worksheets:
                        ws = sheet
                        if f'{str(sheet.cell(row=1,column=1).value)} ({str(sheet.cell(row=1,column=2).value)})' == addName.get():
                                for i in range (1, ws.max_row+1):
                                        if ws.cell(column=7,row=i).value == ws.cell(column=3,row=1).value:
                                                finalAmount = ws.cell(column=5,row=i).value
                                ws['C1'] = int(ws.cell(column=3, row=1).value) + 1
                                ws.append([date.get_date.strftime("%d/%m/%Y"),
                                addService.get(),addCost.get(),
                                addAmount.get(),
                                ((int(finalAmount) + addCost.get()) - addAmount.get()),
                                addComment.get(),
                                ws.cell(column=3, row=1).value,
                                date.get_date.year,date.get_date.month, date.get_date.day])
                                wb.save(file)
                                flag = True
                                messagebox.showinfo('Done','تم التسجيل بنجاح ')
                                self.destroy()
                                break
                if not flag:
                        messagebox.showerror('Not Exists!','الاسم غير موجود')         
        
        Button(self, height = 1, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 20 bold', bd = '8px solid #DBA531', 
                text='اضافه',
                command=add).place(relx=.5, rely=.75,anchor= CENTER)
        
        
        Button(self, height = 1, width = 10, bg = 'grey', fg = 'white',
        activebackground='#43516C', font = 'fantasy 15 bold', bd = '8px solid #DBA531', 
                text='اغلاق',
                command=self.destroy).place(relx=.1, rely=.9,anchor= CENTER)

# -----------------------------------------------------------------------------------------------------------------------------

class Delete(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self.geometry('1600x900')
        self.title('حذف')
        self['bg']='#E5E8C7'


        entries = []
        for sheet in wb.worksheets:
                if sheet.cell(row=1,column=1).value != None :
                        entries.append(f'{str(sheet.cell(row=1,column=1).value)} ({str(sheet.cell(row=1,column=2).value)})')

        if len(entries) < 1:
                entries.append('لا يوجد عملاء')   
                        

        deleteName = StringVar()
        menu = ttk.Combobox(self, font=('Calibri', 20,'bold'),textvariable=deleteName, values = entries )
        menu.place(relx=.5, rely=.25,anchor= CENTER, width=600)
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu['state'] = 'readonly'

        def delete():
                for sheet in wb.worksheets:
                        ws = sheet
                        if f'{str(ws.cell(row=1,column=1).value)} ({str(ws.cell(row=1,column=2).value)})' == deleteName.get():
                                wb.remove(sheet)
                                self.destroy()
                                messagebox.showinfo('Done','تم الحذف بنجاح')
                                wb.save(file)


        Button(self, height = 1, width = 15, bg = 'red', fg = 'white',
        activebackground='#43516C', font = 'fantasy 20 bold', bd = '8px solid #DBA531', 
                text='حذف',
                command=delete).place(relx=.5, rely=.5,anchor= CENTER)
        

        Button(self, height = 1, width = 10, bg = 'grey', fg = 'white',
        activebackground='#43516C', font = 'fantasy 15 bold', bd = '8px solid #DBA531', 
                text='اغلاق',
                command=self.destroy).place(relx=.5, rely=.8,anchor= CENTER)



class Main(tk.Tk):
    def __init__(self):
        super().__init__()

        self.geometry('1600x900')
        self.title('Main Window')
        self['bg']='#E5E8C7'

        Label(self, text='Mr.Wagdy Latif For Accounting Services',
        bg = '#D85426', fg = 'white', font = 'fantasy 30 bold', borderwidth=20, relief="ridge", padx=20, pady=40).place(relx=.5, rely=.1,anchor= CENTER)


        Button(self, height = 1, width = 13, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text='بحث',
                command=self.open_search).place(relx=.5, rely=.4,anchor= CENTER)

        Button(self, height = 2, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text=' خدمه جديده ',
                command=self.open_add).place(relx=.8, rely=.7,anchor= CENTER) 

        Button(self, height = 2, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text='الايرادات',
                command=self.open_revenue).place(relx=.2, rely=.7,anchor= CENTER)       

        Button(self, height = 2, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text=' عميل جديد',
                command=self.open_insert).place(relx=.8, rely=.35,anchor= CENTER)

        Button(self, height = 2, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text=' مدفوعات',
                command=self.open_payment).place(relx=.2, rely=.35,anchor= CENTER)

        Button(self, height = 1, width = 13, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text='حذف',
                command=self.open_delete).place(relx=.5, rely=.6,anchor= CENTER)

        Button(self, height = 1, width = 10, bg = 'grey', fg = 'white',
        activebackground='#43516C', font = 'fantasy 15 bold', bd = '8px solid #DBA531', 
                text='اغلاق',
                command=self.destroy).place(relx=.5, rely=.9,anchor= CENTER)


    def open_revenue(self):
                window = Revenue(self)
                window.grab_set()  

    def open_payment(self):
                window = Payment(self)
                window.grab_set()            

    def open_search(self):
                window = Search(self)
                window.grab_set()

    def open_insert(self):
                window = Insert(self)
                window.grab_set()

    def open_delete(self):
                window = Delete(self)
                window.grab_set()  

    def open_add(self):
                window = Add(self)
                window.grab_set()

if __name__ == "__main__":
        app = Main()
        app.mainloop()