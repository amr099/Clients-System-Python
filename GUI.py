from tkinter import *
from tkinter import Entry, Label, Tk, ttk, messagebox
import tkinter as tk
from openpyxl import load_workbook
from datetime import date, datetime
from openpyxl.xml.constants import MAX_ROW
from tkcalendar import DateEntry


file = 'xl.xlsx' # change ... 
try:
  load_workbook(file)
except:
  messagebox.showerror('File Not Found!','غير موجود xl.xlsx  ملف الاكسيل')

wb = load_workbook(file)

services = ['اتعاب لجنه داخلية','اضافة سياره','اعداد و مراجعة ميزانيه','اقرار ضرائب عامه'
        ,'اقرار ضرائب قيمه مضافه', 'اقرار ضرائب مرتبات', 'بطاقه ضريبيه'
        ,'تجديد اشتراك البوابه الالكترونيه', 'تحت الحساب','تسوية ملف ضريبى', 'تعديل النشاط',
        'جواب مرور', 'حفظ الملف بالضرائب','رسوم استخراج مستوردين', 'رسوم البوابه الالكترونيه'
        ,'رسوم تجديد مستوردين', 'سجل تجارى','سداد ضرائب عامه'
        , 'سداد ضرائب قيمه مضافه','شطب سجل تجارى', 'شهاده بالموقف الضريبى'
        ,'شهادة دخل 1', 'شهادة دخل 2', 'عمل موقع الكترونى', 'غرفه تجاريه',
        'فحص ضرائب عامه', 'فحص ضريبة قيمه مضافه','لجنة طعن ضرائب عامه', 'مركز مالى',
        'ميزانيه عموميه', 'نماذج 41 ض',]


class Expenses(tk.Toplevel):
      def __init__(self, parent):
        super().__init__(parent)

        self.geometry('1600x900')
        self.title('المصروفات')
        self['bg']='#E5E8C7'

        def expenses():
                treev = ttk.Treeview(self, selectmode ='browse', style="mystyle.Treeview", height=900)
                treev.pack()
                
                style = ttk.Style()
                
                style.configure("mystyle.Treeview", background = '#E5E8C7' ,rowheight=40,
                highlightthickness=0, bd=0, font=('Helvetica', 14)) 
                style.configure("mystyle.Treeview.Heading", font=('Helvetica', 20,'bold')) 
                style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})]) 

                verscrlbar = ttk.Scrollbar(self,
                                orient ="vertical",
                                command = treev.yview)

                verscrlbar.pack(side ='left', fill ='x')   
                treev.configure(xscrollcommand = verscrlbar.set)

                treev["columns"] = ("1", "2")

                treev['show'] = 'headings'
        
                treev.column("1", width = 400,anchor ='e')
                treev.column("2", width = 1150,anchor ='e')
                
                treev.heading("1", text ="المصروف")
                treev.heading("2", text ="الحساب")
                
                for sheet in wb.worksheets:
                        exps = []
                        if sheet.title == 'مصروفات اداريه':
                                ws = sheet
                                exp = ''
                                for i in range (1, ws.max_row+1):
                                        if ws.cell(row=i,column=3).value not in exps:
                                                exps.append(ws.cell(row=i,column=3).value)
                                                exp = str(ws.cell(row=i,column=3).value)
                                                amount = 0 
                                                for i in range (1, ws.max_row+1):
                                                        if str(ws.cell(row=i,column=3).value) == exp:
                                                                amount += int(ws.cell(row=i,column=2).value)
                                                treev.insert("", 'end', text ="L7",
                                                        values =(amount,exp), tags = ("expense",))
                                                treev.tag_configure('expense', background='#e1dddd', font=('Helvetica', 26))

                
        Button(self, height = 1, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 20 bold', bd = '8px solid #DBA531', 
                text='المصروفات',
                command=expenses).place(relx=.5, rely=.4,anchor= CENTER)                                

        Button(self, height = 1, width = 10, bg = 'grey', fg = 'white',
        activebackground='#43516C', font = 'fantasy 15 bold', bd = '8px solid #DBA531', 
                text='اغلاق',
                command=self.destroy).place(relx=.5, rely=.9,anchor= CENTER)

# -----------------------------------------------------------------------------------------------------------------------------


class Expenses_form(tk.Toplevel):
      def __init__(self, parent):
        super().__init__(parent)

        self.geometry('1600x900')
        self.title('المصروفات')
        self['bg']='#E5E8C7'


        entries = []
        for sheet in wb.worksheets:
                if sheet.cell(row=1,column=1).value != None and sheet.title != 'مصروفات اداريه' :
                        entries.append(f'{str(sheet.cell(row=1,column=1).value)} ({str(sheet.cell(row=1,column=2).value)})')                
        if len(entries) < 1:
                entries.append('لا يوجد عملاء')   

        Label(self, text='الاسم',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.75, rely=.1,anchor= CENTER, width=200)

        def check_input(event):
                value = event.widget.get()
                if value == '':
                        menu['values'] = entries
                else:
                        data = []
                        for item in entries:
                                if value.lower() in item.lower():
                                        data.append(item)

                        menu['values'] = data

        payName = StringVar()
        menu = ttk.Combobox(self, textvariable= payName, font=('Helvetica', 20,'bold'),values = entries )
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu.place(relx=.5, rely=.1,anchor= CENTER, width=600)
        menu['values'] = entries
        menu.bind('<KeyRelease>', check_input)


        Label(self, text='المصروف',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.75, rely=.25,anchor= CENTER, width=200)

        payAmount = IntVar()
        entry1 = ttk.Entry(self, textvariable = payAmount, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.5, rely=.25,anchor= CENTER, width=600, height=50)
        

        Label(self, text='ملاحظات',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.75, rely=.4,anchor= CENTER, width=200)

        payComment = StringVar()
        entry1 = ttk.Entry(self, textvariable=payComment, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.5, rely=.4,anchor= CENTER, width=600, height=50)

        Label(self, text='التاريخ',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.75, rely=.55,anchor= CENTER, width=200)

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
                        messagebox.showerror('Invalid!','من فضلك قم بادخال الخانات المطلوبه   ')
                        self.destroy()
                        return      

                flag = False
                if payName.get() != '':
                        for sheet in wb.worksheets:
                                ws = sheet
                                if f'{str(sheet.cell(row=1,column=1).value)} ({str(sheet.cell(row=1,column=2).value)})' == payName.get():
                                        ws.append(['','','','','','','','','','',
                                        date.get_date().strftime("%d/%m/%Y"),
                                        payAmount.get(),
                                        payComment.get(),
                                        date.get_date().year, date.get_date().month, date.get_date().day])
                                        wb.save(file)
                                        flag = True
                                        self.destroy()
                                        messagebox.showinfo('Done','تم التسجيل بنجاح ')
                                        break
                                        
                        if not flag:
                                messagebox.showerror('Not Exists!','الاسم غير موجود') 
                                self.destroy()
                else:
                        flag = False
                        for sheet in wb.worksheets:
                                if sheet.title == 'مصروفات اداريه' :
                                        ws = sheet
                                        ws.append([date.get_date().strftime("%d/%m/%Y"),
                                        payAmount.get(),
                                        payComment.get(),
                                        date.get_date().year, date.get_date().month, date.get_date().day])
                                        wb.save(file)
                                        self.destroy()
                                        messagebox.showinfo('Done','تم التسجيل بنجاح ')
                                        flag = True
                                        break
                        if not flag:
                                ws = wb.create_sheet('مصروفات اداريه')                
                                ws.append([date.get_date().strftime("%d/%m/%Y"),
                                        payAmount.get(),
                                        payComment.get(),
                                        date.get_date().year, date.get_date().month, date.get_date().day])
                                wb.save(file)
                                self.destroy()
                                messagebox.showinfo('Done','تم التسجيل بنجاح ')


        Button(self, height = 1, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 20 bold', bd = '8px solid #DBA531', 
                text='تسجيل',
                command=payment).place(relx=.5, rely=.75,anchor= CENTER)


        Button(self, height = 1, width = 10, bg = 'grey', fg = 'white',
        activebackground='#43516C', font = 'fantasy 15 bold', bd = '8px solid #DBA531', 
                text='اغلاق',
                command=self.destroy).place(relx=.1, rely=.9,anchor= CENTER)        
        
# -----------------------------------------------------------------------------------------------------------------------------
        

class Revenue(tk.Toplevel):
      def __init__(self, parent):
        super().__init__(parent)

        self.geometry('1600x900')
        self.title('الايرادات')
        self['bg']='#E5E8C7'


        Label(self, text='من',
        bg = '#43516C', fg = 'white', font = 'fantasy 30 bold',
        borderwidth=5, relief="ridge", padx=20, pady=10).place(relx=.7, rely=.4,anchor= CENTER, width=400)

        startdate = DateEntry(self,width=30,bg="darkblue",fg="white", font = ('fantasy', 25, 'bold'),
        year=datetime.now().year,month=datetime.now().month,day=datetime.now().day
        ,locale='en_US', date_pattern='dd/MM/yyyy')
        startdate.place(relx=.7, rely=.5,anchor= CENTER, width=200, height=50)
        

        Label(self, text='الى',
        bg = '#43516C', fg = 'white', font = 'fantasy 30 bold',
        borderwidth=5, relief="ridge", padx=20, pady=10).place(relx=.3, rely=.4,anchor= CENTER, width=400)

        enddate = DateEntry(self,width=30,bg="darkblue",fg="white", font = ('fantasy', 25, 'bold'),
        year=datetime.now().year,month=datetime.now().month,day=datetime.now().day
        ,locale='en_US', date_pattern='dd/MM/yyyy')
        enddate.place(relx=.3, rely=.5,anchor= CENTER, width=200, height=50)
        
        def clients():
                treev = ttk.Treeview(self, selectmode ='browse', style="mystyle.Treeview", height=900)
                treev.pack()
                
                style = ttk.Style()
                
                style.configure("mystyle.Treeview", background = '#E5E8C7' ,rowheight=40,
                highlightthickness=0, bd=0, font=('Helvetica', 14)) 
                style.configure("mystyle.Treeview.Heading", font=('Helvetica', 20,'bold')) 
                style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})]) 

                verscrlbar = ttk.Scrollbar(self,
                                orient ="vertical",
                                command = treev.yview)

                verscrlbar.pack(side ='left', fill ='x')   
                treev.configure(xscrollcommand = verscrlbar.set)

                treev["columns"] = ("1", "2", "3", "4", "5")

                treev['show'] = 'headings'
        
                treev.column("1", width = 470,anchor ='e')
                treev.column("2", width = 150,anchor ='e')
                treev.column("3", width = 170,anchor ='e')
                treev.column("4", width = 650,anchor ='e')
                treev.column("5", width = 150,anchor ='e')
                
                treev.heading("1", text ="صافى الربح")
                treev.heading("2", text ="المصروف")
                treev.heading("3", text ="الايراد")
                treev.heading("4", text ="الاسم")
                treev.heading("5", text ="الكود")

                for sheet in wb.worksheets:
                        if sheet.title != 'مصروفات اداريه':
                                client_expenses = 0
                                client_revenue = 0
                                amount = 0
                                ws = sheet
                                if (ws.cell(row=1,column=2).value) != None:
                                        for i in range(4, ws.max_row+1):
                                                try:
                                                        celldate = date(ws.cell(row=i, column=8).value, ws.cell(row=i, column=9).value, ws.cell(row=i, column=10).value)
                                                except:
                                                        continue
                                                if celldate >= startdate.get_date() and celldate <= enddate.get_date():
                                                        if ws.cell(row=i, column=4).value != None:
                                                                client_revenue += ws.cell(row=i, column=4).value
                                        for i in range(4, ws.max_row+1):
                                                try:
                                                        celldate = date(ws.cell(row=i, column=14).value, ws.cell(row=i, column=15).value, ws.cell(row=i, column=16).value)
                                                except:
                                                        continue
                                                if celldate >= startdate.get_date() and celldate <= enddate.get_date():
                                                        if ws.cell(row=i, column=13).value != None:
                                                                client_expenses += ws.cell(row=i, column=12).value
                                        amount = client_revenue - client_expenses
                                        treev.insert("", 'end', text ="L7",
                                                        values =(
                                                        '-' if int(amount) == None else str(amount),
                                                        '-' if int(client_expenses) == None else str(client_expenses),
                                                        '-' if int(client_revenue) == None else str(client_revenue),
                                                        '-' if (ws.cell(row=1,column=1).value) == None else str(ws.cell(row=1,column=1).value),
                                                        '-' if (ws.cell(row=1,column=2).value) == None else str(ws.cell(row=1,column=2).value),
                                                        ), tags = ('table',))
                        
                        treev.tag_configure('table', background='#eee', font=('Helvetica', 26)) 

        def revenue():      
                total_revenue = 0
                total_expenses = 0
                income = 0
                for sheet in wb.worksheets:
                        ws = sheet
                        for i in range(4, ws.max_row+1):
                                try:
                                        celldate = date(ws.cell(row=i, column=8).value, ws.cell(row=i, column=9).value, ws.cell(row=i, column=10).value)
                                except:
                                        continue
                                if celldate >= startdate.get_date() and celldate <= enddate.get_date():
                                        if ws.cell(row=i, column=4).value != None:
                                                total_revenue += ws.cell(row=i, column=4).value
                for sheet in wb.worksheets:
                        if sheet.title == 'مصروفات اداريه' :
                                ws = sheet
                                for i in range(0, ws.max_row+1):
                                        try:    
                                                celldate = date(ws.cell(row=i, column=4).value, ws.cell(row=i, column=5).value, ws.cell(row=i, column=6).value)
                                        except:
                                                continue
                                        if celldate >= startdate.get_date() and celldate <= enddate.get_date():
                                                if ws.cell(row=i, column=2).value != None:
                                                        total_expenses += ws.cell(row=i, column=2).value
                        else:
                                ws = sheet
                                for i in range(0, ws.max_row+1):
                                        try:    
                                                celldate = date(ws.cell(row=i, column=14).value, ws.cell(row=i, column=15).value, ws.cell(row=i, column=16).value)
                                        except:
                                                continue
                                        if celldate >= startdate.get_date() and celldate <= enddate.get_date():
                                                if ws.cell(row=i, column=12).value != None:
                                                        total_expenses += ws.cell(row=i, column=12).value                              

                income = total_revenue - total_expenses 

                Label(self, text='الايرادات',
                bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 20 bold').place(relx=.8, rely=.65,anchor= CENTER, width=200)
                Label(self, text='المصروفات',
                bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 20 bold').place(relx=.5, rely=.65,anchor= CENTER, width=200)
                Label(self, text='صافى الربح',
                bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 20 bold').place(relx=.2, rely=.65,anchor= CENTER, width=200)
                
                Label(self, text=str(total_revenue),
                bg = '#E85662', fg = 'white', font = 'fantasy 40 bold',
                borderwidth=5, relief="ridge", padx=20, pady=10).place(relx=.8, rely=.75,anchor= CENTER, width=280)

                Label(self, text='_',
                bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 40 bold').place(relx=.65, rely=.72,anchor= CENTER, width=200)
                
                Label(self, text=str(total_expenses),
                bg = '#E85662', fg = 'white', font = 'fantasy 40 bold',
                borderwidth=5, relief="ridge", padx=20, pady=10).place(relx=.5, rely=.75,anchor= CENTER, width=280)

                Label(self, text='=',
                bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 40 bold').place(relx=.35, rely=.75,anchor= CENTER, width=200)
                

                Label(self, text=str(income),
                bg = '#E85662', fg = 'white', font = 'fantasy 40 bold',
                borderwidth=5, relief="ridge", padx=20, pady=10).place(relx=.2, rely=.75,anchor= CENTER, width=280)
                                

        Button(self, height = 1, width = 18, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 20 bold', bd = '8px solid #DBA531', 
                text='اجمالى الايرادات و المصروفات',
                command=revenue).place(relx=.5, rely=.55,anchor= CENTER)

        Button(self, height = 2, width = 20, bg = '#D85426', fg = 'white',
        activebackground='#D85426', font = 'fantasy 24 bold', bd = '8px solid #DBA531', 
                text='تفاصيل الايرادات',
                command= clients).place(relx=.5, rely=.2,anchor= CENTER)        

        Button(self, height = 1, width = 10, bg = 'grey', fg = 'white',
        activebackground='#43516C', font = 'fantasy 15 bold', bd = '8px solid #DBA531', 
                text='اغلاق',
                command=self.destroy).place(relx=.1, rely=.9,anchor= CENTER)        
        

# -----------------------------------------------------------------------------------------------------------------------------

class Payment(tk.Toplevel):
      def __init__(self, parent):
        super().__init__(parent)

        self.geometry('1600x900')
        self.title('الدفوعات')
        self['bg']='#E5E8C7'


        entries = []
        for sheet in wb.worksheets:
                if sheet.cell(row=1,column=1).value != None and sheet.title != 'مصروفات اداريه':
                        entries.append(f'{str(sheet.cell(row=1,column=1).value)} ({str(sheet.cell(row=1,column=2).value)})')                
        if len(entries) < 1:
                entries.append('لا يوجد عملاء')   

        Label(self, text='الاسم',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.8, rely=.1,anchor= CENTER, width=200)


        def check_input(event):
                value = event.widget.get()
                if value == '':
                        menu['values'] = entries
                else:
                        data = []
                        for item in entries:
                                if value.lower() in item.lower():
                                        data.append(item)

                        menu['values'] = data

        payName = StringVar()
        menu = ttk.Combobox(self, textvariable= payName, font=('Helvetica', 20,'bold'),values = entries )
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu.place(relx=.55, rely=.1,anchor= CENTER, width=600)
        menu['values'] = entries
        menu.bind('<KeyRelease>', check_input)


        Label(self, text='المدفوع',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.8, rely=.25,anchor= CENTER, width=200)

        payAmount = IntVar()
        entry1 = ttk.Entry(self, textvariable = payAmount, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.55, rely=.25,anchor= CENTER, width=600, height=50)
        

        Label(self, text='ملاحظات',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.8, rely=.4,anchor= CENTER, width=200)

        payComment = StringVar()
        entry1 = ttk.Entry(self, textvariable=payComment, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.55, rely=.4,anchor= CENTER, width=600, height=50)

        Label(self, text='التاريخ',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.8, rely=.55,anchor= CENTER, width=200)

        date = DateEntry(self,width=30,bg="darkblue",fg="white",year=datetime.now().year,
        month=datetime.now().month,day=datetime.now().day
        ,locale='en_US', date_pattern='dd/MM/yyyy',font = ('fantasy', 25, 'bold'))
        date.place(relx=.675, rely=.55,anchor= CENTER, width=200, height=50)
        
        
        
        def payment():
                try:
                        payAmount.get() and payName.get()
                except:
                        messagebox.showerror('Invalid!','من فضلك قم بادخال الخانات بطريقه صحيحه ') 
                        self.destroy()
                        return      
                if payAmount.get() == '' or payAmount.get() == None :
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

        self.geometry('1600x850')
        self.title('البحث عن عميل')
        self['bg']='#E5E8C7'

        entries = []
        for sheet in wb.worksheets:
                if sheet.cell(row=1,column=1).value != None and sheet.title != 'مصروفات اداريه':
                        entries.append(f'{str(sheet.cell(row=1,column=1).value)} ({str(sheet.cell(row=1,column=2).value)})')

        if len(entries) < 1:
                entries.append('لا يوجد عملاء') 

        
        def check_input(event):
                value = event.widget.get()
                if value == '':
                        menu['values'] = entries
                else:
                        data = []
                        for item in entries:
                                if value.lower() in item.lower():
                                        data.append(item)

                        menu['values'] = data


        searchName = StringVar()
        menu = ttk.Combobox(self, textvariable= searchName, font=('Helvetica', 20,'bold'),values = entries )
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu.place(relx=.5, rely=.4,anchor= CENTER, width=600)
        menu['values'] = entries
        menu.bind('<KeyRelease>', check_input)


        def search():
                treev = ttk.Treeview(self, selectmode ='browse', style="mystyle.Treeview", height=900)
                treev.pack()
                
                style = ttk.Style()
                
                style.configure("mystyle.Treeview", background = '#E5E8C7' ,rowheight=40,
                highlightthickness=0, bd=0, font=('Helvetica', 14)) 
                style.configure("mystyle.Treeview.Heading", font=('Helvetica', 20,'bold')) 
                style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})]) 

                verscrlbar = ttk.Scrollbar(self,
                                orient ="vertical",
                                command = treev.yview)

                verscrlbar.pack(side ='left', fill ='x')   
                treev.configure(xscrollcommand = verscrlbar.set)

                treev["columns"] = ("1", "2", "3", "4", "5", "6")

                treev['show'] = 'headings'
        
                treev.column("6", width = 200,anchor ='e')
                treev.column("5", width = 420,anchor ='e')
                treev.column("4", width = 160,anchor ='e')
                treev.column("3", width = 160,anchor ='e')
                treev.column("2", width = 160,anchor ='e')
                treev.column("1", width = 500,anchor ='e')
                
                treev.heading("6", text ="التاريخ")
                treev.heading("5", text ="الخدمه")
                treev.heading("4", text ="التكلفه")
                treev.heading("3", text ="المدفوع")
                treev.heading("2", text ="الرصيد")
                treev.heading("1", text ="الملاحظات")

                flag = False
                for sheet in wb.worksheets:
                        ws = sheet
                        if f'{str(ws.cell(row=1,column=1).value)} ({str(ws.cell(row=1,column=2).value)})' == searchName.get():
                                ws = sheet
                                for i in range (4, ws.max_row+1):
                                        if (ws.cell(row=i,column=1).value) != None:
                                                if i%2 == 0:
                                                        treev.insert("", 'end', text ="L7",
                                                                        values =(
                                                                        '-' if (ws.cell(row=i,column=6).value) == None else str(ws.cell(row=i,column=6).value), 
                                                                        '-' if (ws.cell(row=i,column=5).value) == None else str(ws.cell(row=i,column=5).value),
                                                                        '-' if (ws.cell(row=i,column=4).value) == None else str(ws.cell(row=i,column=4).value),
                                                                        '-' if (ws.cell(row=i,column=3).value) == None else str(ws.cell(row=i,column=3).value),
                                                                        '-' if (ws.cell(row=i,column=2).value) == None else str(ws.cell(row=i,column=2).value),
                                                                        '-' if (ws.cell(row=i,column=1).value) == None else str(ws.cell(row=i,column=1).value)), tags = ('even',))
                                                else:
                                                        treev.insert("", 'end', text ="L7",
                                                                        values =(
                                                                        '-' if (ws.cell(row=i,column=6).value) == None else str(ws.cell(row=i,column=6).value), 
                                                                        '-' if (ws.cell(row=i,column=5).value) == None else str(ws.cell(row=i,column=5).value),
                                                                        '-' if (ws.cell(row=i,column=4).value) == None else str(ws.cell(row=i,column=4).value),
                                                                        '-' if (ws.cell(row=i,column=3).value) == None else str(ws.cell(row=i,column=3).value),
                                                                        '-' if (ws.cell(row=i,column=2).value) == None else str(ws.cell(row=i,column=2).value),
                                                                        '-' if (ws.cell(row=i,column=1).value) == None else str(ws.cell(row=i,column=1).value)), tags = ('odd',))                        
                                                flag = True
                                        
                        treev.tag_configure('odd', background='#e1dddd', font=('Helvetica', 26)) 
                        treev.tag_configure('even', background= '#f5f3f3', font=('Helvetica', 26))
                self.title( searchName.get())
                if not flag :
                        self.destroy()
                        messagebox.showerror('Not Exists!','الاسم غير موجود')                 
        

        def clients():
                treev = ttk.Treeview(self, selectmode ='browse', style="mystyle.Treeview", height=900)
                treev.pack()
                
                style = ttk.Style()
                
                style.configure("mystyle.Treeview", background = '#E5E8C7' ,rowheight=40,
                highlightthickness=0, bd=0, font=('Helvetica', 14)) 
                style.configure("mystyle.Treeview.Heading", font=('Helvetica', 20,'bold')) 
                style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})]) 

                verscrlbar = ttk.Scrollbar(self,
                                orient ="vertical",
                                command = treev.yview)

                verscrlbar.pack(side ='left', fill ='x')   
                treev.configure(xscrollcommand = verscrlbar.set)

                treev["columns"] = ("1", "2", "3", "4", "5", "6")

                treev['show'] = 'headings'
        
                treev.column("1", width = 400,anchor ='e')
                treev.column("2", width = 300,anchor ='e')
                treev.column("3", width = 150,anchor ='e')
                treev.column("4", width = 150,anchor ='e')
                treev.column("5", width = 450,anchor ='e')
                treev.column("6", width = 150,anchor ='e')
                
                treev.heading("1", text ="العنوان")
                treev.heading("2", text ="التليفون")
                treev.heading("3", text ="رقم التسجيل")
                treev.heading("4", text ="الرصيد")
                treev.heading("5", text ="الاسم")
                treev.heading("6", text ="الكود")

                counter = 0
                for sheet in wb.worksheets:
                        counter +=1
                        ws = sheet
                        if (ws.cell(row=1,column=2).value) != None and sheet.title != 'مصروفات اداريه':
                                amount = 0
                                for i in range (4, ws.max_row+1):
                                        try:
                                                amount = int(ws.cell(row=i,column=5).value)
                                        except:
                                                continue   
                                if counter % 2 == 0:             
                                        treev.insert("", 'end', text ="L7",
                                                        values =(
                                                        '-' if (ws.cell(row=1,column=6).value) == None else str(ws.cell(row=1,column=6).value),
                                                        '-' if (ws.cell(row=1,column=5).value) == None else str(ws.cell(row=1,column=5).value),
                                                        '-' if (ws.cell(row=1,column=4).value) == None else str(ws.cell(row=1,column=4).value),
                                                        '-' if int(amount) == None else int(amount),
                                                        '-' if (ws.cell(row=1,column=1).value) == None else str(ws.cell(row=1,column=1).value),
                                                        '-' if (ws.cell(row=1,column=2).value) == None else str(ws.cell(row=1,column=2).value),
                                                        ), tags = ('even',))
                                else:             
                                        treev.insert("", 'end', text ="L7",
                                                        values =(
                                                        '-' if (ws.cell(row=1,column=6).value) == None else str(ws.cell(row=1,column=6).value),
                                                        '-' if (ws.cell(row=1,column=5).value) == None else str(ws.cell(row=1,column=5).value),
                                                        '-' if (ws.cell(row=1,column=4).value) == None else str(ws.cell(row=1,column=4).value),
                                                        '-' if int(amount) == None else int(amount),
                                                        '-' if (ws.cell(row=1,column=1).value) == None else str(ws.cell(row=1,column=1).value),
                                                        '-' if (ws.cell(row=1,column=2).value) == None else str(ws.cell(row=1,column=2).value),
                                                        ), tags = ('odd',))
                                                                
                
                treev.tag_configure('even', background='#e1dddd', font=('Helvetica', 26)) 
                treev.tag_configure('odd', background='#f5f3f3', font=('Helvetica', 26)) 



        
        Button(self, height = 1, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 20 bold', bd = '8px solid #DBA531', 
                text='ابحث',
                command= search).place(relx=.5, rely=.5,anchor= CENTER)

        Button(self, height = 2, width = 20, bg = '#D85426', fg = 'white',
        activebackground='#D85426', font = 'fantasy 24 bold', bd = '8px solid #DBA531', 
                text='العملاء',
                command= clients).place(relx=.5, rely=.2,anchor= CENTER)        
        
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
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.9, rely=.1,anchor= CENTER, width=200)

        insertName = tk.StringVar()
        entry1 = ttk.Entry(self, textvariable=insertName, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.825, rely=.17,anchor= CENTER, width=450)

        Label(self, text='الكود',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.5, rely=.1,anchor= CENTER, width=200)

        insertCode = IntVar()
        entry1 = ttk.Entry(self, textvariable=insertCode, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.5, rely=.17,anchor= CENTER, width=200)

        Label(self, text='رقم التسجيل',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.1, rely=.1,anchor= CENTER, width=200)

        recordNumber = tk.IntVar()
        entry1 = ttk.Entry(self, textvariable=recordNumber, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.1, rely=.17,anchor= CENTER, width=200)


        Label(self, text='رقم التليفون',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.9, rely=.7,anchor= CENTER, width=200)

        phone = tk.StringVar()
        entry1 = ttk.Entry(self, textvariable=phone, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.78, rely=.77,anchor= CENTER, width=600)


        Label(self, text='العنوان',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.375, rely=.7,anchor= CENTER, width=200)


        address = tk.StringVar()
        entry1 = ttk.Entry(self, textvariable=address, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.25, rely=.77,anchor= CENTER, width=600)


        Label(self, text='الخدمه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.9, rely=.3,anchor= CENTER, width=200)

        insertService = tk.StringVar()
        menu2 = ttk.Combobox(self, textvariable = insertService, font=('Helvetica', 20,'bold'), values = services) 
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu2.place(relx=.835, rely=.37,anchor= CENTER, width=400)
        

        Label(self, text='التكلفه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.5, rely=.3,anchor= CENTER, width=200)

        insertCost = IntVar()
        entry1 = ttk.Entry(self, textvariable=insertCost, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.5, rely=.37,anchor= CENTER, width=200)


        Label(self, text='المدفوع',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.1, rely=.3,anchor= CENTER, width=200)

        insertAmount = IntVar()
        entry1 = ttk.Entry(self, textvariable=insertAmount, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.1, rely=.37,anchor= CENTER, width=200)
        
        
        Label(self, text='ملاحظات',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.9, rely=.5,anchor= CENTER, width=200)

        insertComment = StringVar()
        entry1 = ttk.Entry(self, textvariable=insertComment, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.78, rely=.57,anchor= CENTER, width=600,)

        Label(self, text='التاريخ',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.375, rely=.5,anchor= CENTER, width=200)

        date = DateEntry(self,width=30,bg="darkblue",fg="white", font = ('fantasy', 25, 'bold'),
        year=datetime.now().year,month=datetime.now().month,day=datetime.now().day
        ,locale='en_US', date_pattern='dd/MM/yyyy')
        date.place(relx=.25, rely=.57,anchor= CENTER, width=600)
        
        
        
        entries = []
        for sheet in wb.worksheets:
                if sheet.cell(row=1,column=1).value != None and sheet.title != 'مصروفات اداريه':
                        entries.append(str(sheet.cell(row=1,column=1).value)) 

        codeentries = []
        for sheet in wb.worksheets:
                if sheet.cell(row=1,column=2).value != None :
                        codeentries.append(sheet.cell(row=1,column=2).value)                

        def insert():
                # get() validation
                try:
                        insertName.get() and insertCode.get() and recordNumber.get() and insertService.get() and insertCost.get() and insertAmount.get() and phone.get() and address.get() and insertComment.get()
                except:
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال الخانات بطريقه صحيحه ') 
                        return

                # not None Validation
                if insertName.get() == None or insertCode.get() == None or insertService.get() == None or insertCost.get() == None or insertAmount.get() == None :
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال الخانات المطلوبه ') 
                        return        

                # not "" validation        
                if insertName.get() == "" or insertService.get() == "" :
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك نأكد من ادخال الاسم و الخدمه ') 
                        return

                # String validation  
                if not insertName.get().isalpha() and insertService.get().isalpha():
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك نأكد من ادخال الاسم و الخدمه بطريقه صحيحه ') 
                        return   
                
                # phone validation
                if not phone.get().isnumeric:
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك نأكد من ادخال التليفون بطريقه صحيحه ') 
                        return  

                # Case validations
                if insertName.get() in entries or insertCode.get() in codeentries:
                        self.destroy()
                        messagebox.showerror('Already Exists!','الاسم او الكود موجودين بالفعل ') 
                        return
                if len(insertName.get()) > 20:
                        self.destroy()
                        messagebox.showerror('Error!','من فضلك استخدم اسم لا تزيد حروفه عن 20 حرف') 
                        return
                
                else:        
                        ws = wb.create_sheet(f'{insertName.get()} ({insertCode.get()})')
                        ws.title = f'{insertName.get()} ( {insertCode.get()} )'
                        ws.append([insertName.get(), insertCode.get(), 1,str(recordNumber.get()),phone.get(),address.get()])
                        ws.append([])
                        ws.append(['التاريخ', 'الخدمه','التكلفه','المدفوع','الرصيد', 'الملاحظات', 'مسلسل'])
                        ws.append([date.get_date().strftime("%d/%m/%Y"),
                        insertService.get(),
                        insertCost.get(),
                        insertAmount.get(), 
                        (insertCost.get()) - insertAmount.get() ,
                        insertComment.get(), 1,
                        date.get_date().year,date.get_date().month, date.get_date().day])
                        wb.save(file)
                        self.destroy()
                        messagebox.showinfo('Done','تم التسجيل بنجاح ')

        
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
                if sheet.cell(row=1,column=1).value != None and sheet.title != 'مصروفات اداريه':
                        entries.append(f'{str(sheet.cell(row=1,column=1).value)} ({str(sheet.cell(row=1,column=2).value)})')
                        
        Label(self, text='الاسم',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold' ).place(relx=.8, rely=.1,anchor= CENTER, width=200)


        def check_input(event):
                value = event.widget.get()
                if value == '':
                        menu['values'] = entries
                else:
                        data = []
                        for item in entries:
                                if value.lower() in item.lower():
                                        data.append(item)

                        menu['values'] = data

        addName = StringVar()
        menu = ttk.Combobox(self, textvariable= addName, font=('Helvetica', 20,'bold'),values = entries )
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu.place(relx=.55, rely=.1,anchor= CENTER, width=600)
        menu['values'] = entries
        menu.bind('<KeyRelease>', check_input)


        Label(self, text='الخدمه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.8, rely=.22,anchor= CENTER, width=200)

        addService = tk.StringVar()
        menu2 = ttk.Combobox(self, textvariable = addService, font=('Helvetica', 20,'bold'), values = services)
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu2.place(relx=.55, rely=.22,anchor= CENTER, width=600)
        

        Label(self, text='التكلفه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.8, rely=.34,anchor= CENTER, width=200)

        addCost = IntVar()
        entry1 = ttk.Entry(self, textvariable=addCost, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.55, rely=.34,anchor= CENTER, width=600, height=50)


        Label(self, text='المدفوع',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.8, rely=.58,anchor= CENTER, width=200)

        addAmount = IntVar()
        entry1 = ttk.Entry(self, textvariable=addAmount, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.55, rely=.58,anchor= CENTER, width=600, height=50)
        

        Label(self, text='ملاحظات',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.8, rely=.7,anchor= CENTER, width=200)

        addComment = StringVar()
        entry1 = ttk.Entry(self, textvariable=addComment, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.55, rely=.7,anchor= CENTER, width=600, height=50)


        Label(self, text='التاريخ',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.8, rely=.46,anchor= CENTER, width=200)

        c_date = DateEntry(self,width=30,bg="darkblue",fg="white", font = ('fantasy', 25, 'bold'),
        year=datetime.now().year,month=datetime.now().month,day=datetime.now().day
        ,locale='en_US', date_pattern='dd/MM/yyyy')
        c_date.place(relx=.675, rely=.46,anchor= CENTER, width=200, height=50)
        
        
        def add():

                try:
                        addAmount.get(), addCost.get(), addService.get(), addName.get() 
                except:
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال الخانات بطريقه صحيحه ')

                if addCost.get() == '' or addName.get() == '' or addService.get() == '':
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال الخانات المطلوبه ')
                        return
                try:
                        addCost.get() 
                except:
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال خانة التكلفه بطريقه صحيحه ')
                        return

                try:
                        addService.get() 
                except:
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال خانة الخدمه بطريقه صحيحه ')  
                        return

                flag = False
                finalAmount = 0
                for sheet in wb.worksheets:
                        ws = sheet
                        if f'{str(sheet.cell(row=1,column=1).value)} ({str(sheet.cell(row=1,column=2).value)})' == addName.get():
                                inserted_date = date(c_date.get_date().year,c_date.get_date().month, c_date.get_date().day)
                                for i in range (4, ws.max_row+1):
                                        try:
                                                row_date = date(ws.cell(column=8, row=i).value,ws.cell(column=9, row=i).value,ws.cell(column=10, row=i).value,)
                                        except:
                                                continue
                                        if inserted_date < row_date:
                                                rows_list = []
                                                for i in range (4, ws.max_row+1):
                                                        try:
                                                                cur_row_date = date(ws.cell(column=8, row=i).value,ws.cell(column=9, row=i).value,ws.cell(column=10, row=i).value,)
                                                        except:
                                                                continue        
                                                        cur_row = [ws.cell(column=1, row=i).value,ws.cell(column=2, row=i).value,ws.cell(column=3, row=i).value,
                                                        ws.cell(column=4, row=i).value,ws.cell(column=5, row=i).value,ws.cell(column=6, row=i).value,
                                                        ws.cell(column=7, row=i).value,ws.cell(column=8, row=i).value,ws.cell(column=9, row=i).value,
                                                        ws.cell(column=10, row=i).value,cur_row_date]
                                                        rows_list.append(cur_row)
                                                        
                                                rows_list.append([c_date.get_date().strftime("%d/%m/%Y"),
                                                addService.get(),addCost.get(),
                                                addAmount.get(),
                                                ((int(finalAmount) + addCost.get()) - addAmount.get()),
                                                addComment.get(),
                                                ws.cell(column=3, row=1).value,c_date.get_date().year,c_date.get_date().month, c_date.get_date().day, date(c_date.get_date().year,c_date.get_date().month, c_date.get_date().day),])  
                                                rows_list.sort(key=lambda x : x[10])
                                                ws.delete_rows(4, ws.max_row)
                                                for row in rows_list:
                                                        ws.append(row)
                                                wb.save(file)
                                                flag = True
                                                self.destroy()
                                                messagebox.showinfo('Done','تم التسجيل بنجاح ')
                                                break                                                
                                if not flag:
                                        finalAmount = ws.cell(column=5,row=int(ws.max_row)).value
                                        ws['C1'] = int(ws.cell(column=3, row=1).value) + 1
                                        ws.append([c_date.get_date().strftime("%d/%m/%Y"),
                                        addService.get(),addCost.get(),
                                        addAmount.get(),
                                        ((int(finalAmount) + addCost.get()) - addAmount.get()),
                                        addComment.get(),
                                        ws.cell(column=3, row=1).value, c_date.get_date().year,c_date.get_date().month, c_date.get_date().day,date(c_date.get_date().year,c_date.get_date().month, c_date.get_date().day),])
                                        wb.save(file)
                                        flag = True
                                        self.destroy()
                                        messagebox.showinfo('Done','تم التسجيل بنجاح ')
                                        break
                if not flag:
                        self.destroy()
                        messagebox.showerror('Not Exists!','الاسم غير موجود')         
        
        Button(self, height = 1, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 20 bold', bd = '8px solid #DBA531', 
                text='اضافه',
                command=add).place(relx=.5, rely=.8,anchor= CENTER)
        
        
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
                if sheet.cell(row=1,column=1).value != None and sheet.title != 'مصروفات اداريه':
                        entries.append(f'{str(sheet.cell(row=1,column=1).value)} ({str(sheet.cell(row=1,column=2).value)})')

        if len(entries) < 1:
                entries.append('لا يوجد عملاء')   

        def check_input(event):
                value = event.widget.get()
                if value == '':
                        menu['values'] = entries
                else:
                        data = []
                        for item in entries:
                                if value.lower() in item.lower():
                                        data.append(item)

                        menu['values'] = data

        deleteName = StringVar()
        menu = ttk.Combobox(self, textvariable= deleteName, font=('Helvetica', 20,'bold'),values = entries )
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu.place(relx=.5, rely=.25,anchor= CENTER, width=600)
        menu['values'] = entries
        menu.bind('<KeyRelease>', check_input)

        def delete():
                flag = False
                for sheet in wb.worksheets:
                        ws = sheet
                        name = f'{str(ws.cell(row=1,column=1).value)} ({str(ws.cell(row=1,column=2).value)})'
                        if  name == deleteName.get() or sheet.title == deleteName.get() :
                                wb.remove(sheet)
                                flag = True
                                self.destroy()
                                messagebox.showinfo('Done','تم الحذف بنجاح')
                                wb.save(file)
                                break
                if not flag:
                        self.destroy()
                        messagebox.showerror('Not Exists!','الاسم غير موجود')         



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

                
        Button(self, height = 2, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text=' عميل جديد',
                command=self.open_insert).place(relx=.8, rely=.35,anchor= CENTER)

        Button(self, height = 2, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text=' خدمه جديده ',
                command=self.open_add).place(relx=.8, rely=.6,anchor= CENTER) 

        Button(self, height = 2, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text='مدفوعات العميل',
                command=self.open_payment).place(relx=.8, rely=.85,anchor= CENTER)

        Button(self, height = 2, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text='العملاء',
                command=self.open_search).place(relx=.5, rely=.45,anchor= CENTER)

        Button(self, height = 2, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text='حذف',
                command=self.open_delete).place(relx=.5, rely=.75,anchor= CENTER)

        Button(self, height = 2, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text='الايرادات',
                command=self.open_revenue).place(relx=.2, rely=.35,anchor= CENTER)       

        Button(self, height = 2, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text='مصروفات جديده',
                command=self.open_exp_form).place(relx=.2, rely=.85,anchor= CENTER)        


        Button(self, height = 2, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text= 'المصروفات',
                command=self.open_exp).place(relx=.2, rely=.6,anchor= CENTER)        



    def open_exp(self):
                window = Expenses(self)
                window.grab_set() 

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
                
    def open_exp_form(self):
                window = Expenses_form(self)
                window.grab_set()           



if __name__ == "__main__":
        app = Main()
        app.mainloop()
