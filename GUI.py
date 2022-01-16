# Change excel file name 
# Change Password

from tkinter import *
from tkinter import Entry, Label, Tk, ttk, messagebox
import tkinter as tk
from urllib import response
from openpyxl import load_workbook
from datetime import date, datetime
from tkcalendar import DateEntry

root = tk.Tk()
width = root.winfo_screenwidth()
height = root.winfo_screenheight()

root.destroy()

file = 'xl.xlsx' 
file0 = 'xl0.xlsx'
try:
  wb = load_workbook(file)
  wb0 = load_workbook(file0)
except:
  messagebox.showerror('File Not Found!','غير موجود xl.xlsx  ملف الاكسيل')


services = ['اتعاب لجنه داخلية','اضافة سياره','اعداد و مراجعة ميزانيه','اقرار ضرائب عامه'
        ,'اقرار ضرائب قيمه مضافه', 'اقرار ضرائب مرتبات', 'بطاقه ضريبيه'
        ,'تجديد اشتراك البوابه الالكترونيه', 'تحت الحساب','تسوية ملف ضريبى', 'تعديل النشاط',
        'جواب مرور', 'حفظ الملف بالضرائب','رسوم استخراج مستوردين', 'رسوم البوابه الالكترونيه'
        ,'رسوم تجديد مستوردين', 'سجل تجارى','سداد ضرائب عامه'
        , 'سداد ضرائب قيمه مضافه','شطب سجل تجارى', 'شهاده بالموقف الضريبى'
        ,'شهادة دخل 1', 'شهادة دخل 2', 'عمل موقع الكترونى', 'غرفه تجاريه',
        'فحص ضرائب عامه', 'فحص ضريبة قيمه مضافه','لجنة طعن ضرائب عامه', 'مركز مالى',
        'ميزانيه عموميه', 'نماذج 41 ض',]


class Expenses_View(tk.Toplevel):
      def __init__(self, parent):
        super().__init__(parent)

        self.geometry(f'{width}x{height}')
        self.title('المصروفات')
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
        
        

        def viewExpenses():
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
                treev.column("2", width = 1200,anchor ='e')
                treev.heading("1", text ="المصروف")
                treev.heading("2", text ="الحساب")

                expenses = []
                try:
                        # if Expenses worksheet exists :
                        ws0 = wb0['مصروفات اداريه']
                        expense = ''
                        for i in range (1, ws0.max_row+1):
                                # check row
                                try:
                                        celldate = date(
                                        ws0.cell(row=i, column=4).value,
                                        ws0.cell(row=i, column=5).value,
                                        ws0.cell(row=i, column=6).value)
                                except:
                                        continue
                                if celldate >= startdate.get_date() and celldate <= enddate.get_date():
                                        if ws0.cell(row=i,column=3).value not in expenses :
                                                expenses.append(ws0.cell(row=i,column=3).value)
                                                expense = str(ws0.cell(row=i,column=3).value)
                                                amount = 0 
                                                for i in range (1, ws0.max_row+1):
                                                        if str(ws0.cell(row=i,column=3).value) == expense:
                                                                amount += int(ws0.cell(row=i,column=2).value)
                                                treev.insert("", 'end', text ="L7",
                                                        values =(amount,expense), tags = ("expense",))
                                                treev.tag_configure('expense', background='#e1dddd', font=('Helvetica', 26))
                except:
                        self.destroy()
                        messagebox.showinfo('Not Found!','لا يوجد مصروفات اداريه حتى الان')        
                
        Button(self, height = 1, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 20 bold', bd = '8px solid #DBA531', 
                text='المصروفات',
                command=viewExpenses).place(relx=.5, rely=.55,anchor= CENTER)                                

        Button(self, height = 1, width = 10, bg = 'grey', fg = 'white',
        activebackground='#43516C', font = 'fantasy 15 bold', bd = '8px solid #DBA531', 
                text='اغلاق',
                command=self.destroy).place(relx=.5, rely=.9,anchor= CENTER)

# -----------------------------------------------------------------------------------------------------------------------------

class Expenses_Form(tk.Toplevel):
      def __init__(self, parent):
        super().__init__(parent)

        self.geometry(f'{width}x{height}')
        self.title('المصروفات')
        self['bg']='#E5E8C7'
        
        entries = []
        for sheet in wb.worksheets:
                if sheet.cell(row=1,column=1).value != None and sheet.title != 'مصروفات اداريه' :
                        entries.append(f'{str(sheet.cell(row=1,column=1).value)} ({str(sheet.cell(row=1,column=2).value)})')                
        if len(entries) < 1:
                entries.append('لا يوجد عملاء')   

        Label(self, text='اسم العميل',
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

        name = StringVar()
        menu = ttk.Combobox(self, textvariable= name, font=('Helvetica', 20,'bold'),values = entries )
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu.place(relx=.5, rely=.1,anchor= CENTER, width=600)
        menu['values'] = entries
        menu.bind('<KeyRelease>', check_input)
        Label(self, text='.اترك خانة (اسم العميل) فارغه اذا كان المصروف غير متعلق بعميل معين',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 12').place(relx=.577, rely=.15,anchor= CENTER)


        Label(self, text='المبلغ',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.75, rely=.25,anchor= CENTER, width=200)

        amount = IntVar()
        entry1 = ttk.Entry(self, textvariable = amount, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.5, rely=.25,anchor= CENTER, width=600, height=50)
        Label(self, text='.برجاء عدم ترك الخانه فارغه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 12').place(relx=.64, rely=.3,anchor= CENTER)
        

        Label(self, text='المصروف',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.75, rely=.4,anchor= CENTER, width=200)

        comment = StringVar()
        entry1 = ttk.Entry(self, textvariable=comment, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.5, rely=.4,anchor= CENTER, width=600, height=50)
        Label(self, text='.برجاء عدم ترك الخانه فارغه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 12').place(relx=.64, rely=.45,anchor= CENTER)


        Label(self, text='التاريخ',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.75, rely=.55,anchor= CENTER, width=200)

        c_date = DateEntry(self,width=30,bg="darkblue",fg="white",year=datetime.now().year,
        month=datetime.now().month,day=datetime.now().day
        ,locale='en_US', date_pattern='dd/MM/yyyy',font = ('fantasy', 25, 'bold'))
        c_date.place(relx=.625, rely=.55,anchor= CENTER, width=200, height=50)
        
        

                
        def addExpenses():
                try:
                        amount.get()
                except:
                        messagebox.showerror('Invalid!','من فضلك قم بادخال الخانات بطريقه صحيحه ') 
                        self.destroy()
                        return      
                if amount.get() == '':
                        messagebox.showerror('Invalid!','من فضلك قم بادخال الخانات المطلوبه   ')
                        self.destroy()
                        return   

                if name.get() == '':
                        try:
                                ws = wb0['مصروفات اداريه']
                        except:
                                ws = wb0.create_sheet('مصروفات اداريه') 
                                ws.append(['التاريخ', 'المبلغ','المصروف'])

                        inserted_date = date(c_date.get_date().year,c_date.get_date().month, c_date.get_date().day)
                        # check if transaction needs to be sorted by comparing date
                        exists_dates = []
                        for i in range (2, ws.max_row+1):
                                try:
                                        exists_dates.append(date(
                                        ws.cell(column=4, row=i).value,
                                        ws.cell(column=5, row=i).value,
                                        ws.cell(column=6, row=i).value,))
                                except:
                                        continue

                                for cell_date in exists_dates:
                                        if inserted_date < cell_date:
                                # add all transactions to rows_list
                                                rows_list = []
                                                for i in range (2, ws.max_row+1):
                                                        try:
                                                                cur_row_date = date(
                                                                ws.cell(column=4, row=i).value,
                                                                ws.cell(column=5, row=i).value,
                                                                ws.cell(column=6, row=i).value,)
                                                        except:
                                                                continue        
                                                        cur_row = [ws.cell(column=1, row=i).value,
                                                        ws.cell(column=2, row=i).value,
                                                        ws.cell(column=3, row=i).value,
                                                        ws.cell(column=4, row=i).value,
                                                        ws.cell(column=5, row=i).value,
                                                        ws.cell(column=6, row=i).value,
                                                        cur_row_date] 
                                                        rows_list.append(cur_row)
                                                
                                                rows_list.append([c_date.get_date().strftime("%d/%m/%Y")
                                                        ,amount.get()
                                                        ,comment.get()
                                                        ,c_date.get_date().year
                                                        ,c_date.get_date().month
                                                        ,c_date.get_date().day, inserted_date]) 
                                
                                                rows_list.sort(key=lambda x : x[6])
                                                # delete rows 
                                                ws.delete_rows(2, ws.max_row)
                                                #add sorted rows
                                                ws.append(rows_list[0])
                                                rows_list.pop(0)
                                                for row in rows_list:
                                                        ws.append(row)

                                                wb0.save(file0)
                                                self.destroy()
                                                messagebox.showinfo('Done','تم الحفظ بنجاح ')
                                                break

                                for cell_date in exists_dates:
                                        if not inserted_date < cell_date:
                                                ws.append([c_date.get_date().strftime("%d/%m/%Y")
                                                        ,amount.get()
                                                        ,comment.get()
                                                        ,c_date.get_date().year
                                                        ,c_date.get_date().month
                                                        ,c_date.get_date().day, inserted_date]) 
                                                wb0.save(file0)
                                                self.destroy()
                                                messagebox.showinfo('Done','تم الحفظ بنجاح ')  
                else:        
                        try:
                                ws = wb0[name.get()]
                        except:
                                messagebox.showerror('Not Exists!','الاسم غير موجود') 
                                self.destroy()

                        inserted_date = date(c_date.get_date().year,c_date.get_date().month, c_date.get_date().day)
                # check if transaction needs to be sorted by comparing date
                        exists_dates = []
                        for i in range (4, ws.max_row+1):
                                try:
                                        exists_dates.append(date(
                                        ws.cell(column=4, row=i).value,
                                        ws.cell(column=5, row=i).value,
                                        ws.cell(column=6, row=i).value,))
                                except:
                                        continue

                                for cell_date in exists_dates:
                                        if inserted_date < cell_date:
                                # add all transactions to rows_list
                                                rows_list = []
                                                for i in range (4, ws.max_row+1):
                                                        try:
                                                                cur_row_date = date(
                                                                ws.cell(column=4, row=i).value,
                                                                ws.cell(column=5, row=i).value,
                                                                ws.cell(column=6, row=i).value,)
                                                        except:
                                                                continue        
                                                        cur_row = [ws.cell(column=1, row=i).value,
                                                        ws.cell(column=2, row=i).value,
                                                        ws.cell(column=3, row=i).value,
                                                        ws.cell(column=4, row=i).value,
                                                        ws.cell(column=5, row=i).value,
                                                        ws.cell(column=6, row=i).value,
                                                        cur_row_date] 
                                                        rows_list.append(cur_row)
                                                
                                                rows_list.append([c_date.get_date().strftime("%d/%m/%Y")
                                                        ,amount.get()
                                                        ,comment.get()
                                                        ,c_date.get_date().year
                                                        ,c_date.get_date().month
                                                        ,c_date.get_date().day, inserted_date]) 
                                
                                                rows_list.sort(key=lambda x : x[6])
                                                # delete rows 
                                                ws.delete_rows(4, ws.max_row)
                                                #add sorted rows
                                                ws.append(rows_list[0])
                                                rows_list.pop(0)
                                                for row in rows_list:
                                                        ws.append(row)

                                                wb0.save(file0)
                                                self.destroy()
                                                messagebox.showinfo('Done','تم الحفظ بنجاح ')
                                                break
                                for cell_date in exists_dates:
                                        if not inserted_date < cell_date:
                                                ws.append([c_date.get_date().strftime("%d/%m/%Y")
                                                        ,amount.get()
                                                        ,comment.get()
                                                        ,c_date.get_date().year
                                                        ,c_date.get_date().month
                                                        ,c_date.get_date().day, inserted_date]) 
                                                wb0.save(file0)
                                                self.destroy()
                                                messagebox.showinfo('Done','تم الحفظ بنجاح ')  
                                                
                        

        Button(self, height = 1, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 20 bold', bd = '8px solid #DBA531', 
                text='حفظ',
                command=addExpenses).place(relx=.5, rely=.75,anchor= CENTER)

        Button(self, height = 1, width = 10, bg = 'grey', fg = 'white',
        activebackground='#43516C', font = 'fantasy 15 bold', bd = '8px solid #DBA531', 
                text='اغلاق',
                command=self.destroy).place(relx=.1, rely=.9,anchor= CENTER)        
        
# -----------------------------------------------------------------------------------------------------------------------------

class Revenues_View(tk.Toplevel):
      def __init__(self, parent):
        super().__init__(parent)

        self.geometry(f'{width}x{height}')
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
        
        def clientsRevenues():
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
        
                treev.column("1", width = 430,anchor ='e')
                treev.column("2", width = 150,anchor ='e')
                treev.column("3", width = 170,anchor ='e')
                treev.column("4", width = 670,anchor ='e')
                treev.column("5", width = 170,anchor ='e')
                
                treev.heading("1", text ="صافى الربح")
                treev.heading("2", text ="المصروف")
                treev.heading("3", text ="الايراد")
                treev.heading("4", text ="الاسم")
                treev.heading("5", text ="الكود")

                for client_sheet in wb.worksheets:
                        client_expenses = 0
                        client_revenue = 0
                        amount = 0
                        ws = client_sheet
                        if (ws.cell(row=1,column=2).value) != None :
                                for i in range(4, ws.max_row+1):
                                        try:
                                                celldate = date(
                                                        ws.cell(row=i, column=8).value,
                                                        ws.cell(row=i, column=9).value,
                                                        ws.cell(row=i, column=10).value)
                                        except:
                                                continue
                                        if celldate >= startdate.get_date() and celldate <= enddate.get_date():
                                                if ws.cell(row=i, column=4).value != None:
                                                        client_revenue += ws.cell(row=i, column=4).value

                        for expenses_sheet in wb0.worksheets:
                                if expenses_sheet.title == client_sheet.title: 
                                        ws = expenses_sheet
                                        for i in range(4, ws.max_row+1):
                                                try:
                                                        celldate = date(
                                                                ws.cell(row=i, column=4).value,
                                                                ws.cell(row=i, column=5).value,
                                                                ws.cell(row=i, column=6).value)
                                                except:
                                                        continue
                                                if celldate >= startdate.get_date() and celldate <= enddate.get_date():
                                                        if ws.cell(row=i, column=2).value != None:
                                                                client_expenses += ws.cell(row=i, column=2).value
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

        def totalRevenues():      
                total_revenue = 0
                total_expenses = 0
                income = 0
                for sheet in wb.worksheets:
                        ws = sheet
                        for i in range(4, ws.max_row+1):
                                try:
                                        celldate = date(
                                                ws.cell(row=i, column=8).value,
                                                ws.cell(row=i, column=9).value,
                                                ws.cell(row=i, column=10).value)
                                except:
                                        continue
                                if celldate >= startdate.get_date() and celldate <= enddate.get_date():
                                        try:
                                                total_revenue += ws.cell(row=i, column=4).value
                                        except:
                                                continue 
                                        
                for expenses_sheet in wb0.worksheets:
                        ws = expenses_sheet
                        for i in range(0, ws.max_row+1):
                                try:    
                                        celldate = date(
                                                ws.cell(row=i, column=4).value,
                                                ws.cell(row=i, column=5).value,
                                                ws.cell(row=i, column=6).value)
                                except:
                                        continue
                                if celldate >= startdate.get_date() and celldate <= enddate.get_date():
                                        try:
                                                total_expenses += ws.cell(row=i, column=2).value
                                        except:
                                                continue
                                                        
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
                command=totalRevenues).place(relx=.5, rely=.55,anchor= CENTER)

        Button(self, height = 2, width = 20, bg = '#D85426', fg = 'white',
        activebackground='#D85426', font = 'fantasy 24 bold', bd = '8px solid #DBA531', 
                text='تفاصيل الايرادات',
                command= clientsRevenues).place(relx=.5, rely=.2,anchor= CENTER)        

        Button(self, height = 1, width = 10, bg = 'grey', fg = 'white',
        activebackground='#43516C', font = 'fantasy 15 bold', bd = '8px solid #DBA531', 
                text='اغلاق',
                command=self.destroy).place(relx=.1, rely=.9,anchor= CENTER)        
        
# -----------------------------------------------------------------------------------------------------------------------------

class Payment_Form(tk.Toplevel):
      def __init__(self, parent):
        super().__init__(parent)

        self.geometry(f'{width}x{height}')
        self.title('المدفوعات')
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

        name = StringVar()
        menu = ttk.Combobox(self, textvariable= name, font=('Helvetica', 20,'bold'),values = entries )
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu.place(relx=.55, rely=.1,anchor= CENTER, width=600)
        menu['values'] = entries
        menu.bind('<KeyRelease>', check_input)
        Label(self, text='.برجاء عدم ترك الخانه فارغه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 12').place(relx=.69, rely=.15,anchor= CENTER)


        Label(self, text='المدفوع',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.8, rely=.25,anchor= CENTER, width=200)

        amount = IntVar()
        entry1 = ttk.Entry(self, textvariable = amount, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.55, rely=.25,anchor= CENTER, width=600, height=50)
        Label(self, text='.برجاء عدم ترك الخانه فارغه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 12').place(relx=.69, rely=.3,anchor= CENTER)
        

        Label(self, text='ملاحظات',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.8, rely=.4,anchor= CENTER, width=200)

        comment = StringVar()
        entry1 = ttk.Entry(self, textvariable=comment, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.55, rely=.4,anchor= CENTER, width=600, height=50)

        Label(self, text='التاريخ',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.8, rely=.55,anchor= CENTER, width=200)

        c_date = DateEntry(self,width=30,bg="darkblue",fg="white",year=datetime.now().year,
        month=datetime.now().month,day=datetime.now().day
        ,locale='en_US', date_pattern='dd/MM/yyyy',font = ('fantasy', 25, 'bold'))
        c_date.place(relx=.675, rely=.55,anchor= CENTER, width=200, height=50)
        
        
        def payment():
                try:
                        amount.get() and name.get()
                except:
                        messagebox.showerror('Invalid!','من فضلك قم بادخال الخانات بطريقه صحيحه ') 
                        self.destroy()
                        return      
                if amount.get() == '' or amount.get() == None :
                        messagebox.showerror('Invalid!','من فضلك قم بادخال خانة المدفوع بطريقه صحيحه ')
                        self.destroy()
                        return
                if name.get() == '':
                        messagebox.showerror('Invalid!','من فضلك قم بادخال خانة الاسم بطريقه صحيحه ')
                        self.destroy()
                        return        

                # for sheet in wb.worksheets:
                try:
                        ws = wb[name.get()]
                except:
                        messagebox.showerror('Not Exists!','الاسم غير موجود') 
                        self.destroy()

                finalAmount = ws.cell(column=5,row=int(ws.max_row)).value
                inserted_date = date(c_date.get_date().year,c_date.get_date().month, c_date.get_date().day)
                # check if transaction needs to be sorted by comparing date
                exists_dates = []
                for i in range (4, ws.max_row+1):
                        try:
                                exists_dates.append(date(
                                ws.cell(column=8, row=i).value,
                                ws.cell(column=9, row=i).value,
                                ws.cell(column=10, row=i).value,))
                        except:
                                continue
                # if transaction needs to be sorted
                for cell_date in exists_dates:
                        if inserted_date < cell_date:
                        # add all transactions to rows_list
                                rows_list = []
                                for i in range (4, ws.max_row+1):
                                        try:
                                                cur_row_date = date(
                                                        ws.cell(column=8, row=i).value,
                                                        ws.cell(column=9, row=i).value,
                                                        ws.cell(column=10, row=i).value,)
                                        except:
                                                continue        
                                        cur_row = [    
                                        ws.cell(column=1, row=i).value,
                                        ws.cell(column=2, row=i).value,
                                        ws.cell(column=3, row=i).value,
                                        ws.cell(column=4, row=i).value,
                                        ws.cell(column=5, row=i).value,
                                        ws.cell(column=6, row=i).value,
                                        ws.cell(column=7, row=i).value,
                                        ws.cell(column=8, row=i).value,
                                        ws.cell(column=9, row=i).value,
                                        ws.cell(column=10, row=i).value,
                                        cur_row_date]

                                        rows_list.append(cur_row)
                                
                                rows_list.append([c_date.get_date().strftime("%d/%m/%Y"),
                                        '-',0,
                                        amount.get(),
                                        finalAmount - amount.get(),
                                        comment.get(),
                                        ws.cell(column=3, row=1).value,
                                        c_date.get_date().year, c_date.get_date().month, c_date.get_date().day, inserted_date]) 
                                rows_list.sort(key=lambda x : x[10])
                                # delete rows 
                                ws.delete_rows(4, ws.max_row)
                                #add sorted rows
                                ws.append([rows_list[0][0],rows_list[0][1],rows_list[0][2],rows_list[0][3],int(rows_list[0][2])-int(rows_list[0][3]),rows_list[0][5],rows_list[0][6],rows_list[0][7],rows_list[0][8],rows_list[0][9],rows_list[0][10],])
                                
                                rows_list.pop(0)
                                for row in rows_list:
                                        try:
                                                ws.append([row[0],row[1],row[2],row[3],int(row[2])-int(row[3])+int(ws.cell(column=5,row=ws.max_row).value),row[5],row[6],row[7],row[8],row[9],row[10],])
                                        except:
                                                continue
                                wb.save(file)
                                self.destroy()
                                messagebox.showinfo('Done','تم الحفظ بنجاح ')
                                break

                for cell_date in exists_dates:
                        if not inserted_date < cell_date:
                                finalAmount = ws.cell(column=5,row=int(ws.max_row)).value
                                ws.append([c_date.get_date().strftime("%d/%m/%Y"),
                                        '-', 0,
                                        amount.get(),
                                        finalAmount - amount.get(), 
                                        comment.get(),
                                        ws.cell(column=3, row=1).value,
                                        c_date.get_date().year, c_date.get_date().month, c_date.get_date().day, inserted_date])
                                wb.save(file)
                                self.destroy()
                                messagebox.showinfo('Done','تم الحفظ بنجاح ')
                                
        Button(self, height = 1, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 20 bold', bd = '8px solid #DBA531', 
                text='حفظ',
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

        
        Label(self, text='اسم العميل',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.63, rely=.35,anchor= CENTER, width=200)

        name = StringVar()
        menu = ttk.Combobox(self, textvariable= name, font=('Helvetica', 20,'bold'),values = entries )
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu.place(relx=.5, rely=.4,anchor= CENTER, width=600)
        menu['values'] = entries
        menu.bind('<KeyRelease>', check_input)
        Label(self, text='.برجاء عدم ترك الخانه فارغه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 12').place(relx=.637, rely=.45,anchor= CENTER)


        flag = False
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

                try:
                        ws = wb[name.get()]
                except:
                        self.destroy()
                        messagebox.showerror('Not Exists!','الاسم غير موجود')                 
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
                        
                treev.tag_configure('odd', background='#e1dddd', font=('Helvetica', 26)) 
                treev.tag_configure('even', background= '#f5f3f3', font=('Helvetica', 26))
        self.title( name.get())
                


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
                treev.column("1", width = 480,anchor ='e')
                treev.column("2", width = 230,anchor ='e')
                treev.column("3", width = 140,anchor ='e')
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
                command= search).place(relx=.5, rely=.55,anchor= CENTER)

        Button(self, height = 2, width = 20, bg = '#D85426', fg = 'white',
        activebackground='#D85426', font = 'fantasy 24 bold', bd = '8px solid #DBA531', 
                text='العملاء',
                command= clients).place(relx=.5, rely=.2,anchor= CENTER)        
        
        Button(self, height = 1, width = 10, bg = 'grey', fg = 'white',
        activebackground='#43516C', font = 'fantasy 15 bold', bd = '8px solid #DBA531', 
                text='اغلاق',
                command=self.destroy).place(relx=.5, rely=.8,anchor= CENTER)

# --------------------------------------------------------------------------------------------------------------------------------------------

class Client_Form(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self.geometry(f'{width}x{height}')
        self.title('اضافة عميل')
        self['bg']='#E5E8C7'

        Label(self, text='الاسم',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.9, rely=.1,anchor= CENTER, width=200)

        name = tk.StringVar()
        entry1 = ttk.Entry(self, textvariable=name, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.825, rely=.17,anchor= CENTER, width=450)
        Label(self, text='.برجاء عدم ترك الخانه فارغه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 12').place(relx=.92, rely=.22,anchor= CENTER)

        Label(self, text='الكود',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.5, rely=.1,anchor= CENTER, width=200)

        code = IntVar()
        entry1 = ttk.Entry(self, textvariable=code, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.5, rely=.17,anchor= CENTER, width=200)
        Label(self, text='.برجاء عدم ترك الخانه فارغه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 12').place(relx=.51, rely=.22,anchor= CENTER)

        Label(self, text='رقم التسجيل',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.1, rely=.1,anchor= CENTER, width=200)

        record = tk.IntVar()
        entry1 = ttk.Entry(self, textvariable=record, justify = LEFT, font = ('fantasy', 25, 'bold'))
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

        service = tk.StringVar()
        menu2 = ttk.Combobox(self, textvariable = service, font=('Helvetica', 20,'bold'), values = services) 
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu2.place(relx=.835, rely=.37,anchor= CENTER, width=400)
        Label(self, text='.برجاء عدم ترك الخانه فارغه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 12').place(relx=.912, rely=.42,anchor= CENTER)
        

        Label(self, text='التكلفه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.5, rely=.3,anchor= CENTER, width=200)

        cost = IntVar()
        entry1 = ttk.Entry(self, textvariable=cost, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.5, rely=.37,anchor= CENTER, width=200)
        Label(self, text='.برجاء عدم ترك الخانه فارغه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 12').place(relx=.51, rely=.42,anchor= CENTER)


        Label(self, text='المدفوع',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.1, rely=.3,anchor= CENTER, width=200)

        amount = IntVar()
        entry1 = ttk.Entry(self, textvariable=amount, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.1, rely=.37,anchor= CENTER, width=200)
        Label(self, text='.برجاء عدم ترك الخانه فارغه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 12').place(relx=.11, rely=.42,anchor= CENTER)
        
        
        Label(self, text='ملاحظات',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.9, rely=.5,anchor= CENTER, width=200)

        comment = StringVar()
        entry1 = ttk.Entry(self, textvariable=comment, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.78, rely=.57,anchor= CENTER, width=600,)

        Label(self, text='التاريخ',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.375, rely=.5,anchor= CENTER, width=200)

        c_date = DateEntry(self,width=30,bg="darkblue",fg="white", font = ('fantasy', 25, 'bold'),
        year=datetime.now().year,month=datetime.now().month,day=datetime.now().day
        ,locale='en_US', date_pattern='dd/MM/yyyy')
        c_date.place(relx=.25, rely=.57,anchor= CENTER, width=600)
        
        
        
        entries = []
        for sheet in wb.worksheets:
                if sheet.cell(row=1,column=1).value != None and sheet.title != 'مصروفات اداريه':
                        entries.append(str(sheet.cell(row=1,column=1).value)) 

        codeentries = []
        for sheet in wb.worksheets:
                if sheet.cell(row=1,column=2).value != None :
                        codeentries.append(sheet.cell(row=1,column=2).value)                

        def insertClient():
                # get() validation
                try:
                        name.get() and code.get() and record.get() and service.get() and cost.get() and amount.get() and phone.get() and address.get() and comment.get()
                except:
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال الخانات بطريقه صحيحه ') 
                        return

                # not None Validation
                if name.get() == None or code.get() == None or service.get() == None or cost.get() == None or amount.get() == None :
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال الخانات المطلوبه ') 
                        return        

                # not "" validation        
                if name.get() == "" or service.get() == "" :
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك نأكد من ادخال الاسم و الخدمه ') 
                        return

                # String validation  
                if not name.get().isalpha() and service.get().isalpha():
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك نأكد من ادخال الاسم و الخدمه بطريقه صحيحه ') 
                        return   
                
                # phone validation
                if not phone.get().isnumeric:
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك نأكد من ادخال التليفون بطريقه صحيحه ') 
                        return  

                # Case validations
                if name.get() in entries or code.get() in codeentries:
                        self.destroy()
                        messagebox.showerror('Already Exists!','الاسم او الكود موجودين بالفعل ') 
                        return
                if len(name.get()) > 20:
                        self.destroy()
                        messagebox.showerror('Error!','من فضلك استخدم اسم لا تزيد حروفه عن 20 حرف') 
                        return
                
                else:        
                        ws = wb.create_sheet(f'{name.get()} ({code.get()})')
                        ws.title = f'{name.get()} ({code.get()})'
                        ws.append([name.get(), code.get(), 1,str(record.get()),phone.get(),address.get()])
                        ws.append([])
                        ws.append(['التاريخ', 'الخدمه','التكلفه','المدفوع','الرصيد', 'الملاحظات', 'مسلسل',])
                        ws.append([c_date.get_date().strftime("%d/%m/%Y"),
                        service.get(),
                        cost.get(),
                        amount.get(), 
                        (cost.get()) - amount.get() ,
                        comment.get(), 1,
                        c_date.get_date().year,c_date.get_date().month, c_date.get_date().day])
                        wb.save(file)

                        wb0 = load_workbook(filename='xl0.xlsx')
                        ws0 = wb0.create_sheet(f'{name.get()} ({code.get()})')
                        ws0.title = f'{name.get()} ({code.get()})'
                        ws0.append([name.get(), code.get(), 1,str(record.get()),phone.get(),address.get()])
                        ws0.append([])
                        ws0.append(['التاريخ', 'المبلغ','المصروف'])
                        wb0.save(filename='xl0.xlsx')
                        # wb0.close()

                        self.destroy()
                        messagebox.showinfo('Done','تم الحفظ بنجاح ')

        
        Button(self, height = 1, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 20 bold', bd = '8px solid #DBA531', 
                text='حفظ',
                command=insertClient).place(relx=.5, rely=.85,anchor= CENTER)
        
        
        Button(self, height = 1, width = 10, bg = 'grey', fg = 'white',
        activebackground='#43516C', font = 'fantasy 15 bold', bd = '8px solid #DBA531', 
                text='اغلاق',
                command=self.destroy).place(relx=.1, rely=.9,anchor= CENTER)

# -------------------------------------------------------------------------------------------------------

class Service_Form(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self.geometry(f'{width}x{height}')
        self.title('اضافة خدمه')
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

        name = StringVar()
        menu = ttk.Combobox(self, textvariable= name, font=('Helvetica', 20,'bold'),values = entries )
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu.place(relx=.55, rely=.1,anchor= CENTER, width=600)
        menu['values'] = entries
        menu.bind('<KeyRelease>', check_input)
        Label(self, text='.برجاء عدم ترك الخانه فارغه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 12').place(relx=.69, rely=.15,anchor= CENTER)


        Label(self, text='الخدمه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.8, rely=.22,anchor= CENTER, width=200)

        service = tk.StringVar()
        menu2 = ttk.Combobox(self, textvariable = service, font=('Helvetica', 20,'bold'), values = services)
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu2.place(relx=.55, rely=.22,anchor= CENTER, width=600)
        Label(self, text='.برجاء عدم ترك الخانه فارغه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 12').place(relx=.69, rely=.27,anchor= CENTER)
        

        Label(self, text='التكلفه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.8, rely=.34,anchor= CENTER, width=200)

        cost = IntVar()
        entry1 = ttk.Entry(self, textvariable=cost, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.55, rely=.34,anchor= CENTER, width=600, height=50)
        Label(self, text='.برجاء عدم ترك الخانه فارغه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 12').place(relx=.69, rely=.39,anchor= CENTER)


        Label(self, text='المدفوع',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.8, rely=.58,anchor= CENTER, width=200)

        amount = IntVar()
        entry1 = ttk.Entry(self, textvariable=amount, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.55, rely=.58,anchor= CENTER, width=600, height=50)
        Label(self, text='.برجاء عدم ترك الخانه فارغه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 12').place(relx=.69, rely=.63,anchor= CENTER)
        

        Label(self, text='ملاحظات',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.8, rely=.7,anchor= CENTER, width=200)

        comment = StringVar()
        entry1 = ttk.Entry(self, textvariable=comment, justify = LEFT, font = ('fantasy', 25, 'bold'))
        entry1.place(relx=.55, rely=.7,anchor= CENTER, width=600, height=50)


        Label(self, text='التاريخ',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.8, rely=.46,anchor= CENTER, width=200)

        c_date = DateEntry(self,width=30,bg="darkblue",fg="white", font = ('fantasy', 25, 'bold'),
        year=datetime.now().year,month=datetime.now().month,day=datetime.now().day
        ,locale='en_US', date_pattern='dd/MM/yyyy')
        c_date.place(relx=.675, rely=.46,anchor= CENTER, width=200, height=50)
        
        
        def addService():
                try:
                        amount.get(), cost.get(), service.get(), name.get() 
                except:
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال الخانات بطريقه صحيحه ')

                if cost.get() == '' or name.get() == '' or service.get() == '':
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال الخانات المطلوبه ')
                        return
                try:
                        cost.get() 
                except:
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال خانة التكلفه بطريقه صحيحه ')
                        return

                try:
                        service.get() 
                except:
                        self.destroy()
                        messagebox.showerror('Invalid!','من فضلك قم بادخال خانة الخدمه بطريقه صحيحه ')  
                        return

                finalAmount = 0
                try:
                        ws = wb[name.get()]
                except:
                        self.destroy()
                        messagebox.showerror('Not Exists!','الاسم غير موجود')         

                inserted_date = date(c_date.get_date().year,c_date.get_date().month, c_date.get_date().day)
                exists_dates = []
                        # check if transaction needs to be sorted by comparing date
                for i in range (4, ws.max_row+1):
                        try:
                                exists_dates.append (date(
                                ws.cell(column=8, row=i).value,
                                ws.cell(column=9, row=i).value,
                                ws.cell(column=10, row=i).value,))
                        except:
                                continue
                # if transaction needs to be sorted
                for cell_date in exists_dates:
                        if inserted_date < cell_date:
                                # add all transactions to rows_list
                                rows_list = []
                                for i in range (4, ws.max_row+1):
                                        try:
                                                cur_row_date = date(
                                                        ws.cell(column=8, row=i).value,
                                                        ws.cell(column=9, row=i).value,
                                                        ws.cell(column=10, row=i).value,)
                                        except:
                                                continue        
                                        cur_row = [    
                                        ws.cell(column=1, row=i).value,
                                        ws.cell(column=2, row=i).value,
                                        ws.cell(column=3, row=i).value,
                                        ws.cell(column=4, row=i).value,
                                        ws.cell(column=5, row=i).value,
                                        ws.cell(column=6, row=i).value,
                                        ws.cell(column=7, row=i).value,
                                        ws.cell(column=8, row=i).value,
                                        ws.cell(column=9, row=i).value,
                                        ws.cell(column=10, row=i).value,
                                        cur_row_date]

                                        rows_list.append(cur_row)
                                # add the new transaction
                                rows_list.append([
                                c_date.get_date().strftime("%d/%m/%Y"),
                                service.get(),
                                cost.get(),
                                amount.get(),
                                (cost.get() - amount.get()),
                                comment.get(),
                                ws.cell(column=3, row=1).value,
                                c_date.get_date().year,c_date.get_date().month, c_date.get_date().day,
                                inserted_date,])  

                                # sort rows_list by date
                                rows_list.sort(key=lambda x : x[10])
                                # delete rows 
                                ws.delete_rows(4, ws.max_row)
                                #add sorted rows
                                
                                ws.append([rows_list[0][0],rows_list[0][1],rows_list[0][2],rows_list[0][3],int(rows_list[0][2])-int(rows_list[0][3]),rows_list[0][5],rows_list[0][6],rows_list[0][7],rows_list[0][8],rows_list[0][9],rows_list[0][10],])
                                
                                rows_list.pop(0)
                                for row in rows_list:
                                        try:
                                                ws.append([row[0],row[1],row[2],row[3],row[2]-row[3]+int(ws.cell(column=5,row=ws.max_row).value),row[5],row[6],row[7],row[8],row[9],row[10],])
                                        except:
                                                continue
                                wb.save(file)
                                self.destroy()
                                messagebox.showinfo('Done','تم الحفظ بنجاح ')
                                break 

                # if transaction doesn't need to be sorted
                for cell_date in exists_dates:
                        if not inserted_date < cell_date:
                                finalAmount = ws.cell(column=5,row=int(ws.max_row)).value  
                                ws.append([c_date.get_date().strftime("%d/%m/%Y"),
                                service.get(),
                                cost.get(),
                                amount.get(),
                                ((int(finalAmount) + (cost.get()) - amount.get())),
                                comment.get(),
                                ws.cell(column=3, row=1).value,
                                c_date.get_date().year,c_date.get_date().month, c_date.get_date().day,
                                inserted_date,])

                                wb.save(file)
                                self.destroy()
                                messagebox.showinfo('Done','تم الحفظ بنجاح ')

        
        Button(self, height = 1, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 20 bold', bd = '8px solid #DBA531', 
                text='حفظ',
                command=addService).place(relx=.5, rely=.8,anchor= CENTER)
        
        Button(self, height = 1, width = 10, bg = 'grey', fg = 'white',
        activebackground='#43516C', font = 'fantasy 15 bold', bd = '8px solid #DBA531', 
                text='اغلاق',
                command=self.destroy).place(relx=.1, rely=.9,anchor= CENTER)

# -----------------------------------------------------------------------------------------------------------------------------

class Delete(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self.geometry(f'{width}x{height}')
        self.title('حذف عميل')
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

        Label(self, text='اسم العميل',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 30 bold').place(relx=.63, rely=.3,anchor= CENTER)

        name = StringVar()
        menu = ttk.Combobox(self, textvariable= name, font=('Helvetica', 20,'bold'),values = entries )
        text_font = ('Courier New', '20', 'bold')
        app.option_add('*TCombobox*Listbox.font', text_font)
        menu.place(relx=.5, rely=.35,anchor= CENTER, width=600)
        menu['values'] = entries
        menu.bind('<KeyRelease>', check_input)
        Label(self, text='.برجاء عدم ترك الخانه فارغه',
        bg = '#E5E8C7', fg = 'black' ,font = 'fantasy 12').place(relx=.64, rely=.4,anchor= CENTER)

        def delete():
                response = messagebox.askquestion('Deleted!', f'{name.get()} هل تريد ان تحذف كشف حساب العميل')
                if response == 'yes' :
                        try:
                                del wb[name.get()]
                        except:
                                self.destroy()
                                messagebox.showerror('Not Exists!','الاسم غير موجود')         
                        self.destroy()
                        messagebox.showinfo('Done!','تم الحذف بنجاح')
                        wb.save(file)
                else :
                        self.destroy()
                        messagebox.showinfo('Fail!', 'لم يتم الحذف')       


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

        self.geometry(f'{width}x{height}')
        self.title('Main Window')
        self['bg']='#E5E8C7'

        Label(self, text='Mr.Wagdy Latif For Accounting Services',
        bg = '#D85426', fg = 'white', font = 'fantasy 30 bold', borderwidth=20, relief="ridge", padx=20, pady=40).place(relx=.5, rely=.1,anchor= CENTER)
                
        Button(self, height = 1, width = 13, bg = 'green', fg = 'white',
        activebackground='green', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text='عملاء',
                command=self.open_insert).place(relx=.8, rely=.35,anchor= CENTER)

        Button(self, height = 1, width = 13, bg = 'green', fg = 'white',
        activebackground='green', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text='خدمات',
                command=self.open_add).place(relx=.8, rely=.5,anchor= CENTER) 

        Button(self, height = 1, width = 13, bg = 'green', fg = 'white',
        activebackground='green', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text='مدفوعات',
                command=self.open_payment).place(relx=.8, rely=.65,anchor= CENTER)

        Button(self, height = 1, width = 13, bg = 'green', fg = 'white',
        activebackground='green', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text='مصروفات',
                command=self.open_exp_form).place(relx=.8, rely=.8,anchor= CENTER) 

        Button(self, height = 2, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text='العملاء',
                command=self.open_search).place(relx=.5, rely=.35,anchor= CENTER)

        Button(self, height = 2, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text='الإيرادات',
                command=self.open_revenue).place(relx=.5, rely=.6,anchor= CENTER)       

        Button(self, height = 2, width = 15, bg = '#05659E', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text= 'المصروفات',
                command=self.open_exp).place(relx=.5, rely=.85,anchor= CENTER)        

        Button(self, height = 1, width = 13, bg = 'red', fg = 'white',
        activebackground='#43516C', font = 'fantasy 30 bold', bd = '8px solid #DBA531', 
                text='حذف',
                command=self.open_delete).place(relx=.2, rely=.6,anchor= CENTER)


    def open_exp(self):
                window = Expenses_View(self)
                window.grab_set() 

    def open_revenue(self):
                window = Revenues_View(self)
                window.grab_set()  

    def open_payment(self):
                window = Payment_Form(self)
                window.grab_set()            

    def open_search(self):
                window = Search(self)
                window.grab_set()

    def open_insert(self):
                window = Client_Form(self)
                window.grab_set()

    def open_delete(self):
                window = Delete(self)
                window.grab_set()  

    def open_add(self):
                window = Service_Form(self)
                window.grab_set()
                
    def open_exp_form(self):
                window = Expenses_Form(self)
                window.grab_set()           

if __name__ == "__main__":
        app = Main()
        app.mainloop()