if inserted_date in exists_dates:
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
                                finalAmount - amount.get(), #<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                comment.get(),
                                ws.cell(column=3, row=1).value,
                                c_date.get_date().year, c_date.get_date().month, c_date.get_date().day, inserted_date]) 
                        print(rows_list)        
                        rows_list.sort(key=lambda x : x[10])
                        print(rows_list)        
                        # delete rows 
                        ws.delete_rows(4, ws.max_row)
                        #add sorted rows
                        ws.append(rows_list[0])
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

                else:
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
                        break
        # if not flag:
        #         messagebox.showerror('Not Exists!','الاسم غير موجود') 
        #         self.destroy()
        