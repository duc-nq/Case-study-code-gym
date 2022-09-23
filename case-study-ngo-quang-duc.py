# %%
import tkinter as tk
from tkinter import ttk, messagebox
import mysql.connector
from sqlalchemy import create_engine
from tkinter import *
import urllib.parse
from sqlalchemy import column, null
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
from tkinter.messagebox import showerror
from tkinter.filedialog import asksaveasfile
import pandas as pd
from tkinter import filedialog
from PIL import Image, ImageTk
from tkinter import Toplevel, Button, Tk, Menu  

def GetValue(event):
    e1.delete(0, END)
    e2.delete(0, END)
    e3.delete(0, END)
    e4.delete(0, END)
    row_id = listBox.selection()[0]
    select = listBox.set(row_id)
    e1.insert(0,select['id'])
    e2.insert(0,select['empname'])
    e3.insert(0,select['mobile'])
    e4.insert(0,select['salary'])
 
 
def Add():
    studid = e1.get()
    studname = e2.get()

    coursename = e3.get()
    feee = e4.get()
 
    mysqldb=mysql.connector.connect(host="103.82.21.72",user="codegym",password="Datcang1!",database="codegym")
    mycursor=mysqldb.cursor()
 
    try:
       sql = "INSERT INTO  registation (id,empname,mobile,salary) VALUES (%s, %s, %s, %s)"
       val = (studid,studname,coursename,feee)
       mycursor.execute(sql, val)
       mysqldb.commit()
       lastid = mycursor.lastrowid
       messagebox.showinfo("information", "Thông tin nhân viên đã được thêm mới")
       e1.delete(0, END)
       e2.delete(0, END)
       e3.delete(0, END)
       e4.delete(0, END)
       e1.focus_set()
    except Exception as e:
       print(e)
       mysqldb.rollback()
       mysqldb.close()
    show()
 
 
def update():
    studid = e1.get()
    studname = e2.get()
    coursename = e3.get()
    feee = e4.get()
    mysqldb=mysql.connector.connect(host="103.82.21.72",user="codegym",password="Datcang1!",database="codegym")
    mycursor=mysqldb.cursor()
 
    try:
        mycursor.execute(f"SELECT * FROM data_kiot_viet.registation where id={studid}")
        records = mycursor.fetchall()
        if len(records) == 0:
            messagebox.showinfo("Information", "Mã nhân viên không tồn tại")
        else:
            sql = "Update registation set empname= %s,mobile= %s,salary= %s where id= %s"
            val = (studname,coursename,feee,studid)
            mycursor.execute(sql, val)
            mysqldb.commit()
            lastid = mycursor.lastrowid
            messagebox.showinfo("information", "Thông tin nhân viên đã được cập nhật")
            e1.delete(0, END)
            e2.delete(0, END)
            e3.delete(0, END)
            e4.delete(0, END)
            e1.focus_set()
    except Exception as e:
       print(e)
       mysqldb.rollback()
       mysqldb.close()
    show()
 
def delete():
    studid = e1.get()
    mysqldb=mysql.connector.connect(host="103.82.21.72",user="codegym",password="Datcang1!",database="codegym")
    mycursor=mysqldb.cursor()
 
    try:
            mycursor.execute(f"SELECT * FROM data_kiot_viet.registation where id={studid}")
            records = mycursor.fetchall()
            if len(records) == 0:
                messagebox.showinfo("Information", "Mã nhân viên không tồn tại")
            else:  
                sql = "delete from registation where id = %s"
                val = (studid,)
                mycursor.execute(sql, val)
                mysqldb.commit()
                lastid = mycursor.lastrowid
                messagebox.showinfo("Information", "Thông tin nhân viên đã được xóa")
                e1.delete(0, END)
                e2.delete(0, END)
                e3.delete(0, END)
                e4.delete(0, END)
                e1.focus_set()
    except Exception as e:
        print(e)
        mysqldb.rollback()
        mysqldb.close()
    show()
 
def show():
        remove_many()
        mysqldb = mysql.connector.connect(host="103.82.21.72",user="codegym",password="Datcang1!",database="codegym")
        mycursor = mysqldb.cursor()
        mycursor.execute("SELECT id,empname,mobile,salary FROM registation")
        records = mycursor.fetchall()
        print(records)

        for i, (id,stname, course,fee) in enumerate(records, start=1):
            listBox.insert("", "end", values=(id, stname, course, fee))
            mysqldb.close()

def remove_many():
    listBox.delete(*listBox.get_children())

def find():
    remove_many()
    studid = e1.get()
    mysqldb=mysql.connector.connect(host="103.82.21.72",user="codegym",password="Datcang1!",database="codegym")
    mycursor=mysqldb.cursor()
    
    try:
       mycursor.execute(f"SELECT * FROM data_kiot_viet.registation where id={studid}")
       records = mycursor.fetchall()
       #messagebox.showinfo("Information", "Tìm kiếm thành công")
       if len(records) == 0:
           messagebox.showinfo("Information", "Mã nhân viên không tồn tại")
       else:
           print(records)
           for i, (id,stname, course,fee) in enumerate(records, start=1):
                    listBox.insert("", "end", values=(id, stname, course, fee))
                    mysqldb.close()
 
    except Exception as e:
       mysqldb.rollback()
       mysqldb.close()


def save():
    file = filedialog.asksaveasfilename(defaultextension=".xlsx")
    engine = create_engine('mysql+mysqldb://codegym:%s@103.82.21.72:3306/codegym' % urllib.parse.quote('Datcang1!'), echo = False)
    df = pd.read_sql("SELECT * FROM data_kiot_viet.registation;", con=engine)
    df.to_excel(file,index=False)

def import_file():
    file = filedialog.askopenfilename(defaultextension=".xlsx")
    engine = create_engine('mysql+mysqldb://codegym:%s@103.82.21.72:3306/codegym' % urllib.parse.quote('Datcang1!'), echo = False)
    df = ['id','empname','mobile','salary']
    df = pd.read_excel(file)
    id = df['id'].values.tolist()
    if len(id) == 0:
        messagebox.showinfo("Information", "File nhập dữ liệu không chứa mã nhân viên!")
    else:
        for i in id:
            try:
                delete_id = pd.read_sql(f"delete from data_kiot_viet.registation where id={i};", con=engine)
            except:
                continue
        df.to_sql(name = 'registation', con = engine, if_exists = 'append', index = False)
        messagebox.showinfo("Information", "Nhập dữ liệu thành công")
        show()

def about():
    messagebox.showinfo("Information", "HRM v1.0 - Ngô Quang Đức - Học viên CodeGym")

root = Tk()
root.geometry("840x450")
root.title('HRM v1.0')
icon_master = PhotoImage(file = r'C:\Users\Admin\New folder\selection.png')
root.iconphoto(False, icon_master)

#add the menu to the menubar
menubar = Menu(root)
root.config(menu=menubar)
file_menu = Menu(menubar,tearoff=0)

file_menu.add_command(
    label='Exit',
    command=root.destroy,
)
menubar.add_cascade(
    label="Hệ Thống",
    menu=file_menu,
    underline=0
)
help_menu = Menu(
    menubar,
    tearoff=0
)

help_menu.add_command(label='Về chúng tôi',command=about)

# add the Help menu to the menubar
menubar.add_cascade(
    label="Hỗ trợ",
    menu=help_menu,
    underline=0
)
my_img = tk.PhotoImage(file = r'C:\Users\Admin\New folder\hiring.png')
l2 = tk.Label(root, image=my_img )
l2.place(x=700, y=15)

global e1
global e2
global e3
global e4
 
tk.Label(root, text="QUẢN LÝ NHÂN SỰ", fg='#ba3271', font=(None, 30)).place(x=280, y=35)
tk.Label(root, text="Mã nhân viên").place(x=10, y=10)
Label(root, text="Tên nhân viên").place(x=10, y=40)
Label(root, text="Số điện thoại").place(x=10, y=70)
Label(root, text="Mức lương").place(x=10, y=100)

e1 = Entry(root)
e1.place(x=140, y=10)
 
e2 = Entry(root)
e2.place(x=140, y=40)
 
e3 = Entry(root)
e3.place(x=140, y=70)
 
e4 = Entry(root)
e4.place(x=140, y=100)

my_font=('tahoma', 10, 'bold')
Button(root, text="Thêm mới",command = Add,height=3, width= 13, bg='#CB2777',fg='white',font=my_font).place(x=30, y=130)
Button(root, text="Cập nhật",command = update,height=3, width= 13, bg='#ED4B66',fg='white',font=my_font).place(x=140, y=130)
Button(root, text="Xóa",command = delete,height=3, width= 13, bg='#FF7557',fg='white',font=my_font).place(x=250, y=130)
Button(root, text="Tìm kiếm",command = find,height=3, width= 13, bg='#FFA14E',fg='white',font=my_font).place(x=360, y=130)
Button(root, text="Xuất dữ liệu",command = save,height=3, width= 13, bg='#FFCD54',fg='white',font=my_font).place(x=470, y=130)
Button(root, text="Nhập dữ liệu",command = import_file,height=3, width= 13, bg='#009690',fg='white',font=my_font).place(x=580, y=130)
cols = ('id', 'empname', 'mobile','salary')
listBox = ttk.Treeview(root, columns=cols, show='headings')

vsb = ttk.Scrollbar(root, orient="vertical", command=listBox.yview)
vsb.place(x=810, y=200, height=200+20)

listBox.configure(yscrollcommand=vsb.set)

listBox.heading('id', text='Mã nhân viên')
listBox.heading('empname', text='Tên nhân viên')
listBox.heading('mobile', text='Số điện thoại')
listBox.heading('salary', text='Mức lương')
show()
listBox.bind('<Double-Button-1>',GetValue)
for col in cols:
    #listBox.heading(col, text=col)
    listBox.grid(row=1, column=0, columnspan=2)
    listBox.place(x=10, y=200)


 
root.mainloop()

# %%





