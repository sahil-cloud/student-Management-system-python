import random
from tkinter import *
import time
import datetime
from tkinter import Toplevel, messagebox, ttk ,filedialog
from tkinter.ttk import Treeview
import pymysql
import pandas
import openpyxl

######### Functions Area Started ###########


# Label function
def label(place_, text, col, col2, x1, y1, w, h):
    l = Label(place_, text=text, font="luicida 15 bold", bg=col, fg=col2)
    l.place(x=x1, y=y1, width=w, height=h)


# Button function
def btn(place_, text, col, col2, abg, afg, x1, y1, w, h, cmd):
    b = Button(place_, text=text, font="luicida 15 bold", bg=col,
               activebackground=abg, activeforeground=afg, fg=col2, command=cmd)
    b.place(x=x1, y=y1, width=w, height=h)


# timing in the clock
def clock():
    now = time.strftime("%H:%M:%S")
    d = time.strftime("%d:%m:%y")
    # label(root, f"Date: {d} \nTime: {now}", "lightblue", "white", 10, 10, 200, 82)
    clk.config(text=f"Date: {d} \nTime: {now}")
    clk.after(200, clock)


# creating function for rotating text
def rotate1():
    global textvar, count, tex

    if count >= len(textvar):
        count = 0
        tex = ""
    else:
        tex = tex+textvar[count]
        head.config(text=tex)
        count = count+1
    head.after(200, rotate1)


# connect to database
def dbConnect():
    host = StringVar()
    user = StringVar()
    passw = StringVar()

    # submit data into database
    def submitDb():
        # getting the values
        global conn, mycursor

        host1 = host.get()
        user1 = user.get()
        passw1 = passw.get()
        # print(host1, passw1, user1)

        # connecting to the database
        try:
            conn = pymysql.connect(host=host1, user=user1, password=passw1)
            mycursor = conn.cursor()
        except:
            messagebox.showerror(
                "Warning!", "Data is incorrect..Please try again")
            return

        # creating db and making table
        try:
            q1 = 'create database studentManagementRecord'
            mycursor.execute(q1)
            q2 = 'use studentManagementRecord'
            mycursor.execute(q2)
            q3 = 'create table studentData (id int,name varchar(30),email varchar(40),mobile varchar(12),address varchar(80),dob varchar(30),gender varchar(10),date varchar(30),time varchar(30))'
            mycursor.execute(q3)
            q4 = 'alter table studentData modify column id int not null'
            mycursor.execute(q4)
            q5 = 'alter table studentData modify column id int primary key'
            mycursor.execute(q5)
            messagebox.showinfo(
                'Success', "Database created and connected to it successfully", parent=top)

        except:
            q = 'use studentManagementRecord'
            mycursor.execute(q)
            messagebox.showinfo(
                'Success', "Connected to database successfully", parent=top)
        top.destroy()

    top = Toplevel()
    top.title("Add to Database")
    top.geometry("400x270+400+150")
    top.grab_set()
    top.config(bg="lightblue")
    top.resizable(False, False)

    # creating labels
    label(top, "Enter Host: ", "lightblue", "black", 5, 5, 140, 50)
    e1 = Entry(top, textvariable=host, font="luicida 15 bold")
    e1.place(x=160, y=12, width=200, height=40)

    label(top, "Enter User: ", "lightblue", "black", 5, 60, 140, 50)
    e2 = Entry(top, textvariable=user, font="luicida 15 bold")
    e2.place(x=160, y=62, width=200, height=40)

    label(top, "Enter Password: ", "lightblue", "black", 5, 110, 180, 50)
    e3 = Entry(top, textvariable=passw, font="luicida 15 bold")
    e3.place(x=180, y=122, width=200, height=40)

    btn(top, "Submit", "Lightgreen", "black",
        "lightgreen", "white", 140, 190, 100, 60, submitDb)

    top.mainloop()


# left frame functions

def addStudent():
    id_ = StringVar()
    name = StringVar()
    email = StringVar()
    mobile = StringVar()
    gender = StringVar()
    dob = StringVar()
    address = StringVar()

    def addDatabase():
        id_1 =  id_.get()
        name1 =  name.get()
        email1 =  email.get()
        mobile1 =  mobile.get()
        gender1 =  gender.get()
        dob1 =  dob.get()
        address1 =  address.get()
        now = time.strftime("%H:%M:%S")
        d = time.strftime("%d:%m:%y")
        # print(id_1,mobile1)
        try:
            q = 'insert into studentdata values(%s,%s,%s,%s,%s,%s,%s,%s,%s)'
            mycursor.execute(q,(id_1,name1,email1,mobile1,address1,dob1,gender1,d,now))
            conn.commit()
            res = messagebox.askyesnocancel('Success',f'ID {id_1} and name {name1} added successfully \ndo you want to clean the form?',parent=top1)
            if(res == True):
                id_.set('')
                name.set('')
                email.set('')
                mobile.set('')
                gender.set('')
                dob.set('')
                address.set('')
        except:
            messagebox.showwarning('Warning!','Id already exist try with some other',parent=top1)

        query = 'Select * from studentdata'
        mycursor.execute(query)
        datas = mycursor.fetchall()

        # deleting the old data and retreiving new data
        stTable.delete(*stTable.get_children())

        for data in datas:
            vv = [data[0],data[1],data[2],data[3],data[3],data[4],data[4],data[5],data[6],data[7],data[8]]
            stTable.insert('',END,values=vv) 

    top1 = Toplevel()
    top1.title("Add to Database")
    top1.geometry("400x430+400+150")
    top1.grab_set()
    top1.config(bg="lightblue")
    top1.resizable(False, False)

    lab = ['Enter Id: ', 'Enter Name: ', "Enter Email: ",
           'Enter Mobile: ', 'Enter Gender: ', 'Enter Dob: ', 'Enter Address: ']
    var_name = [id_, name, email, mobile, gender, dob, address]
    a = 5
    for i in range(0, len(lab)):
        label(top1, lab[i], "lightblue", "black", 2, a, 200, 50)
        Entry(top1, textvariable=var_name[i], font="luicida 15 bold").place(
            x=210, y=a+10, width=180, height=35)
        a += 50
    btn(top1, "Submit", "darkslateblue", "black",
        "darkslateblue", "white", 160, 370, 90, 50, addDatabase)


def searchStudent():
    id_ = StringVar()
    name = StringVar()
    email = StringVar()
    mobile = StringVar()
    gender = StringVar()
    dob = StringVar()
    address = StringVar()
    date = StringVar()

    def searchDatabase():
        id_1 = id_.get()
        name1 = name.get()
        email1 = email.get()
        mobile1 = mobile.get()
        gender1 = gender.get()
        dob1 = dob.get()
        address1 = address.get()
        date1 = date.get()
        # print(id_1,mobile1)
        try:
            if(id_1 != ''):
                q = 'select * from studentdata where id=%s'
                mycursor.execute(q, (id_1))
                datas = mycursor.fetchall()
                stTable.delete(*stTable.get_children())
                for data in datas:
                    vv = [data[0], data[1], data[2], data[3], data[3],
                        data[4], data[4], data[5], data[6], data[7], data[8]]
                    stTable.insert('', END, values=vv)
            elif (name1 != ''):
                q = 'select * from studentdata where name=%s'
                mycursor.execute(q, (name1))
                datas = mycursor.fetchall()
                stTable.delete(*stTable.get_children())
                for data in datas:
                    vv = [data[0], data[1], data[2], data[3], data[3],
                          data[4], data[4], data[5], data[6], data[7], data[8]]
                    stTable.insert('', END, values=vv)
            
            elif (email1 != ''):
                q = 'select * from studentdata where email=%s'
                mycursor.execute(q, (email1))
                datas = mycursor.fetchall()
                stTable.delete(*stTable.get_children())
                for data in datas:
                    vv = [data[0], data[1], data[2], data[3], data[3],
                          data[4], data[4], data[5], data[6], data[7], data[8]]
                    stTable.insert('', END, values=vv)

            elif (mobile1 != ''):
                q = 'select * from studentdata where mobile=%s'
                mycursor.execute(q, (mobile1))
                datas = mycursor.fetchall()
                stTable.delete(*stTable.get_children())
                for data in datas:
                    vv = [data[0], data[1], data[2], data[3], data[3],
                          data[4], data[4], data[5], data[6], data[7], data[8]]
                    stTable.insert('', END, values=vv)

            elif (address1 != ''):
                q = 'select * from studentdata where address=%s'
                mycursor.execute(q, (address1))
                datas = mycursor.fetchall()
                stTable.delete(*stTable.get_children())
                for data in datas:
                    vv = [data[0], data[1], data[2], data[3], data[3],
                          data[4], data[4], data[5], data[6], data[7], data[8]]
                    stTable.insert('', END, values=vv)

            elif (dob1 != ''):
                q = 'select * from studentdata where dob=%s'
                mycursor.execute(q, (dob1))
                datas = mycursor.fetchall()
                stTable.delete(*stTable.get_children())
                for data in datas:
                    vv = [data[0], data[1], data[2], data[3], data[3],
                          data[4], data[4], data[5], data[6], data[7], data[8]]
                    stTable.insert('', END, values=vv)

            elif (gender1 != ''):
                q = 'select * from studentdata where gender=%s'
                mycursor.execute(q, (gender1))
                datas = mycursor.fetchall()
                stTable.delete(*stTable.get_children())
                for data in datas:
                    vv = [data[0], data[1], data[2], data[3], data[3],
                          data[4], data[4], data[5], data[6], data[7], data[8]]
                    stTable.insert('', END, values=vv)

            elif (date1 != ''):
                q = 'select * from studentdata where date=%s'
                mycursor.execute(q, (date1))
                datas = mycursor.fetchall()
                stTable.delete(*stTable.get_children())
                for data in datas:
                    vv = [data[0], data[1], data[2], data[3], data[3],
                          data[4], data[4], data[5], data[6], data[7], data[8]]
                    stTable.insert('', END, values=vv)
        except:
            messagebox.showwarning(
                'Warning!', 'some error occured', parent=top1)

    top1 = Toplevel()
    top1.title("Add to Database")
    top1.geometry("400x480+400+150")
    top1.grab_set()
    top1.config(bg="lightblue")
    top1.resizable(False, False)

    lab = ['Enter Id: ', 'Enter Name: ', "Enter Email: ",
           'Enter Mobile: ', 'Enter Gender: ', 'Enter Dob: ', 'Enter Address: ', 'Enter Date: ']
    var_name = [id_, name, email, mobile, gender, dob, address, date]
    a = 5
    for i in range(0, len(lab)):
        label(top1, lab[i], "lightblue", "black", 2, a, 200, 50)
        Entry(top1, textvariable=var_name[i], font="luicida 15 bold").place(
            x=210, y=a+10, width=180, height=35)
        a += 50
    btn(top1, "Submit", "darkslateblue", "black",
        "darkslateblue", "white", 160, 420, 90, 50, searchDatabase)


def deleteStudent():
    cc = stTable.focus()
    content = stTable.item(cc)
    id_d = content['values'][0]
    q = 'delete from studentdata where id=%s'
    mycursor.execute(q,(id_d))
    conn.commit()

    messagebox.showinfo('Success',f'the item with id {id_d} is deleted successfully ')
    query = 'Select * from studentdata'
    mycursor.execute(query)
    datas = mycursor.fetchall()

    # deleting the old data and retreiving new data
    stTable.delete(*stTable.get_children())

    for data in datas:
        vv = [data[0], data[1], data[2], data[3], data[3],
              data[4], data[4], data[5], data[6], data[7], data[8]]
        stTable.insert('', END, values=vv)




def updateStudent():
    id_ = StringVar()
    name = StringVar()
    email = StringVar()
    mobile = StringVar()
    gender = StringVar()
    dob = StringVar()
    address = StringVar()
    date1 = StringVar()
    time1 = StringVar()


    cc = stTable.focus()
    content = stTable.item(cc)
    items = content['values']


    id_.set(items[0])
    name.set(items[1])
    email.set(items[2])
    mobile.set(items[3])
    address.set(items[4])
    gender.set(items[6])
    dob.set(items[5])
    date1.set(items[7])
    time1.set(items[8])

    def updateDatabase():
        id_1 = id_.get()
        name1 = name.get()
        email1 = email.get()
        mobile1 = mobile.get()
        gender1 = gender.get()
        dob1 = dob.get()
        address1 = address.get()
        date2 = date1.get()
        time2 = time1.get()

        q = 'update studentdata set name = %s,email=%s,mobile=%s,address=%s,dob=%s ,gender=%s,date=%s,time=%s where id =%s'
        mycursor.execute(q,(name1,email1,mobile1,address1,dob1,gender1,date2,time2,id_1))
        conn.commit()
        messagebox.showinfo("Success",f"name {name1} is modified",parent=top1)

        query = 'Select * from studentdata'
        mycursor.execute(query)
        datas = mycursor.fetchall()

        # deleting the old data and retreiving new data
        stTable.delete(*stTable.get_children())

        for data in datas:
            vv = [data[0], data[1], data[2], data[3], data[3],
                data[4], data[4], data[5], data[6], data[7], data[8]]
            stTable.insert('', END, values=vv)



    
    top1 = Toplevel()
    top1.title("Add to Database")
    top1.geometry("400x530+400+120")
    top1.grab_set()
    top1.config(bg="lightblue")
    top1.resizable(False, False)

    lab = ['Enter Id: ', 'Enter Name: ', "Enter Email: ",
           'Enter Mobile: ', 'Enter Gender: ', 'Enter Dob: ', 'Enter Address: ', 'Enter Date: ', 'Enter Time: ']
    
    var_name = [id_, name, email, mobile, gender, dob, address, date1, time1]
    a = 5
    for i in range(0, len(lab)):
        label(top1, lab[i], "lightblue", "black", 2, a, 200, 50)
        Entry(top1, textvariable=var_name[i], font="luicida 15 bold").place(
            x=210, y=a+10, width=180, height=35)
        a += 50
    btn(top1, "Submit", "darkslateblue", "black",
        "darkslateblue", "white", 160, 470, 90, 50, updateDatabase)


def showAll():
    # global conn, mycursor
    query = 'Select * from studentdata'
    mycursor.execute(query)
    datas = mycursor.fetchall()

    # deleting the old data and retreiving new data
    stTable.delete(*stTable.get_children())

    for data in datas:
        vv = [data[0], data[1], data[2], data[3], data[3],
              data[4], data[4], data[5], data[6], data[7], data[8]]
        stTable.insert('', END, values=vv)


def exportData():
    ff = filedialog.asksaveasfile()
    gg = stTable.get_children()

    id,name,email,mobile,address,gender,dob,date,time1 = [],[],[],[],[],[],[],[],[]

    for i in gg:
        content = stTable.item(i)
        pp = content['values']
        id.append(pp[0]), name.append(pp[1]), mobile.append(pp[3]), email.append(pp[2]), address.append(pp[4]), gender.append(pp[5]),
        dob.append(pp[6]), date.append(pp[7]), time1.append(pp[8])

    # print(ff.name)
    dd = ['Id', 'Name', 'Email', 'Mobile', 'Address',
          'Gender', 'D.O.B', 'Added Date', 'Added Time']
    df = pandas.DataFrame(list(zip(id, name, email, mobile,
                                   address, gender, dob, date, time1)), columns=dd)
    
    paths = r'{}.csv'.format(ff.name)
    df.to_csv(paths, index=False)
    messagebox.showinfo(
        'Notifications', 'Student data is Saved {}'.format(paths))

    # openpyxl




def Exit():
    res = messagebox.askyesnocancel(
        "Warning!", "Do you want to close the window")
    if (res == True):
        root.destroy()

######### Functions Area Ended ###########


# creating basic layout
root = Tk()
root.title("Student Management Program")
# root.iconbitmap('') include later
root.geometry('1100x650+100+20')
root.config(bg="lightgray")
root.resizable(False, False)


# Creating Frames within the root

dataFrame = Frame(root, bg='darkgray')
dataFrame.place(x=10, y=110, width=450, height=520)

studentFrame = Frame(root, bg='darkgray')
studentFrame.place(x=480, y=110, width=600, height=520)


#  creating top bar with clock heading and add db button
clk = Label(root, text="", bg="lightblue", fg="white", font="luicida 15 bold")
clk.place(x=10, y=10, width=200, height=82)
clock()

textvar = "Welcome To Student Management System"
count = 0
tex = ""
head = Label(root, text=textvar, bg="lightgreen",
             fg="black", font="luicida 15 bold")
head.place(x=270, y=10, width=500, height=60)
rotate1()

btn(root, "Add To Database", "lightblue", "white",
    "lightblue", "black", 860, 10, 200, 70, dbConnect)


# ###  creating the left side frame
Label(dataFrame, text="--Welcome--", font="chiller 25 bold",
      bg="darkgray").pack(side=TOP, expand=True)

btn_names = ["Add Student", "Search Student", "Delete Student",
             "Update Student", "Show All", "Export Data", "Exit"]
btn_cmd = [addStudent, searchStudent, deleteStudent,
           updateStudent, showAll, exportData, Exit]

for btns in range(0, len(btn_names)):
    Button(dataFrame, text=btn_names[btns], font="luicida 16 bold",
           bg="lightslateblue", command=btn_cmd[btns]).pack(side=TOP, expand=True)


# ### Creating the right frame
scroll_x = Scrollbar(studentFrame, orient=HORIZONTAL)
scroll_y = Scrollbar(studentFrame, orient=VERTICAL)

# columns
column = ['ID', "Name", "Email", "Phone", "D.O.B",
          "Address", "Date Added", "Time Added"]

stTable = Treeview(studentFrame, columns=column,
                   xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)
scroll_x.pack(side=BOTTOM, fill=X)
scroll_y.pack(side=RIGHT, fill=Y)
scroll_x.config(command=stTable.xview)
scroll_y.config(command=stTable.yview)


# setting heading
for i in range(0, len(column)):
    stTable.heading(column[i], text=column[i])

stTable['show'] = "headings"

# for styling
style = ttk.Style()
style.configure("Treeview.Heading", font="luicida 15 bold", foreground="black")
style.configure("Treeview", font="luicida 15 ",
                foreground="white", background="lightslateblue")

stTable.pack(fill=BOTH, expand=1,ipady=4,anchor='ne')
root.mainloop()
