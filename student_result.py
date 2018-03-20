from tkinter import *
import tkinter
from  xlwt import *
import smtplib
wb=Workbook()
top=Tk()
f0=Frame(top)
f1=Frame(top,height=800,width=700)
import mysql.connector
conn = mysql.connector.connect(user='root', password='',  host='127.0.0.1', database='stud')
cursor = conn.cursor()
#cursor.execute("create table student7575 (snm varchar(20),roll int,acy varchar(20),dcn int,la int,android int,se int ,algo int,aj int, net int,status varchar(20))")

print("connection done")

def disable():
    
    menubar.entryconfig("feedback", state="disabled")
    menubar.entryconfig("File", state="disabled")
    menubar.entryconfig("signout", state="disabled")
    menubar.entryconfig("Help", state="disabled")
cv=StringVar()
snm=StringVar()
sr=StringVar()
acy=StringVar()
m1=StringVar()
m2=StringVar()
m3=StringVar()
m4=StringVar()
m5=StringVar()
m6=StringVar()
m7=StringVar()
xlnm=StringVar()
emr=StringVar()
r=0
c=0

f=Frame(top)
fo=Frame(top,width=400,height=400)

def forgetpass():
    receiver='prabhatkushwaha7575@yahoo.com'
    sender='prabhatkushwaha7575@gmail.com'
    senderpass='8286641344'
    message='your password is prabhat'
    print("sender:",sender,"reciver:",receiver,"sendpass",senderpass,"massage:",message)
    try:
        server=smtplib.SMTP('smtp.gmail.com',587)
        server.starttls()
        server.login(sender,senderpass)
        server.sendmail(sender, receiver, message)
        server.close()
        messagebox.showinfo(":)","Thank You.......!","plz check your email your password send on ragister email id ")
            
    except:
        messagebox.showinfo(":(","No proper Connection or Email id is incorrect")
    return None


def openf():
    fo.grid()
    f1.grid_forget()
    data="select * from student75"
    data1=cursor.execute(data)
    data1 = cursor.fetchall()
        
    #[int(dcn),int(la),int(android),int(se),int(algo),int(aj),int(net)]
    Label(fo,text="Name",font="Times 15 bold").grid(row=0,column=0)
    Label(fo,text="Roll No",font="Times 15 bold").grid(row=0,column=1)
    Label(fo,text="Year",font="Times 15 bold").grid(row=0,column=2)
    Label(fo,text="DCN",font="Times 15 bold").grid(row=0,column=3)
    Label(fo,text="LinearAlgebra",font="Times 15 bold").grid(row=0,column=4)
    Label(fo,text="Android",font="Times 15 bold").grid(row=0,column=5)
    Label(fo,text="SoftwearEng",font="Times 15 bold").grid(row=0,column=6)
    Label(fo,text="Algorithm",font="Times 15 bold").grid(row=0,column=7)
    Label(fo,text="Adv. Java",font="Times 15 bold").grid(row=0,column=8)
    Label(fo,text="Dot Net",font="Times 15 bold").grid(row=0,column=9)
    Label(fo,text="Status",font="Times 15 bold").grid(row=0,column=10)
#---------------------------------------------
    Lb1=Listbox(fo,width=10,height=40)
    Lb1.grid(row=1,column=0)
    Lb2=Listbox(fo,width=10,height=40)
    Lb2.grid(row=1,column=1)
    Lb3=Listbox(fo,width=10,height=40)
    Lb3.grid(row=1,column=2)
    Lb4=Listbox(fo,width=10,height=40)
    Lb4.grid(row=1,column=3)
    Lb5=Listbox(fo,width=10,height=40)
    Lb5.grid(row=1,column=4)
    Lb6=Listbox(fo,width=10,height=40)
    Lb6.grid(row=1,column=5)
    Lb7=Listbox(fo,width=10,height=40)
    Lb7.grid(row=1,column=6)
    Lb8=Listbox(fo,width=10,height=40)
    Lb8.grid(row=1,column=7)
    Lb9=Listbox(fo,width=10,height=40)
    Lb9.grid(row=1,column=8)
    Lb10=Listbox(fo,width=10,height=40)
    Lb10.grid(row=1,column=9)
    Lb11=Listbox(fo,width=10,height=40)
    Lb11.grid(row=1,column=10)
    b1=Button(fo,text="goback",font="Times 20 bold",command=goback).grid(row=1,column=11)
    
    a=1
    for i in data1:
        Lb1.insert(a,i[0])
        Lb2.insert(a,i[1])
        Lb3.insert(a,i[2])
        Lb4.insert(a,i[3])
        Lb5.insert(a,i[4])
        Lb6.insert(a,i[5])
        Lb7.insert(a,i[6])
        Lb8.insert(a,i[7])
        Lb9.insert(a,i[8])
        Lb10.insert(a,i[9])
        Lb11.insert(a,i[10])
        a+=1

    
def goback():
    fo.grid_forget()
    f.grid_forget()
    f1.grid()
def fb():
    f.grid()
    f1.grid_forget()

    def massage():
       
        receiver='prabhatkushwaha7575@yahoo.com'
        sender=semail.get()
        senderpass=spas.get()
        message=feed.get()
        print("sender:",sender,"reciver:",receiver,"sendpass",senderpass,"massage:",message)
        try:
            server=smtplib.SMTP('smtp.gmail.com',587)
            server.starttls()
            server.login(str(sender),str(senderpass))
            server.sendmail(sender, receiver, message)
            server.close()
            messagebox.showinfo(":)","Thank You.......!")
            
        except:
                messagebox.showinfo(":(","No proper Connection or Email id is incorrect")
        return None


    feed=StringVar()
    semail=StringVar()
    spas=StringVar()    
    
   
    flb=Label(f,text="Feedback",font="Times 15 bold").grid(row=0,column=0)
    fe=Entry(f,font="Times 20 bold",textvariable=feed).grid(row=0,column=1)
    feed.set(" ")
    emsl=Label(f,text="Enter your Email Id",font="Times 15 bold").grid(row=1,column=0)
    es=Entry(f,font="Times 20 bold",textvariable=semail).grid(row=1,column=1)
    epl=Label(f,text="Enter your Email Id Password",font="Times 15 bold").grid(row=2,column=0)
    epas=Entry(f,font="Times 20 bold",textvariable=spas,show="*").grid(row=2,column=1)
   
    b=Button(f,text="Send",font="Times 20 bold",command=massage).grid(row=3,column=1)
    b1=Button(f,text="goback",font="Times 20 bold",command=goback).grid(row=4,column=1)
    
        
sheet=wb.add_sheet('result',cell_overwrite_ok=True)
def xlsheet():
    data="select * from student75"
    data1=cursor.execute(data)
    data1 = cursor.fetchall()
    st=xlnm.get()
    print("select query execute........")
    print("-----------------------------")
    first=0
    print(data1)
    #-----------------for heding [int(dcn),int(la),int(android),int(se),int(algo),int(aj),int(net)] 
    nm=["name","roll","year","dcn","la","android","se","algo","aj","dot net","status"]
    for x in range(0,11):
        sheet.write(0,x,nm[x])
    alldata=[]
    temp=1
    while(first<len(data1)):
        sec=0
        #temp=[]
        while(sec<11):
            sheet.write(temp,sec,data1[first][sec])
            print(data1[first][sec])
            wb.save(str(st)+'.xls')
            sec=sec+1
        temp=temp+1        
        first=first+1
def delete():
    r=sr.get()
    q="delete from student75 where roll=%d"%(int(r))
    print(q)
    data=cursor.execute(q)
    cursor.fetchone()
    snm.set('')
    sr.set('')
    acy.set('')
    m1.set('')
    m2.set('')
    m3.set('')
    m4.set('')
    m5.set('')
    m6.set('')
    m7.set('')
    

    
def insert():
    n=snm.get()
    r=sr.get()
    y=acy.get()
    dcn=m1.get()
    la=m2.get()
    android=m3.get()
    se=m4.get()
    algo=m5.get()
    aj=m6.get()
    net=m7.get()
    state=[int(dcn),int(la),int(android),int(se),int(algo),int(aj),int(net)]
    chk=0
    temp2=0
    print(state)
    while temp2<7:
        if state[temp2]<35:
            chk=chk+1
        temp2+=1
    if(chk>3):
        q='insert into student75(snm,roll,acy,dcn,la,android,se,algo,aj,net,status) values("%s",%d,"%s",%d,%d,%d,%d,%d,%d,%d,"%s")'%(n,int(r),y,int(dcn),int(la),int(android),int(se),int(algo),int(aj),int(net),"Fail")
    else:
        q='insert into student75(snm,roll,acy,dcn,la,android,se,algo,aj,net,status) values("%s",%d,"%s",%d,%d,%d,%d,%d,%d,%d,"%s")'%(n,int(r),y,int(dcn),int(la),int(android),int(se),int(algo),int(aj),int(net),"pass")
    print(q)
    data=cursor.execute(q)
    cursor.fetchone()
    snm.set('')
    sr.set('')
    acy.set('')
    m1.set('')
    m2.set('')
    m3.set('')
    m4.set('')
    m5.set('')
    m6.set('')
    m7.set('')
    

    
def update():
    r=sr.get()
    data="select * from student75 where roll=%d"%(int(r))
    data1=cursor.execute(data)
    data1 = cursor.fetchall()
    print(data1)
    sl=[None,None,None,None,None,None,None,None,None,None,None]
    
    
    #---------------------------------------------
    n=snm.get()
    r=sr.get()
    y=acy.get()
    dcn=m1.get()
    la=m2.get()
    android=m3.get()
    se=m4.get()
    algo=m5.get()
    aj=m6.get()
    net=m7.get()
    state=[int(dcn),int(la),int(android),int(se),int(algo),int(aj),int(net)]
    chk=0
    temp2=0
    print(state)
    while temp2<7:
        if state[temp2]<35:
            chk=chk+1
        temp2+=1
    if chk>3:
        q="update student75 set  snm='%s',roll=%d,acy='%s',dcn=%d,la=%d,android=%d,se=%d,algo=%d,aj=%d,net=%d,status ='%s' where roll=%d"%(n,int(r),y,int(dcn),int(la),int(android),int(se),int(algo),int(aj),int(net),"Fail",int(r))
    else:
        q="update student75 set  snm='%s',roll=%d,acy='%s',dcn=%d,la=%d,android=%d,se=%d,algo=%d,aj=%d,net=%d,status ='%s' where roll=%d"%(n,int(r),y,int(dcn),int(la),int(android),int(se),int(algo),int(aj),int(net),"pass",int(r))
    print(q)
    data=cursor.execute(q)
    cursor.fetchone()
    snm.set('')
    sr.set('')
    acy.set('')
    m1.set('')
    m2.set('')
    m3.set('')
    m4.set('')
    m5.set('')
    m6.set('')
    m7.set('')
    
        
def login():
    a=var1.get()
    b=var.get()
    
    try:
        if(a=='prabhat' and b=="prabhat"):
            lable.grid_forget()
        
            f0.grid_forget()
            f1.grid()
            menubar.entryconfig("feedback", state="normal")
            menubar.entryconfig("File", state="normal")
            menubar.entryconfig("signout", state="normal")
            menubar.entryconfig("Help", state="normal")
            
            #t1=Label(f1,text="Enter Subject Marks ..........",font="Times 20 bold").grid(row=0,column=0)
            t2=Label(f1,text="Student \n Name",font="Times 15 bold").grid(row=0,column=0)
            e2=Entry(f1,font="Times 20 bold",textvariable=snm).grid(row=0,column=1)
            
            t3=Label(f1,text="Student \n RollNo",font="Times 15 bold").grid(row=0,column=2)
            e3=Entry(f1,font="Times 20 bold",textvariable=sr).grid(row=0,column=3)

            t4=Label(f1,text="Student \n Acadmic year",font="Times 15 bold").grid(row=0,column=4)
            e4=Entry(f1,font="Times 20 bold",textvariable=acy).grid(row=0,column=5)

            t1=Label(f1,text="Enter Subject Marks ..........",font="Times 20 bold").grid(row=1,column=1)
            #-------------------------------------------------------
            t5=Label(f1,text="DCN",font="Times 15 bold").grid(row=2,column=0)
            e5=Entry(f1,font="Times 20 bold",textvariable=m1).grid(row=2,column=1)
            
            t6=Label(f1,text="Linear \n Algebra",font="Times 15 bold").grid(row=3,column=0)
            e6=Entry(f1,font="Times 20 bold",textvariable=m2).grid(row=3,column=1)

            t7=Label(f1,text="Android",font="Times 15 bold").grid(row=4,column=0)
            e7=Entry(f1,font="Times 20 bold",textvariable=m3).grid(row=4,column=1)


            t8=Label(f1,text="SoftWare\nEnginering",font="Times 15 bold").grid(row=5,column=0)
            e8=Entry(f1,font="Times 20 bold",textvariable=m4).grid(row=5,column=1)
            
            t9=Label(f1,text="Algorithm",font="Times 15 bold").grid(row=6,column=0)
            e9=Entry(f1,font="Times 20 bold",textvariable=m5).grid(row=6,column=1)

            t10=Label(f1,text="Adv Java",font="Times 15 bold").grid(row=7,column=0)
            e10=Entry(f1,font="Times 20 bold",textvariable=m6).grid(row=7,column=1)

            #------------------------------------------------------------------------------
            t11=Label(f1,text="Dot Net",font="Times 15 bold").grid(row=8,column=0)
            e11=Entry(f1,font="Times 20 bold",textvariable=m7).grid(row=8,column=1)
             #------------------------------
            
            bins= Button(f1, text ="    insert   ",font="Times 20 bold", command=insert)
            bins.grid(row=3, column=3)
            bup= Button(f1, text ="   update  ",font="Times 20 bold", command=update)
            bup.grid(row=4, column=3)
            bdel= Button(f1, text ="    delete   ",font="Times 20 bold", command=delete)
            bdel.grid(row=5, column=3)
            
            e12=Entry(f1,font="Times 20 bold",textvariable=xlnm).grid(row=6,column=3)
            xlnm.set("enter xlheet name")

            bxl= Button(f1, text ="  Create  xlsheet  ",font="Times 20 bold", command=xlsheet)
            bxl.grid(row=6, column=4)
            
            
            
        else:
            tkinter.messagebox.showinfo(" Hello Sir", " PLEASE ENTER CORRECT USERNAME AND PASSWORD")
    except :
        tkinter.messagebox.showinfo(" Hello Sir", " PLEASE ENTER CORRECT USERNAME AND PASSWORD")



 #--------------------------forget admin------------------------

#--------------------------------content of login frame
photo=PhotoImage(file="pic2.png")
lable=Label(image=photo)
lable.grid()

f0.grid()
tag=Label(f0,font="Times 70 bold",text="    Admin Login   ").grid(row=0,column=1,columnspan=2)
usnm=Label(f0,font="Times 20 bold",text="Username")
pas= Label(f0,font="Times 20 bold", text="Password")
usnm.grid(row=1, column=1)
pas.grid(row=2, column=1)
var1=StringVar()
usnme= Entry(f0,width=30,bd="3px",font="Times 20 bold",textvariable=var1,relief=SUNKEN) 
usnme.grid(row=1, column=2)
var=StringVar()
pase= Entry(f0,textvariable=var,font="Times 20 bold",width=30,show="*",bd="3px",relief=SUNKEN) 
pase.grid(row=2, column=2)

bnext= Button(f0, text ="Login",font="Times 20 bold", command=login )
bnext.grid(row=3, column=2)

bforpass= Button(f0, text ="Forget Password",font="Times 20 bold", command=forgetpass )
bforpass.grid(row=4, column=2)
#---------------------------

#-------------------------------in login frame--
def donothing():
    print("do nathing ")
#--------help
def about():
    tkinter.messagebox.showinfo(" about ", " Software created by Prabhat kushwaha ")
def helps():
    tkinter.messagebox.showinfo(" help ", " please visit to www.studyhub.com ")
    
def signouts():
    bforpass.grid()
    lable.grid()
    f1.grid_forget()
    fo.grid_forget()
    f0.grid()
#-----------------menu bar ------------------
menubar=Menu(top)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="Open", command=openf)
filemenu.add_separator()
filemenu.add_command(label="Exit", command=top.quit)
menubar.add_cascade(label="File", menu=filemenu)

helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="Help Index", command=helps)
helpmenu.add_command(label="About...", command=about)
menubar.add_cascade(label="Help", menu=helpmenu)
feedback=Menu(menubar,tearoff=0)
feedback.add_command(label="feedback",command=fb)
menubar.add_cascade(label="feedback", menu=feedback)

signout=Menu(menubar, tearoff=0)
signout.add_command(label="signout",command=signouts)
menubar.add_cascade(label="signout", menu=signout)

top.config(menu=menubar)


disable()
#----------------------
