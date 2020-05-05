import smtplib, ssl
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
from tkinter import *
import datetime
import numpy as np
import cv2
from matplotlib import pyplot as plt
import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
import time
from tkinter import messagebox
import sys
from tkinter import font
from tkinter import ttk
import pandas as pd
import re
import xlrd
import time
from tkinter.messagebox import *
from validate_email import validate_email
import io
root = Tk()

root.geometry("400x450")
pathname = os.path.dirname(sys.argv[0])
root.title("Auto send")
root.iconbitmap(os.path.abspath(pathname) + "\\auto.ico")
root.resizable(width=False, height=False)
xcel_path = None
text_path = None
pdf_path = None
comboExample = None
password = None
email2 = None
subject1 = None
imageback1 = None
wb = None

x = datetime.datetime.now()
m = x.day
b = x.month

sheet_names=None
pathname = os.path.dirname(sys.argv[0])
context = ssl.create_default_context()
def get_contacts():

    sheet1 = comboExample.get()
    sheet = wb.sheet_by_name(sheet1)
    emails = []
    def is_valid_email(email):
        if type(email) == str :
            if len(email) > 7:
                return bool(re.match("^.+@(\[?)[a-zA-Z0-9-.]+.([a-zA-Z]{2,3}|[0-9]{1,3})(]?)$", email))
        else:
            return

    for i in range(sheet.nrows):
        for s in range(sheet.ncols):
            if is_valid_email(sheet.cell_value(i, s)) == True:
                emails.append(sheet.cell_value(i, s))
            else:
                continue



    return emails
def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
         template_file_content = template_file.read()
    return (template_file_content)
def main():
    global frame4
    frame2.place_forget()
    frame4 = Frame(root,borderwidth=5, relief=RIDGE)
    frame4.place(x=0,y=160,width=400,height = 290)
    frame3 = Frame(frame4,borderwidth=1)
    frame3.place(x=0,y=0,width=390,height = 200)
    Exit = Button(frame4, text="Exit",bd = 3, command=ask_quit)
    Exit.place(x=300, y=240, width=80, height=30)
    def go():
        frame4.destroy()
        getall()
    Back = Button(frame4, text="Back",bd = 3, command=go)
    Back.place(x=20, y=240, width=80, height=30)
    list = Listbox(frame3, height=50, width=60)
    scroll = Scrollbar(frame3, command=list.yview)
    list.configure(yscrollcommand=scroll.set)
    list.pack(side=LEFT)
    list.update()
    scroll.pack(side=RIGHT, fill=Y)
    progress = ttk.Progressbar(frame4, orient=HORIZONTAL, length=380, mode='determinate')
    progress.place(x= 10, y= 205,width =370,height =13)
    emails = get_contacts()  # read contacts
    r = len(emails)

    def check():
        if r == 0:
            messagebox.showerror("Error", "No email were found, Upload another Excel file")
            frame4.place_forget()
            getall()
        else:
            return
    check()






    message_template = read_template(text_path)
    MY_ADDRESS = email2.get()
    PASSWORD = password.get()
    subject = subject1.get()
    s = smtplib.SMTP_SSL('smtp.gmail.com', 465,context=context)
    s.login(MY_ADDRESS, PASSWORD)
    g = 0
    progress['maximum'] = 100

    for email in emails:

        msg = MIMEMultipart()  # create a message
        message = message_template
        msg['From'] = MY_ADDRESS
        msg['To'] = email
        msg['Subject'] = subject
        filename = pdf_path
        fo = open(filename, "rb")
        attach = MIMEApplication(fo.read(), _subtype="ppt")
        fo.close()
        lastnamepath = os.path.basename(os.path.normpath(filename))
        attach.add_header('Content-Disposition', 'attachment', filename=lastnamepath)
        msg.attach(MIMEText(message))
        msg.attach(attach)
        list.insert(END, email)
        list.update()
        g+=(100/r)
        progress['value'] = g
        frame4.update_idletasks()
        time.sleep(0.5)

        s.send_message(msg)
        if progress['value'] == 100:
            done = Label(frame4, text= "Done")
            done.place(x=300, y=220, width=80, height=20)

        del msg

    s.quit()


def ask_quit():
        if messagebox.askokcancel("Quit", "ARE YOU SURE YOU WANT TO QUIT ? "):
            root.destroy()
            sys.exit(0)
def getall():
    global frame2,subject1
    frame1.place_forget()
    frame2 = Frame(root,borderwidth=5, relief=RIDGE)
    frame2.place(x=0,y=180,width=400,height = 270)
    def picklist():
        global comboExample,wb,sheet_names
        labelTop = tk.Label(frame2, text="Choose your sheet ")
        labelTop.place(x=10, y=35, width=120, height=25)
        if len(xcel_path) != 0:
            wb = xlrd.open_workbook(xcel_path)

            sheet_names = wb.sheet_names()
            comboExample = ttk.Combobox(frame2, values=sheet_names)
            comboExample.place(x=140, y=35, width=230, height=25)
            comboExample.current(0)
        else:
            return
    def xcel():
        global xcel_path
        xcel1.delete(0, 'end')
        xcel_path = filedialog.askopenfilename(filetypes=(('Excel files', 'xlsx'),))
        xcel1.insert(0, xcel_path)
        picklist()
    def text():
        global text_path
        text1.delete(0, 'end')
        text_path = filedialog.askopenfilename(filetypes=(('text files', 'txt'),))
        #print(text_path)
        text1.insert(0, text_path)
    def pdf():
        global pdf_path
        pdf1.delete(0, 'end')
        pdf_path = filedialog.askopenfilename(filetypes=(('PDF files', 'pdf'),))
        #print(pdf_path)
        pdf1.insert(0, pdf_path)
    getfile1 = Button(frame2, text="Upload Excel",bd = 3, command=xcel)
    getfile1.place(x=10, y=10, width=120, height=25)
    xcel1 = Entry(frame2)
    xcel1.place(x=140, y=10, width=230, height=25)
    getfile2 = Button(frame2, text="Upload Email content ", bd=3, command=text)
    getfile2.place(x=10, y=70, width=120, height=25)
    text1 = Entry(frame2)
    text1.place(x=140, y=70, width=230, height=25)
    getfile3 = Button(frame2, text="Upload CV", bd=3, command=pdf)
    getfile3.place(x=10, y=110, width=120, height=25)
    pdf1 = Entry(frame2)
    pdf1.place(x=140, y=110, width=230, height=25)
    subject = Label(frame2, text="Subject : ", bd=3)
    subject.place(x=10, y=170, width=120, height=25)
    subject1 = Entry(frame2)
    subject1.place(x=140, y=170, width=230, height=25)
    def goback():
        frame2.place_forget()
        start()
    def gonext():
        if (((xcel_path != None) and (len(xcel_path)!= 0)) and  ((text_path != None)and (len(text_path)!= 0)) and ((pdf_path!=None)and(len(pdf_path)!= 0))):
            main()
        else:
            fill = Label(frame2,text="Fill all empty square", fg = "red")
            fill.place(x=120, y=200, width=230, height=20)
            frame2.after(2000,fill.destroy)

    done = Button(frame2, text="Next",bd = 3, command=gonext)
    done.place(x=300, y=220, width=80, height=30)
    exit = Button(frame2, text="Back",bd = 3, command=goback)
    exit.place(x=20, y=220, width=80, height=30)
def start():
    global email2,password ,imageback1,frame1
    all = cv2.imread(os.path.abspath(pathname) + "\\auto1.png")
    imagebackground = cv2.resize(all, (100, 100))
    bg_image = Image.fromarray(imagebackground)
    imageback1 = ImageTk.PhotoImage(image=bg_image)
    x = Label(root, image=imageback1)
    x.place(x=150, y=50, width=100, height=100)
    frame1 = Frame(root)
    frame1.place(x=0,y=200,width = 400,height = 250)
    email1 = Label(frame1, text="Email")
    email1.configure(font=("Calisto MT", 20, "bold"))
    email1.place(x=50, y=10, width=300, height=25)
    email2 = Entry(frame1)
    email2.place(x=50, y=50, width=300, height=25)
    password1 = Label(frame1, text="Password")
    password1.configure(font=("Calisto MT", 20, "bold"))
    password1.place(x=50, y=90, width=300, height=25)
    password = Entry(frame1, show="*")
    password.place(x=50, y=130, width=300, height=25)
    def next():
        def is_valid_email(email):
            if len(email) > 7:
                return bool(re.match("^.+@(\[?)[a-zA-Z0-9-.]+.([a-zA-Z]{2,3}|[0-9]{1,3})(]?)$", email))
        f = is_valid_email(email2.get())
        if (f == True and len(password.get()) > 4) :
            getall()
        else:
            error = Label(frame1,text = "Wrong email or password ",fg = "red")
            error.place(x=50, y=160, width=300, height=25)
            frame1.after(2000, error.destroy)
            return

    def info():
        root1 = Toplevel(root)
        root1.geometry("200x100")
        root1.title("Information")
        root1.iconbitmap(os.path.abspath(pathname) + "\\auto.ico")
        root.resizable(width=False, height=False)
        contact = Label(root1,text = " contact me : ")
        contact.pack()
        contact1 = Label(root1, text=" noor.ziad1994@gmail.com ")
        contact1.pack()
        root1.lift()
        root1.grab_set()

        


    done = Button(frame1, text="Next",bd = 3, command=next)
    done.place(x=300, y=200, width=80, height=30)
    exit = Button(frame1, text="Info",bd = 3, command=info)
    exit.place(x=20, y=200, width=80, height=30)


if (m <= 20 and b <= 6):
    start()
else:
    qe = Label(root, text="THE PROGRAM EXPIRED ")
    qe.configure(font=("Arial", 20, "bold"))
    qe.pack(padx=0, pady=20)
    qt = Label(root, text="REACTIVATE IT ")
    qt.configure(font=("Arial", 20, "bold"))
    qt.pack()
    qo = Label(root, text="Contact us : noor.ziad1994@gmail.com")
    qo.configure(font=("Arial", 15, "bold"))
    qo.pack(padx=0, pady=100)




root.protocol("WM_DELETE_WINDOW", ask_quit)
root.mainloop()
