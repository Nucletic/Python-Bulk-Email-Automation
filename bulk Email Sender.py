# dftlozwxupycipab
from email.message import EmailMessage
import ssl
import smtplib
from tkinter.filedialog import *
from tkinter import *
from openpyxl import load_workbook
win = Tk()
win.geometry("900x600")


def send_mail():
    # Loop to send email
    for i in EmailReciever:
        Email = EmailMessage()
        Email["From"] = EmailSender
        Email["To"] = i
        Email["Subject"] = EmailSubject
        Email.set_content(EmailBody)
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
            smtp.login(EmailSender, EmailPassword)
            smtp.sendmail(EmailSender, i, Email.as_string())
        print("email sent to", i)


def take_input():
    global em_text_reciver, em_send_input, em_send_pass_input, EmailSender, EmailPassword, EmailReciever, em_body_lab_input
    global em_subject_lab_input, first_index_lab_input, last_index_lab_input, EmailSubject, EmailBody
    EmailSender = em_send_input.get().strip()
    EmailPassword = em_send_pass_input.get("1.0", 'end-1c')
    EmailSubject = em_subject_lab_input.get()
    EmailBody = em_body_lab_input.get("1.0", 'end-1c')
    first_index = first_index_lab_input.get()
    last_index = last_index_lab_input.get()
    text_email_reciver = em_text_reciver.get("1.0", 'end-1c')
    if len(last_index) > 0:
        fi = int(first_index)
        li = int(last_index)
        fi -= 1
        EmailReciever = id_list[fi:li]
        send_mail()
    else:
        if len(text_email_reciver) == 0:
            pass
        else:
            x = text_email_reciver.split()
            EmailReciever = list(x)
            del x
            send_mail()

# select excel file


def select_file():
    global id_list
    excel_file = askopenfilename()
    if len(excel_file) == 0:
        pass
    else:
        Label(win, text=excel_file).place(x=220, y=60)
        load_file = load_workbook(excel_file)
        files = load_file.active
        ran = files
        id_list = list()
        for i in ran:
            for x in i:
                id_list.append(x.value)


# Email
em_send_lab = Label(win, text="Email:").place(x=10, y=11)
em_send_input = Entry(win, width=80)
em_send_input.place(x=90, y=13)
# password
em_send_pass_lab = Label(win, text="Password:").place(x=10, y=31)
em_send_pass_input = Text(win, height=1, width=60)
em_send_pass_input.place(x=90, y=33)
# reciver
em_reciver_lab = Label(win, text="Reciver Email:").place(x=10, y=60)
em_reciver_lab_input = Button(
    win, text="Select File", width=15, command=select_file)
em_reciver_lab_input.place(x=90, y=58)
# text reciver
em_text_reciver_lab = Label(win, text="Reciver Email:").place(x=10, y=87)
em_text_reciver = Text(win, width=60, height=5)
em_text_reciver.place(x=90, y=90)
# first index
first_index_lab = Label(win, text="First Index:").place(x=10, y=178)
first_index_lab_input = Entry(win, text="First Index:", width=15)
first_index_lab_input.place(x=90, y=180)
# last index
last_index_lab = Label(win, text="last Index:").place(x=10, y=199)
last_index_lab_input = Entry(win, text="Last Index:", width=15)
last_index_lab_input.place(x=90, y=202)
# Subject
em_subject_lab = Label(win, text="Subject:").place(x=10, y=221)
em_subject_lab_input = Entry(win, width=80)
em_subject_lab_input.place(x=90, y=224)
# Body
em_body_lab = Label(win, text="Body:").place(x=10, y=244)
em_body_lab_input = Text(win, height=10, width=80)
em_body_lab_input.place(x=90, y=246)
# Send Mail
submit_top = Button(win, text="Send Email", width=15,
                    command=take_input).place(x=90, y=420)

# Run Tkinter
win.mainloop()
