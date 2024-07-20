#!/usr/bin/env python
# coding: utf-8

# In[22]:


import tkinter
from tkinter import ttk
from tkinter import *
from docxtpl import DocxTemplate
from tkcalendar import Calendar
from datetime import datetime
from docx2pdf import convert

window = tkinter.Tk()
window.title("Generate edSlash Certificate")
frame = tkinter.Frame(window) 
frame.pack(padx=60, pady=40)

def exit_app():
    window.destroy()

def generat_certificate():
    
    doc = DocxTemplate("Certificate_template.docx")
    name = sname_entry.get()
    application_no = application_entry.get()
    cname = combo_box.get()
    now = datetime.now()
    current_date = now.date()
    formatted_date = current_date.strftime('%d %B %Y')  
  
    doc.render({"name" :name,
                "application_no" : application_no,
            "course_name" :cname,
            "date" : formatted_date})
    
    exit_app()

    doc_name = name +application_no +".docx"
    doc.save(doc_name)
    pdf_name = name +application_no +".docx"
    convert(pdf_name)
    
    print("Certificate Generate Successful ")

sname_label = tkinter.Label(frame, text = "Name")
sname_label.grid(row = 0,column = 0,padx = 20, pady=5,sticky ="w")
sname_entry = tkinter.Entry(frame)
sname_entry.grid(row = 0,column = 1,padx = 20, pady=5)

application_lebel = tkinter.Label(frame, text = "Application no.")
application_lebel.grid(row =1 ,column = 0,padx = 20,pady = 5,sticky = "w")
application_entry = tkinter.Entry(frame)
application_entry.grid(row = 1, column = 1, padx = 50,pady = 5)

course_label = tkinter.Label(frame, text = "Course Name")
course_label.grid(row = 2,column = 0,padx = 20, pady=5,sticky ="w")
combo_box = ttk.Combobox(frame, values=["Basic and Advance C programming",
                                        "Basic and Advance C++ programming",
                                        "Basic and Advance Python programming",
                                        "Basic and Advance SQL programming",
                                        "Data Science","Data Analysis",
                                        "Machine Learning","Generative AI",])
combo_box.grid(row = 2,column = 1,padx = 20, pady=5)

generate_receipt_button = tkinter.Button(frame , text = "generate_certificate", command = generat_certificate)
generate_receipt_button.grid(row=3, column=0, columnspan = 2, sticky ="news",padx = 20, pady=5) 

window.mainloop()

