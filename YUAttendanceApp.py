# -*- coding: utf-8 -*-
"""
Created on Tue Jul  7 23:58:17 2020

@author: ayman
"""
import pandas as pd
import xlsxwriter
from tkinter import filedialog
#import PIL.Image
#import PIL.ImageTk
#from PIL import ImageTk,Image  
import io
import tkinter as tk
root = tk.Tk()
names=[]
numbers=[]
lst=[]
text_file=""
excel_file="attendance.xlsx"
stu_number=2
global date
#encoding='utf8'
#canvas1 = tk.Canvas(root, width = 300, height = 300)
#canvas1.pack()
def fun ():  
  global conf
  conf.pack_forget()
  date = dat.get()
#    label1 = tk.Label(root, text= 'Hello World!', fg='green', font=('helvetica', 12, 'bold'))
#    canvas1.create_window(150, 200, window=label1)  
  if(text_file == ""):
      return
  f = io.open(text_file, "r",encoding='utf8')
  for l in f:
      name_i_f=l.find("From")+6
      name_i_l = l.find(':',name_i_f)
      num_i_f=l.find("201",name_i_l+1)
      number=l[num_i_f:num_i_f+10]
      the_name=l[name_i_f:name_i_l]
      if(number.isdecimal):
          if(len(number)==10):
              if(number in numbers):
                  continue
              else:
                      names.append(the_name)
                      numbers.append(number)
                      lst.append([the_name,number])
  writer = pd.ExcelWriter("log.xlsx", engine = 'xlsxwriter')
  dfw = pd.DataFrame.from_dict({'name':names,'stu_number':numbers })
  dfw.to_excel(writer,sheet_name = "att", header=False, index=False)
  writer.save()
  writer.close()
  df_attend = pd.read_excel (r''+str(excel_file),header=None) 
#[][2]stu_number
  cols={}
  temp_lst=[]
  col_len=len(df_attend.columns)
  for o in range(col_len):
      for i in range(len(df_attend)):
          temp_lst.append(df_attend.iloc[i][o])
      cols[o]=temp_lst
      temp_lst=[]
  temp_lst=[]
  temp=cols[stu_number]
  for i in range(len(temp)):
      if(i == 0):
          temp_lst.append(str(date))
      else:
          if(str(temp[i]) in numbers):
              temp_lst.append(1)
          else:
              temp_lst.append(0)
  cols[col_len]=temp_lst
  writer = pd.ExcelWriter("attendance.xlsx", engine = 'xlsxwriter')
  dfw = pd.DataFrame.from_dict(cols)
  dfw.to_excel(writer,sheet_name = "attendance", header=False, index=False)
  writer.save()
  writer.close()
  conf=tk.Label(root,text="Attendance Report has been Generated :)", font='Helvetica 9 bold',fg="Dark Green",bg=bg_color)
  conf.pack()
#####################################end fun
def txt_browser():
    global text_file
    try:
        text_file = filedialog.askopenfilename(initialdir =  "/", title = "Select A File", filetype =(("Text files","*.txt"),("all files","*.*")) )
#    text_file=text_file.replace("/","//")
    except:
        print("file not found!")
    if(text_file != ""):
        worning_txt.pack_forget()
#    print(text_file)
def excel_browser():
    global excel_file
    excel_file = filedialog.askopenfilename(initialdir =  "/", title = "Select A File", filetype =(("Excel files","*.xlsx .csv"),("all files","*.*")) )
    if(excel_file != ""):  
        if(str(excel_file).find(".csv")):
            read_file = pd.read_csv (r''+excel_file,encoding='utf-8',header=None)
            read_file.to_excel (r'attendance.xlsx', index = None, header=False)
            excel_file="attendance.xlsx"
#        excel_file =excel_file.replace("/","//")
#    else:
#        excel_file="attendance.xlsx"
#        worning_exc.pack_forget()
#    print(excel_file)
bg_color="light cyan"
root.title("YU Zoom's Attendance")
#root.iconbitmap("Logo.ico")
frame = tk.Frame(root,padx=10,pady=10,bg=bg_color)
frame.pack()
label1=tk.Label(frame,text="Zoom's Chat File", font='Helvetica 9 bold',bg=bg_color)
#txt=tk.Entry(frame)
txt = tk.Button(frame,text='Browse', font='Helvetica 9 bold',command=lambda :txt_browser(),pady=5)
label2=tk.Label(frame,text="Attendance Excel Sheet",bg=bg_color, font='Helvetica 9 bold')
excel= tk.Button(frame,text='Browse', font='Helvetica 9 bold',command=lambda :excel_browser(),pady=5)
button1 = tk.Button(frame,text='Generate Attendance Report', font='Helvetica 9 bold',command=fun,pady=5)
label3=tk.Label(frame,text="Attendance Date", font='Helvetica 9 bold',bg=bg_color)
dat=tk.Entry(frame)
label1.grid(row=1,column=1,sticky="w",padx=5)
txt.grid(row=1,column=2,pady=5,sticky="ew")
label2.grid(row=2,column=1,sticky="w",padx=5)
excel.grid(row=2,column=2,sticky="ew",pady=5)
label3.grid(row=3,column=1,sticky="w",padx=5)
dat.grid(row=3,column=2,pady=5,sticky="ew")
button1.grid(columnspan=2,row=4,column=1,sticky="ew",padx=5)
global worning_txt
worning_txt=tk.Label(root,text="Please Select Zoom’s Chat File to Proceed!", font='Helvetica 9 bold',fg="red",bg=bg_color)
worning_txt.pack()
conf=tk.Label(root,text="", font='Helvetica 9 bold',fg="Dark Green",bg=bg_color)
#canvas = tk.Canvas(root, width = 300, height = 300)      
#canvas.pack()    
#try:
#    img = ImageTk.PhotoImage(Image.open("./Logo.png").resize((150,120),Image.ANTIALIAS))
#    label4=tk.Label(frame,image=img)
#    label4.grid(rowspan=4,row=0,column=0,sticky="w")
#    ph = ImageTk.PhotoImage(Image.open("./yu.png"))
#    root.wm_iconphoto(False,ph)
#except:
#    print("no image")
root.configure(bg=bg_color)


#canvas.create_image(20,20, image=img) 
#global worning_exc
#worning_exc=tk.Label(root,text="Be sh!", font='Helvetica 9 bold',fg="orange",bg=bg_color)
#worning_exc.pack()
root.mainloop()
