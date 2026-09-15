import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk  # pip install pillow
from tkinter.filedialog import asksaveasfilename
import barcode as br
 
from barcode.writer import ImageWriter
from tkinter.colorchooser import askcolor

from tkinter import ttk



import tempfile
import os
from db_AddItems import Database


from tkcalendar import Calendar # calendar
from tkcalendar import DateEntry

import random # billNumber
import time  # date time

from db_AddItemsReal import Database
from dbTotal_TempReal import DatabaseTotal
from dbSalesReportReal import DatabaseSales
from dborder import DatabaseOrder
from dbWarehouse import warehouse


class FirstPage(tk.Frame):
     def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        load = Image.open("login1.png")
        photo = ImageTk.PhotoImage(load)
        label = tk.Label(self, image=photo)
        label.image=photo
        label.place(x=0,y=0)
        lblTitle=tk.Label(self,text="ABK Login System", font=("arial", 45, 'bold'),fg="black",bg="#fcc203")
        lblTitle.place(x=450,y=50)
        bordercolor=tk.Frame(self,bg="black",width=800,height=400)
        bordercolor.place(x=300,y=200)
        mainframe=tk.Frame(bordercolor, bg="#fcc203",width=800, height=400)
        mainframe.pack(padx=20,pady=20)
        text1 = tk.StringVar()
        text2 = tk.StringVar()
        L1 = tk.Label(mainframe,text="Username", font=("arial", 30, "bold"),bg="#fcc203") 
        L1.place (x=100,y=50)
        L2=tk.Label(mainframe,text="Password", font=("arial", 30, "bold"),bg="#fcc203") 
        L2.place(x=100,y=150)
        T1=tk.Entry(mainframe,width=12,bd=2, font=("arial", 30),textvariable=text1)
        T1.place(x=400, y=50)
        T2=tk.Entry(mainframe,width=12,bd=2, font=("arial", 30),show="*",textvariable=text2)
        T2.place(x=400,y=150)
        
          

          

# ----------------------------- Old form -----------------------


     #     border = tk.LabelFrame(self, text='Login', bg='ivory', bd = 10, font=("Arial", 20))
     #     border.pack(fill="both", expand="yes", padx = 150, pady=150)
#
     #     L1 = tk.Label(border, text="Username", font=("Arial Bold", 17), bg='ivory')
     #     L1.place(x=50, y=20)
     #     T1 = tk.Entry(border, width = 30, bd = 5)
     #     T1.place(x=180, y=20)
#
     #     L2 = tk.Label(border, text="Password", font=("Arial Bold", 17), bg='ivory')
     #     L2.place(x=50, y=80)
     #     T2 = tk.Entry(border, width = 30, show='*', bd = 5)
     #     T2.place(x=180, y=80)
        
        def verify():
            if T1.get() == "admin" and T2.get() == "123":
               text1.set('')
               text2.set('')
               controller.show_frame(SecondPage)
            else: 
               try:
                    with open("credential.txt", "r") as f:
                         info = f.readlines()
                         i  = 0
                         for e in info:
                             u, p =e.split(",")
                             if u.strip() == T1.get() and p.strip() == T2.get():
                                 # -------------------  Login Design --------------------
                                 controller.show_frame(SixPage)
                                 i = 1
                                 break
                         if i==0:
                              messagebox.showinfo("Error", "Please provide correct username and password!!")
               except:
                      
                    messagebox.showinfo("Error", "Please provide correct username and password!!")
        tk.Button(mainframe,text="Login",height="2",font=("Arial Bold", 15),width="18",bg="#ed3833",fg="white",bd=0,command=verify).place(x=150,y=250)
        tk.Button(mainframe,text="Exit",height="2",font=("Arial Bold", 15),width="18",bg="#00bd56",fg="white",bd=0,command=self.destroy).place(x=450,y=250)
          #B1 = tk.Button(border, text="Submit", font=("Arial", 15), command=verify)
          #B1.place(x=320, y=115)
        
        

# ------------------------------------  Main Form (second page)  ------------------------------------





class SecondPage(tk.Frame):
     def __init__(self, parent, controller):
          tk.Frame.__init__(self, parent)
          
          # -------------------------  Background Photo Added --------------------------   
          #load = Image.open("el.jpg")
          #photo = ImageTk.PhotoImage(load)
          #label = tk.Label(self, image=photo)
          #label.image=photo
          #label.place(x=0,y=0)
          load = Image.open("shop1.jpg")
          photo = ImageTk.PhotoImage(load)
          label = tk.Label(self, image=photo)
          label.image=photo
          label.place(x=0,y=0)
          # ------------- Design Layout --------------------
          entries_frame = tk.Frame(self, bg="#a8f0a8")
          entries_frame.pack(fill="x")
          lbltitle = tk.Label(entries_frame, text="ABK Mini Mart", font=("Calibri", 55, "bold"), fg="Black", bg="#a8f0a8")
          lbltitle.grid(row=0, column=2, padx=315, pady=5)      
          lblDashboard = tk.Label(entries_frame,text="Dashboard",font=("Calibri", 17, "bold"), fg="white",bg="#0366fc")
          lblDashboard.grid(row=0, column=1, padx=10, pady=20)
          btnNotification = tk.Button(entries_frame,text="Notification",font=("Calibri", 17, "bold"), fg="white",bg="#0366fc")
          btnNotification.grid(row=0, column=3,padx=0,pady=20)
          # image photo
          # -------------------------- Side Frame -------------------- 
          # 2 frame 
          body_frame = tk.Frame(self,highlightbackground="black", highlightthickness=13,bg='darkblue') # left side
          body_frame.pack(fill="y",side="left",expand=False)
          color = 'blue'
          
          lbldasboard = tk.Button(body_frame, text="Bill", font=("Calibri", 17, "bold"), fg="yellow",bg=color,width=15,command=lambda: controller.show_frame(SixPage))
          lbldasboard.grid(row=0, column=0, padx=10, pady=16)
          btnAddItem = tk.Button(body_frame, text="AddItem", font=("Calibri", 17, "bold"), fg="yellow",bg=color,width=15,command=lambda: controller.show_frame(FivePage))
          btnAddItem.grid(row=1, column=0, padx=10, pady=16)
          
          btnBarcode = tk.Button(body_frame, text="Barcode", font=("Calibri", 17, "bold"), fg="yellow",bg=color,width=15,command=lambda: controller.show_frame(FourthPage))
          btnBarcode.grid(row=2, column=0, padx=10, pady=16)
          
          lblAddUser = tk.Button(body_frame, text="AddUser", font=("Calibri", 17, "bold"), fg="yellow",bg=color,width=15,command=lambda: controller.show_frame(SevenPage))
          lblAddUser.grid(row=3, column=0, padx=10, pady=16)

          lblWarehouse = tk.Button(body_frame, text="Warehouse", font=("Calibri", 17, "bold"), fg="yellow",bg=color,width=15,command=lambda: controller.show_frame(EightPage))
          lblWarehouse.grid(row=4, column=0, padx=10, pady=16)

          lblNotification = tk.Button(body_frame, text="Notification", font=("Calibri", 17, "bold"), fg="yellow",bg=color,width=15)
          lblNotification.grid(row=5, column=0, padx=10, pady=16)

          lblReport = tk.Button(body_frame, text="Report", font=("Calibri", 17, "bold"), fg="yellow",bg=color,width=15,command=lambda: controller.show_frame(NinePage))
          lblReport.grid(row=6, column=0, padx=10, pady=16)
# ---      ------------------------  TOP  right ----------------------
# 3 frame
          top_right_frame = tk.Frame(self, bg="#ddfa4b",highlightbackground="White", highlightthickness=13) # left side
          top_right_frame.pack(fill="x",side="top",expand=False)
          color = 'black'
          color1 = "white"
          lbl3Sales = tk.Label(top_right_frame, text="Today Sales", font=("Calibri", 24, "bold"), fg="red",bg="#ddfa4b",width=10)
          lbl3Sales.grid(row=0, columnspan=2, padx=5, pady=8)
          lbl3SalesSumm = tk.Label(top_right_frame, text="Sales Summary", font=("Calibri", 18, "bold"), fg="green",bg="#ddfa4b",width=15)
          lbl3SalesSumm.grid(row=1,column=0,columnspan=4, padx=5, pady=8)
          lbl3Sales = tk.Label(top_right_frame, text="Order Summary", font=("Calibri", 18, "bold"), fg="blue",bg="#ddfa4b",width=15)
          lbl3Sales.grid(row=1, column=5,columnspan=4, padx=5, pady=8)
          
          # ------------
          btnSales = tk.Button(top_right_frame, text="ရောင်းရငွေ", font=("Calibri", 10, "bold"), fg="Black",bg=color1,width=13,relief="groove",border=5)
          btnSales.grid(row=2, column=0, padx=20, pady=10)
          txtSales = tk.Entry(top_right_frame,font=("Calibri", 12, "bold"), fg="black",bg=color1,width=10,relief="groove",border=3)
          txtSales.grid(row=2, column=1, padx=20, pady=10)
          btnCosts = tk.Button(top_right_frame, text="အခြားကုန်ငွေ", font=("Calibri", 10, "bold"), fg="black",bg=color1,width=13,relief="groove",border=5)
          btnCosts.grid(row=2, column=2, padx=20, pady=10)
          txtCosts = tk.Entry(top_right_frame,font=("Calibri", 12, "bold"), fg="black",bg=color1,width=10,relief="groove",border=5)
          txtCosts.grid(row=2, column=3, padx=20, pady=10)
          btnOrderCost = tk.Button(top_right_frame, text="Orderကုန်ကျငွေ", font=("Calibri", 10, "bold"), fg="black",bg=color1,width=13,relief="groove",border=5)
          btnOrderCost.grid(row=2, column=4, padx=20, pady=10)
          txtSales = tk.Entry(top_right_frame,font=("Calibri", 12, "bold"), fg="black",bg=color1,width=10,relief="groove",border=5)
          txtSales.grid(row=2, column=5, padx=20, pady=10)
          btnOrder = tk.Button(top_right_frame, text="Orderအရေအတွက်", font=("Calibri", 10, "bold"), fg="black",bg=color1,width=15,relief="groove",border=5)
          btnOrder.grid(row=2, column=6, padx=20, pady=10)
          txtSales = tk.Entry(top_right_frame,font=("Calibri", 12, "bold"), fg="black",bg=color1,width=10,relief="groove",border=5)
          txtSales.grid(row=2, column=7, padx=20, pady=10)
#----      -------------------------------------------------------------
# 4 f      rame 
          bottom_left_frame = tk.Frame(self, bg="#6bfa75",highlightbackground="White", highlightthickness=13) # left side
          bottom_left_frame.pack(fill="y",side="left",expand=False)
          color = 'black'
          lbldasboard = tk.Label(bottom_left_frame, text="Best Sales", font=("Calibri", 17, "bold"), fg="black",bg="#d3dbd3",width=30)
          lbldasboard.grid(row=0, column=0, padx=10, pady=18)
# ___      __________________________________
#    
# 5 f      rame 
          bottom_right_frame2 = tk.Frame(self, bg="#04c404",highlightbackground="White", highlightthickness=13) # left side
          bottom_right_frame2.pack(fill="x",side="top",expand=False)
          color = 'black'
          lbldasboard = tk.Label(bottom_right_frame2, text="Order", font=("Calibri", 20, "bold"), fg="white",bg="#04c404",width=40)
          lbldasboard.grid(row=0, column=0, padx=110, pady=22)

#class SecondPage(tk.Frame):
#     def __init__(self, parent, controller):
#          tk.Frame.__init__(self, parent)
#          
#          # -------------------------  Background Photo Added --------------------------   
#          #load = Image.open("d:/0.work/Yangon Project/ABK Mart/example prg/custom tkinter 1/custom_tkinter_login-master/Google__G__Logo.svg.webp")
#          #photo = ImageTk.PhotoImage(load)
#          #label = tk.Label(self, image=photo)
#          #label.image=photo
#          #label.place(bordermode="outside", height=100, width=100)
#          
#          # ------------- Design Layout --------------------
#          colorF21="#a8f0a8"
#          entries_frame = tk.Frame(self, bg=colorF21,relief="sunken",border=15)
#          entries_frame.pack(fill="x")
#          lbltitle = tk.Label(entries_frame, text="ABK Mini Mart", font=("Calibri", 55, "bold"), fg="#f56942",bg=colorF21)
#          lbltitle.grid(row=0, column=2, padx=380, pady=20)      
#          lblDashboard = tk.Label(entries_frame,text="Dashboard",font=("Calibri", 17, "bold"), fg="white",bg=colorF21)
#          lblDashboard.grid(row=0, column=1, padx=10, pady=20)
#          btnNotification = tk.Button(entries_frame,text="Notification",font=("Calibri", 17, "bold"), fg="white",bg="#f5e942")
#          btnNotification.grid(row=0, column=3,padx=0,pady=20)
#          # image photo
#          # -------------------------- Side Frame -------------------- 
#          # 2 frame 
#          body_frame = tk.Frame(self, bg="#42f58d") # left side
#          body_frame.pack(fill="y",side="left",expand=False,pady=20)
#          color = "#f57b42"
#          lbldasboard = tk.Button(body_frame, text="Bill", font=("Calibri", 17, "bold"), fg="yellow",bg=color,width=15,command=lambda: controller.show_frame(SixPage))
#          lbldasboard.grid(row=0, column=0, padx=10, pady=17)
#          btnAddItem = tk.Button(body_frame, text="AddItem", font=("Calibri", 17, "bold"), fg="yellow",bg=color,width=15,command=lambda: controller.show_frame(FivePage))
#          btnAddItem.grid(row=1, column=0, padx=10, pady=17)
#          
#          btnBarcode = tk.Button(body_frame, text="Barcode", font=("Calibri", 17, "bold"), fg="yellow",bg=color,width=15,command=lambda: controller.show_frame(FourthPage))
#          btnBarcode.grid(row=2, column=0, padx=10, pady=17)
#          
#          lblAddUser = tk.Button(body_frame, text="AddUser", font=("Calibri", 17, "bold"), fg="yellow",bg=color,width=15,command=lambda: controller.show_frame(SevenPage))
#          lblAddUser.grid(row=3, column=0, padx=10, pady=17)
#
#          lblWarehouse = tk.Button(body_frame, text="Warehouse", font=("Calibri", 17, "bold"), fg="yellow",bg=color,width=15,command=lambda: controller.show_frame(EightPage))
#          lblWarehouse.grid(row=4, column=0, padx=10, pady=17)
#
#          lblNotification = tk.Button(body_frame, text="Notification", font=("Calibri", 17, "bold"), fg="yellow",bg=color,width=15)
#          lblNotification.grid(row=5, column=0, padx=10, pady=17)
#
#          lblReport = tk.Button(body_frame, text="Report", font=("Calibri", 17, "bold"), fg="yellow",bg=color,width=15)
#          lblReport.grid(row=6, column=0, padx=10, pady=17)
## ---      ------------------------  TOP  right ----------------------
## 3 f      rame
#          top_right_frame = tk.Frame(self, bg="#42adf5",highlightbackground="White", highlightthickness=13) # left side
#          top_right_frame.pack(fill="x",side="top",expand=False,pady=20)
#          color = 'black'
#          color1 = "white"
#          lbl3Sales = tk.Label(top_right_frame, text="Today Sales", font=("Calibri", 25, "bold"), fg="yellow",bg=color,width=15)
#          lbl3Sales.grid(row=0, column=0, padx=10, pady=10)
#          lbl3SalesSumm = tk.Label(top_right_frame, text="Sales Summary", font=("Calibri", 20, "bold"), fg="green",bg=color,width=15)
#          lbl3SalesSumm.grid(row=1, column=0, padx=10, pady=10)
#          lbl3Sales = tk.Label(top_right_frame, text="Order Summary", font=("Calibri", 20, "bold"), fg="blue",bg=color,width=15)
#          lbl3Sales.grid(row=1, column=2, padx=10, pady=10)
#          btnSales = tk.Button(top_right_frame, text="ရောင်းရငွေ", font=("Calibri", 15, "bold"), fg="yellow",bg=color,width=15,relief="groove",border=10)
#          btnSales.grid(row=2, column=0, padx=10, pady=10)
#          txtSales = tk.Entry(top_right_frame,font=("Calibri", 15, "bold"), fg="yellow",bg=color1,width=15,relief="groove",border=10)
#          txtSales.grid(row=2, column=1, padx=10, pady=10)
#          btnCosts = tk.Button(top_right_frame, text="အခြားကုန်ငွေ", font=("Calibri", 15, "bold"), fg="yellow",bg=color,width=15,relief="groove",border=10)
#          btnCosts.grid(row=2, column=2, padx=10, pady=10)
#          txtCosts = tk.Entry(top_right_frame,font=("Calibri", 15, "bold"), fg="yellow",bg=color1,width=15,relief="groove",border=10)
#          txtCosts.grid(row=2, column=3, padx=10, pady=10)
#          btnOrderCost = tk.Button(top_right_frame, text="Orderကုန်ကျငွေ", font=("Calibri", 15, "bold"), fg="yellow",bg=color,width=15,relief="groove",border=10)
#          btnOrderCost.grid(row=2, column=4, padx=10, pady=10)
#          txtSales = tk.Entry(top_right_frame,font=("Calibri", 15, "bold"), fg="yellow",bg=color1,width=15,relief="groove",border=10)
#          txtSales.grid(row=2, column=5, padx=10, pady=10)
#          btnOrder = tk.Button(top_right_frame, text="Orderအရေအတွက်", font=("Calibri", 15, "bold"), fg="yellow",bg=color,width=15,relief="groove",border=10)
#          btnOrder.grid(row=2, column=6, padx=10, pady=10)
#          txtSales = tk.Entry(top_right_frame,font=("Calibri", 15, "bold"), fg="yellow",bg=color1,width=15,relief="groove",border=10)
#          txtSales.grid(row=2, column=7, padx=10, pady=10)
##----      -------------------------------------------------------------
## 4 f      rame 
#          bottom_left_frame = tk.Frame(self, bg="#42adf5",highlightbackground="White", highlightthickness=13) # left side
#          bottom_left_frame.pack(fill="y",side="left",expand=False,pady=10)
#          color = 'black'
#          lbldasboard = tk.Label(bottom_left_frame, text="Best Sales", font=("Calibri", 17, "bold"), fg="yellow",bg=color,width=35)
#          lbldasboard.grid(row=0, column=0, padx=10, pady=22)
## ___      __________________________________
##    
## 5 f      rame 
#          bottom_right_frame2 = tk.Frame(self, bg="#42adf5",highlightbackground="White", highlightthickness=13) # left side
#          bottom_right_frame2.pack(fill="x",side="top",expand=False,pady=10)
#          color = 'black'
#          lbldasboard = tk.Label(bottom_right_frame2, text="Order", font=("Calibri", 17, "bold"), fg="yellow",bg=color,width=55)
#          lbldasboard.grid(row=0, column=0, padx=110, pady=22)




        
class ThirdPage(tk.Frame):
     def __init__(self, parent, controller):
          tk.Frame.__init__(self, parent)

          self.configure(bg='Tomato')

          Label = tk.Label(self, text="Store some content related to your \n project or what your application made for. \n All the best!!", bg = "orange", font=("Arial Bold", 25))
          Label.place(x=40, y=150)

          Button = tk.Button(self, text="Home", font=("Arial", 15), command=lambda: controller.show_frame(FirstPage))
          Button.place(x=650, y=450)

          Button = tk.Button(self, text="Back", font=("Arial", 15), command=lambda: controller.show_frame(SecondPage))
          Button.place(x=100, y=450)

          Button = tk.Button(self, text="Fourth", font=("Arial", 15), command=lambda: controller.show_frame(FourthPage))
          Button.place(x=300, y=450)


# ----------------------------------- Barcode ----------------------------
# 
# 
#          
class FourthPage(tk.Frame):
     def __init__(self, parent, controller):
          tk.Frame.__init__(self, parent)

          self.configure(bg='#34e8eb')

          barcode_style = {
               "module_width":0.2,
               "module_height":15.0,
               "quiet_zone":2.45,
               "font_size":10,
               "text_distance":5.0,
               "background":"white",
               "foreground":"black",
               "write_text":True,
               "text":"",
          }
          def generate_barcode():
              global img
              img_writer = ImageWriter()
              temp_file = open('_temp_img.png','wb')
              br.Code128(code=txtdigit.get(), writer=img_writer).write(temp_file,barcode_style)
              temp_file.close()
              img = ImageTk.PhotoImage(Image.open('_temp_img.png').resize((300,200),Image.ANTIALIAS))
              barcode_lb.config(image=img)


          # function change color.....
          def set_fg_color():
              color = askcolor()
              if color:
                   btnColor.config(bg=color[1])
                   barcode_style['background']=color[1]
                   generate_barcode() 

          def set_bg_color():
              color = askcolor()
              if color:
                   btnFore.config(bg=color[1])
                   barcode_style['foreground']=color[1]
                   generate_barcode()         

          # humnam reable function (for check box)
          def set_human_readable():
              if human_readable.get():
                  barcode_style['write_text']=True
                  txtadd.config(state=tk.DISABLED)
              else:
                  barcode_style['write_text']=False
                  txtadd.config(state=tk.NORMAL)
              generate_barcode()

          # set_text function
          def set_text():
              barcode_style['text']=txtadd.get()
              generate_barcode()

          # save as function
          def save_barcode():
              file_name=asksaveasfilename()
              if file_name:
                  br.generate(name='code128',code=digit.get(),
                              writer=ImageWriter(),writer_options=barcode_style,
                              output=file_name)
                  

          btnMain = tk.Button(self,command=lambda: controller.show_frame(SecondPage), text="Main", width=15, font=("Calibri", 16, "bold"), fg="black",
                            bg="yellow",
                            bd=0)
          btnMain.pack(pady=5,padx=20)

          main_frame = tk.Frame(self)
          
          digit = tk.Label(main_frame,text='တန်ဖိုးထည့်ပါ',font=('bold,14'))
          digit.pack(pady=13)
          txtdigit = tk.Entry(main_frame,font=('Bold',15),justify=tk.CENTER)
          txtdigit.pack(pady=10)

          generate_btn= tk.Button(main_frame,text='Generate Barcode', font=('Bold',13),bg='blue',fg='white',command=generate_barcode)
          generate_btn.pack(pady=20)

          main_frame.pack(side=tk.TOP,padx=550,pady=20)

          # Add style frame
          style_frame = tk.Frame(self)
          fgframe = tk.Label(style_frame,text="နောက်ခံအရောင်ရွေးချယ်ပါ",font=('Bold',13))
          fgframe.place(x=5,y=10)
          btnFore = tk.Button(style_frame,bg='black',command=set_fg_color)
          btnFore.place(x=10,y=40,width=80,height=25)

          bgframe = tk.Label(style_frame,text="စာလုံးအရောင်ရွေးချယ်ပါ",font=('Bold',13))
          bgframe.place(x=5,y=80)
          btnColor = tk.Button(style_frame,bg='white',command=set_bg_color)
          btnColor.place(x=10,y=110,width=80,height=25)

          #human reable
          human_readable = tk.BooleanVar()
          human_readable.set(True)
          btnread = tk.Checkbutton(style_frame,text="စာသားအဖြစ်ပုံဖော်သည်",font=('Bold,13'),variable=human_readable,command=set_human_readable)
          btnread.place(x=5,y=140)

          addtext = tk.Label(style_frame,text="Add Text",font=('Bold',13))
          addtext.place(x=30,y=170)
          txtadd = tk.Entry(style_frame,font=('Bold',15),state=tk.DISABLED)
          txtadd.place(x=30,y=200)
          # textbox ကနေပြီး button လိုမျိုး function call လို့ရတယ်
          txtadd.bind('<KeyRelease>',lambda e: set_text())


          btnSave = tk.Button(style_frame,text="Save Barcode",bg='#1877f2',fg='white',command=save_barcode)
          btnSave.place(x=0,y=240,width=100,height=30)

          btnDefault = tk.Button(style_frame,text="Default",bg='red',fg='white')
          btnDefault.place(x=130,y=240,width=100,height=30)


          style_frame.pack(side=tk.TOP)
          style_frame.pack_propagate(False)
          style_frame.configure(width=250,height=270)

          # barcode image
          barcode_lb = tk.Label(self,image=None)
          barcode_lb.place(x=400,y=250)


# ----------------------------------- Add Items Data to the store Form ----------
#
#
#
class FivePage(tk.Frame):
     def __init__(self, parent, controller):
          tk.Frame.__init__(self, parent)
          # set background color
          db = Database("real/AddItemsReal.db")
          
          self.config(bg="#2c3e50")
          

          ItemCode = tk.StringVar()
          Name = tk.StringVar()
          Price = tk.StringVar()
          


          # Entries Frame
          colorMain="#15cde6"
          entries_frame = tk.Frame(self, bg=colorMain)
          entries_frame.pack(side="top", fill="x")
          title = tk.Label(entries_frame, text="Adding ItemsCode to Store", font=("Calibri", 25, "bold"), bg=colorMain, fg="white")
          title.grid(row=0,columnspan="6", padx=500, pady=20, sticky="w")

          lblItemCode = tk.Label(entries_frame, text="Item Code", font=("Calibri", 16), bg=colorMain, fg="black")
          lblItemCode.grid(row=1, column=0, padx=5, pady=10, sticky="w")
          txtItemCode= tk.Entry(entries_frame, textvariable=ItemCode, font=("Calibri", 16), width=25)
          txtItemCode.focus()
          txtItemCode.grid(row=1, column=1, padx=5, pady=10, sticky="w")
          

          lblName = tk.Label(entries_frame, text="Name", font=("Calibri", 16), bg=colorMain, fg="black")
          lblName.grid(row=1, column=2, padx=5, pady=10, sticky="w")
          txtName = tk.Entry(entries_frame, textvariable=Name, font=("Calibri", 16), width=25)
          txtName.grid(row=1, column=3, padx=5, pady=10, sticky="w")

          lblPrice = tk.Label(entries_frame, text="Price", font=("Calibri", 16), bg=colorMain, fg="black")
          lblPrice.grid(row=1, column=4, padx=5, pady=10, sticky="w")
          txtPrice = tk.Entry(entries_frame, textvariable=Price, font=("Calibri", 16), width=25)
          txtPrice.grid(row=1, column=5, padx=5, pady=10, sticky="w")

          

          def getData(event):
              selected_row = tv.focus()
              data = tv.item(selected_row)
              global row
              row = data["values"]
              #print(row)
              ItemCode.set(row[1])
              Name.set(row[2])
              Price.set(row[3])
              



          def dispalyAll():
              tv.delete(*tv.get_children())
              for row in db.fetch():
                  tv.insert("", "end", values=row)


          def add_employee():
              if txtItemCode.get() == "" or txtName.get() == "" or txtPrice.get() == "":
                  messagebox.showerror("Erorr in Input", "Please Fill All the Details")
                  return
              db.insert(txtItemCode.get(),txtName.get(), txtPrice.get())
              messagebox.showinfo("Success", "Record Inserted")
              clearAll()
              dispalyAll()
              txtItemCode.focus()



          def update_employee():
              if txtItemCode.get() == "" or txtName.get() == ""  or txtPrice.get() == "":
                  messagebox.showerror("Erorr in Input", "Please Fill All the Details")
                  return
              db.update(row[0],txtItemCode.get(),txtName.get(),txtPrice.get())

              messagebox.showinfo("Success", "Record Update")
              clearAll()
              dispalyAll()


          def delete_employee():
              db.remove(row[0])
              clearAll()
              dispalyAll()


          def clearAll():
              ItemCode.set("")
              Name.set("")
              Price.set("")
              


          btn_frame = tk.Frame(entries_frame, bg=colorMain)
          btn_frame.grid(row=6, column=0, columnspan=7, padx=225, pady=35, sticky="w")
          btnAdd = tk.Button(btn_frame, command=add_employee, text="Add", width=15, font=("Calibri", 16, "bold"), fg="white",
                          bg="#16a085", bd=0).grid(row=0, column=0,padx=15)
          btnEdit = tk.Button(btn_frame, command=update_employee, text="Update Details", width=15, font=("Calibri", 16, "bold"),
                           fg="white", bg="#2980b9",
                           bd=0).grid(row=0, column=1, padx=15)
          btnDelete = tk.Button(btn_frame, command=delete_employee, text="Delete Details", width=15, font=("Calibri", 16, "bold"),
                             fg="white", bg="#c0392b",
                             bd=0).grid(row=0, column=2, padx=15)
          btnClear = tk.Button(btn_frame, command=clearAll, text="Clear Details", width=15, font=("Calibri", 16, "bold"), fg="white",
                            bg="#f39c12",
                            bd=0).grid(row=0, column=3, padx=15)
          btnMain = tk.Button(btn_frame,command=lambda: controller.show_frame(SecondPage), text="Main", width=15, font=("Calibri", 16, "bold"), fg="black",
                            bg="white",
                            bd=0).grid(row=0, column=4, padx=15)
          

          # Table Frame
          tree_frame = tk.Frame(self, bg="#e6a312")
          tree_frame.place(x=235, y=300, width=894, height=320) # အပေါ်တိုး လို့ရတယ်၊ ချ လို့လည်းရတယ်။
          style = ttk.Style()
          style.theme_use("clam")
          style.configure("mystyle.Treeview", font=('Calibri', 13,'italic'),rowheight=50,fg="yellow",background='red')  # Modify the font of the body
          style.configure("mystyle.Treeview.Heading", font=('Calibri', 16),fieldbackground='red') 
           # Modify the font of the headings
          
          tv =ttk.Treeview(tree_frame, columns=(1, 2, 3, 4), style="mystyle.Treeview",selectmode="browse")
          treeScroll = ttk.Scrollbar(tree_frame,orient="vertical")
          treeScroll.configure(command=tv.yview)
          tv.configure(yscrollcommand=treeScroll.set)
          tv.heading("1", text="ID")
          tv.column("1", minwidth=100, width=100, stretch=False)
          tv.heading("2", text="Item Code")
          tv.column("2", minwidth=220, width=250, stretch=False,anchor="center")
          tv.heading("3", text="Name")
          tv.column("3", minwidth=260, width=290, stretch=False,anchor="center")
          tv.heading("4", text="Price")
          tv.column("4", minwidth=220, width=250, stretch=False,anchor="center")
          
          

          tv['show'] = 'headings'
          tv.bind("<ButtonRelease-1>", getData)
          tv.pack(fill="x")
          
          
          dispalyAll()



# ---------------------------------  Bill Form ------------------------
#
#

class SixPage(tk.Frame):
     def __init__(self, parent, controller):
          tk.Frame.__init__(self, parent)
          


          
          #self.configure('Billing Manangement System')

          # full screen size 

          #w, h = self.winfo_screenwidth(), self.winfo_screenheight()
          #self.config("%dx%d+0+0" % (w, h))

          # =------------------ create database --------------
          db = Database("real/AddItemsReal.db")
           #ချိတ်ထားတဲ့ database file ရဲ့ classname ကို ခေါ်ပြီး db ဆောက်
           # database for total ဘောက်ချာထဲထည့်ဖို့
          dbTotal = DatabaseTotal("real/AddTotalReal.db")
          # database for နေ့စဉ် စာရင်းပြုစုဖို့..
          dbSales = DatabaseSales("real/SalesReport.db")

          # Update Quantity
          dbWarehouse = warehouse("real/Warehouse.db")

          bg_color='#2D9290'
          color1 = 'black'

          # ----------------------------- 
          #   discount txt , tax txt ထည့်ဖို့ကျန်နေသေးတယ်


          #=====================variables===================
          

          ItemCode = tk.StringVar()
          
          TotalVar = tk.StringVar()
          
          Cashback = tk.StringVar()

          Quantity = tk.StringVar()

          start = "\033[1m"
          end = "\033[0;0m"

          # ===========Function===============
          def total():
              #if Bread.get()==0 and Wine.get()==0 and Rice.get()==0 and Gal.get()==0:
              #    messagebox.showerror('Error','Please select number of quantity')
              #else:
              #    b=Bread.get()
              #    w=Wine.get()
              #    r=Rice.get()
              #    g=Gal.get()
          #
              #    t=float(b*1.89+w*8.99+r*2.10+g*4.50)
              #    Total.set(b + w + r + g)
              #    total_cost.set('$ ' + str(round(t, 2)))
          #
              #    cb.set('$ '+str(round(b * 1.89, 2)))
              #    cw.set('$ '+str(round(w*8.99,2)))
              #    cr.set('$ '+str(round(r*2.10,2)))
              #    cg.set('$ '+str(round(g*4.50,2)))
              total=0
              values= dbTotal.fetch()
              temp = convertTurpleInt(values) # textarea မှာ string ပဲပြချင်လို့
              Cash = int(txtCashback.get())
              total = str(temp - Cash)
              
              #messagebox.showinfo("Result",total)
              #print(total)
              textarea.insert("end",'******************************************************\n')
              textarea.insert("end",'    ''CashBack:\t\t\t'+txtCashback.get()+'\n') 
              textarea.insert("end",'******************************************************\n')
              textarea.insert("end",'    ''Total Cost:\t\t\t'+total+'\n') # strig ပြောင်းပြီးရေးထားပါတယ်
              textarea.insert("end",'******************************************************\n')
              
              # decrease quantity form warehouse
              #item = txtItems.get()
              #decItem=dbTotal.DecItem(item)
              #
              
              #TotalVar.set(item)
              #dbWarehouse.updateQuantity(item,realQun)
              
              
              item = b
              decItem=dbTotal.DecItem(item)
              #TotalVar.set(decItem)
              mainQun = dbWarehouse.getQuntity(item)
              realQun = convertTurpleInt(mainQun)-convertTurpleInt(decItem)
              dbWarehouse.updateQuantity(b,realQun)   


          def receipt():
              # txtbox ကို ထည့်တာ ပြန်ပြင်ထားတဲ့ CODE 

              #items =db.getItemsCodePrice(txtItems.get())
              #if items==None:
              #    messagebox.showerror("Error","This Items code is not valid")
              #else:
              #    values = convertTurple(items)
              #    values= int(values)
              #    names = db.getItemsCodeName(txtItems.get())
              #    names = convertTurple(names)
              ## condition for non itemcoe ---------------
              #    textarea.delete(1.0,END)
              #textarea.insert(END,' Items\tNumber of Items\t  Cost of Items\n')
                  #textarea.insert(END,f'\nBread\t\t {items} \t  {items}')


                  #x=random.randint(100,10000)
                  #billnumber='BILL'+str(x)
                  #date=time.strftime('%d/%m/%Y')

                  # အကုန်ပြီးမှ bill form ထဲ့ထည့်ရမဲ့အပိုင်း 
                  #textarea.insert(END,'    ' 'Receipt Ref:\t\t'+billnumber+'\t\t'+date+'\n')
                  #textarea.insert(END,'*******************************************************************\n')
                  #textarea.insert(END,'    ''Items:\t\t Cost Of Items(Kyats)\n')
                  #textarea.insert(END,'*******************************************************************\n')

                  #items = {txtItems.get():{"name":names,"price":values},}

                  # ---------------  total form

                  dataAll = db.fetch()


                  #total_bill = txtItems.get()


                      #textarea.insert(END,'    ' 'Receipt Ref:\t\t'+billnumber+'\t\t'+date+'\n')
                      #textarea.insert(END,'*******************************************************************\n')
                      #textarea.insert(END,'    ''Items:\t\t Cost Of Items(Kyats)\n')
                      #textarea.insert(END,'*******************************************************************\n')
                      #print("Scan your barcode")
                  value = txtItems.get()
                  name = ''
                  quan = txtQty.get()

                  for i in range(len(dataAll)):
                      if value==dataAll[i][1]:
                          amount = int(quan)* int(dataAll[i][3])
                          amount = str(amount)
                             #print("Item found",dataAll[i][2],"is",dataAll[i][8])
                          textarea.insert("end",dataAll[i][2]+'\t'+txtQty.get()+'\t'+dataAll[i][3]+'\t'+amount+'\n')
                          name = dataAll[i][2]
                          price = int(dataAll[i][3])
                          #Quantity.set(int(dataAll[i][8]))


                              #messagebox.showinfo("Result",total_bill)
                      #else:
                      #    messagebox.showerror("Error","This item doesn't not exist")
                      #    break
                  getItemcode(value)
                  totalShow(value,name,price,txtQty.get(),amount)
                  ItemCode.set('')
                  
          def getItemcode(a):
              global b
              b = a
              return b
          # -------------------------------------------            
          def totalShow(itemCode,name,price,qty,amount):
                seller = combo.get()
                dbTotal.insert(itemCode,name,amount)
                bill = billnumber
                dbTotal.insertBill(bill,itemCode,name,qty,price,amount,seller)
              #Quantity.set(total)
          #-------------------------- database for Bill Number Search -------
          
          # ---------------  Receipt Button -----------
          #             
          def Receipt_Button():
               global billnumber
               x=random.randint(100,10000)
               billnumber='BILL'+str(x)
               date=time.strftime('%d/%m/%Y')
              #အကုန်ပြီးမှ bill form ထဲ့ထည့်ရမဲ့အပိုင်း 
               textarea.insert("end",'\t'+"       အောင်ဘုန်းခန့် MiniMart"+ '\n')
               textarea.tag_add("here", "1.0", "2.5")
          #textarea.tag_add("start", "1.8", "1.13")
               textarea.tag_config("here", background="yellow", foreground="blue",font='arial 15 bold')
               #textarea.tag_config("start", background="black", foreground="green")
       
               textarea.insert("end",'**************     Ph:09777775706      *************\n')
               textarea.insert("end",'    ' 'Receipt Ref:\t\t'+billnumber+'\t'+date+'\n')
               textarea.insert("end",'******************************************************\n')
               textarea.insert("end",'    ''Items:\t\t Qty \t Price\t Amount\n')
               textarea.insert("end",'******************************************************\n')
               return billnumber    

                  #if txtItems.get()!='0':
                  #    textarea.insert(END,f'\n{names}\t\t {values}  \n')

              #textarea.insert(END,f'\n\nWine\t\t{Wine.get()}\t  {cw.get()}')
              #textarea.insert(END,f'\n\nRice\t\t{Rice.get()}\t  {cr.get()}')
              #textarea.insert(END,f'\n\nMilk\t\t{Gal.get()}\t  {cg.get()}')
              #textarea.insert(END, f"\n\n================================")
              #textarea.insert(END,f'\nTotal Price\t\t{Total.get()}\t{total_cost.get()}')
              #textarea.insert(END, f"\n================================")

          # convert turple to string
          def convertTurple(tup):
              str=''
              for item in tup:
                  str = str + item
              return str

          # conver turple to int
          def convertTurpleInt(tup):
              str=0
              for item in tup:
                  str = str + item
              return str

          #-------------------------------------------

          def print():
              q=textarea.get('1.0','end-1c')
              filename=tempfile.mktemp('.txt')
              open(filename,'w',encoding="utf-8").write(q) # using unicode - encoding="utf-8" ထည့်ပေးရမယ်။
              os.startfile(filename,'Print')


          def reset():
              textarea.delete(1.0,"end")

              # ---------------------  Create Sales Total Databases
              billnumber = Receipt_Button()
              total=0
              values= dbTotal.fetch()
              total = str(convertTurpleInt(values)) 
              date = cal.get_date()
              seller = combo.get()
              dbSales.insert(date,total,billnumber,seller)
              dbTotal.remove()
              #--------------- txtTotal add--------------------
              a = dbSales.Total(date)
              a=str(convertTurpleInt(a)) # str
              TotalVar.set(a)


          # ---------------------------------------    
          def exit():
              if messagebox.askyesno('Exit','Do you really want to exit'):
                  self.destroy()

          title=tk.Frame(self,bd=8,bg='#d6b22d',relief="groove")
          title.place(x=0,y=0,width=1355,height=70)

          title1=tk.Label(title,text="Billing Manangement System",bd=7,bg='#d6b22d',fg='white',font=('times new roman', 28 ,'bold'))
          title1.pack()



          # -------------------- လေ့လာစရာအသစ် ------------
          # cursor ကျနေစေချင်လို့ပါ။ 
          # takefocus က tag နဲ့ရွှေ့လို့မရအောင်ပိတ်တာပါ။ 

          # color အဝါ #d6b22d
          #       အနီ  red4
          #       အစိမ်း original


          #===============Product Details=================
          F1 = tk.LabelFrame(self, text='Product Details', font=('times new romon', 18, 'bold'), fg='gold',bg='green',bd=15,relief="ridge")
          F1.place(y=70,width=870,height=260)

          #=====================Heading==========================
          lblDate=tk.Label(F1, text='Date', font=('Helvetic',22, 'bold'), fg=color1,bg=bg_color)
          lblDate.grid(row=0,column=0,padx=10,pady=20)
          # takefocus က tag နဲ့ရွှေ့လို့မရအောင်ပိတ်တာပါ။
          # hide date format.....
          cal = DateEntry(F1, width=15,background='darkblue', foreground='white', borderwidth=2,font=('Helvetic',17, 'bold'),showweeknumbers=False,
                          showothermonthdays=False)
          cal.grid(row=0,column=1,padx=10,pady=20)
          #txtDate=Entry(F1,font='arial 15 bold',relief=SUNKEN,bd=7,takefocus=False) # textvariable=Wine,justify=CENTER)
          #txtDate.grid(row=0,column=1,padx=10,pady=20)

          lblName=tk.Label(F1, text='Name', font=('Helvetic',20, 'bold'), fg=color1,bg=bg_color)
          lblName.grid(row=0,column=2,padx=10,pady=15)

          combo = ttk.Combobox(F1,width=15,font=('Helvetic',20),
            state="readonly",
            values=["YMTZ", "CMP", "TTS"]
          )

          combo.grid(row=0,column=3)

          lblQty=tk.Label(F1, text='Quantity', font=('Helvetic',20, 'bold'), fg=color1,bg=bg_color)
          lblQty.grid(row=1,column=0,padx=10,pady=15)
          txtQty=tk.Entry(F1,font='arial 15 bold',relief="sunken",bd=7,takefocus=False,textvariable=Quantity)#,textvariable=Wine,justify=CENTER)
          txtQty.grid(row=1,column=1,padx=10,pady=15)
          Quantity.set(1)

          lblItems=tk.Label(F1, text='Items', font=('Helvetic',20, 'bold'), fg=color1,bg=bg_color)
          lblItems.grid(row=1,column=2,padx=10,pady=15)
          txtItems=tk.Entry(F1,font='arial 15 bold',relief="sunken",bd=7,takefocus=True,textvariable=ItemCode)#,textvariable=Wine,justify=CENTER)
          txtItems.focus() # cursor ကျနေစေချင်လို့ပါ။ 
          txtItems.bind('<Return>',lambda e: receipt()) # enter
          txtItems.grid(row=1,column=3,padx=10,pady=15)

          lblCashback=tk.Label(F1, text='CashBack', font=('Helvetic',20, 'bold'), fg=color1,bg=bg_color)
          lblCashback.grid(row=2,column=0,padx=10,pady=15)
          txtCashback=tk.Entry(F1,font='arial 15 bold',relief="sunken",bd=7,takefocus=False,textvariable=Cashback)#,textvariable=Wine,justify=CENTER)
          txtCashback.grid(row=2,column=1,padx=10,pady=15)
          Cashback.set(0)

          lblTotal=tk.Label(F1, text='TotalSales', font=('Helvetic',20, 'bold'), fg=color1,bg="white")
          lblTotal.grid(row=2,column=2,padx=10,pady=15)
          txtTotal=tk.Entry(F1,font='arial 15 bold',relief="sunken",bd=7,takefocus=False,textvariable=TotalVar)#,textvariable=Wine,justify=CENTER)
          txtTotal.grid(row=2,column=3,padx=10,pady=15)

          




          #===============Calculator============
          operator='' #7+9
          def buttonClick(numbers): #9
              global operator
              operator=operator+numbers
              calculatorField.delete(0,"end")
              calculatorField.insert("end",operator)

          def clear():
              global operator
              operator=''
              calculatorField.delete(0,"end")

          def answer():
              global operator
              result=str(eval(operator))
              calculatorField.delete(0,"end")
              calculatorField.insert(0,result)
              operator=''

          calculatorFrame=tk.Frame(self,bd=13,relief="ridge",bg='red4')
          calculatorFrame.place(y=340,width=870,height=260)

          calculatorField=tk.Entry(calculatorFrame,font=('arial',17,'bold'),width=28,bd=5,relief="ridge")
          calculatorField.grid(row=0,column=0,columnspan=4,pady=10)

          button7=tk.Button(calculatorFrame,text='7',font=('arial',13,'bold'),fg='yellow',bg='red4',bd=6,width=6,command=lambda:buttonClick('7'))
          button7.grid(row=1,column=0)

          button8=tk.Button(calculatorFrame,text='8',font=('arial',13,'bold'),fg='yellow',bg='red4',bd=6,width=6,command=lambda:buttonClick('8'))
          button8.grid(row=1,column=1)

          button9=tk.Button(calculatorFrame,text='9',font=('arial',13,'bold'),fg='yellow',bg='red4',bd=6,width=6,command=lambda:buttonClick('9'))
          button9.grid(row=1,column=2)

          buttonPlus=tk.Button(calculatorFrame,text='+',font=('arial',13,'bold'),fg='yellow',bg='red4',bd=6,width=10,command=lambda:buttonClick('+'))
          buttonPlus.grid(row=1,column=3)

          button4=tk.Button(calculatorFrame,text='4',font=('arial',13,'bold'),fg='yellow',bg='red4',bd=6,width=6
                         ,command=lambda:buttonClick('4'))
          button4.grid(row=2,column=0)

          button5=tk.Button(calculatorFrame,text='5',font=('arial',13,'bold'),fg='red4',bg='white',bd=6,width=6
                         ,command=lambda:buttonClick('5'))
          button5.grid(row=2,column=1)

          button6=tk.Button(calculatorFrame,text='6',font=('arial',13,'bold'),fg='red4',bg='white',bd=6,width=6
                         ,command=lambda:buttonClick('6'))
          button6.grid(row=2,column=2)

          buttonMinus=tk.Button(calculatorFrame,text='-',font=('arial',13,'bold'),fg='yellow',bg='red4',bd=6,width=10
                             ,command=lambda:buttonClick('-'))
          buttonMinus.grid(row=2,column=3)

          button1=tk.Button(calculatorFrame,text='1',font=('arial',13,'bold'),fg='yellow',bg='red4',bd=6,width=6
                         ,command=lambda:buttonClick('1'))
          button1.grid(row=3,column=0)

          button2=tk.Button(calculatorFrame,text='2',font=('arial',13,'bold'),fg='red4',bg='white',bd=6,width=6
                         ,command=lambda:buttonClick('2'))
          button2.grid(row=3,column=1)

          button3=tk.Button(calculatorFrame,text='3',font=('arial',13,'bold'),fg='red4',bg='white',bd=6,width=6
                         ,command=lambda:buttonClick('3'))
          button3.grid(row=3,column=2)

          buttonMult=tk.Button(calculatorFrame,text='x',font=('arial',13,'bold'),fg='yellow',bg='red4',bd=6,width=10
                            ,command=lambda:buttonClick('*'))
          buttonMult.grid(row=3,column=3)

          buttonAns=tk.Button(calculatorFrame,text='Ans',font=('arial',13,'bold'),fg='yellow',bg='red4',bd=6,width=6,
                           command=answer)
          buttonAns.grid(row=4,column=0)

          buttonClear=tk.Button(calculatorFrame,text='Clear',font=('arial',13,'bold'),fg='yellow',bg='red4',bd=6,width=6
                             ,command=clear)
          buttonClear.grid(row=4,column=1)

          button0=tk.Button(calculatorFrame,text='0',font=('arial',13,'bold'),fg='yellow',bg='red4',bd=6,width=6
                         ,command=lambda:buttonClick('0'))
          button0.grid(row=4,column=2)

          buttonDiv=tk.Button(calculatorFrame,text='/',font=('arial',13,'bold'),fg='yellow',bg='red4',bd=6,width=10,
                           command=lambda:buttonClick('/'))
          buttonDiv.grid(row=4,column=3)



          image=Image.open('logo.jpg')
          resized_image= image.resize((70,40), Image.ANTIALIAS)
          pic=ImageTk.PhotoImage(resized_image)

          lblphoto = tk.Label(calculatorFrame,image=pic,bd=6)
          lblphoto.grid(row=0,column=4,padx=180,pady=0,rowspan=2)

          lblProduct = tk.Label(calculatorFrame,text='MSH',font=('forte',20,'bold'),fg='#1958f7',bg='white',bd=6,width=7)
          lblProduct.grid(row=2, column=4,padx=130,pady=10,rowspan=2)

          lblcontact = tk.Label(calculatorFrame,text='contact number-09444418843',font=('forte',15),fg='#1958f7',bg='white',bd=6,width=25)
          lblcontact.grid(row=4, column=4,padx=110,pady=10)



          #=====================Bill areea====================
          F2=tk.Frame(self,relief="groove",bd=18,bg='#d6b22d')
          F2.place(x=870,y=70,width=490,height=635)
          bill_title=tk.Label(F2,text='Receipt',font='arial 15 bold',bd=7,relief="groove",bg='green',fg='red4').pack(fill="x")
          scrol_y=tk.Scrollbar(F2,orient="vertical")
          scrol_y.pack(side="right",fill="y")
          textarea=tk.Text(F2,font='arial 15',yscrollcommand=scrol_y.set,takefocus=False)
          textarea.pack(fill="both")
          scrol_y.config(command=textarea.yview)



          #=====================Buttons========================
          F3 =tk.Frame(self,bg=bg_color,bd=15,relief="ridge")
          F3.place(y=610,width=870,height=90)

          btn1 = tk.Button(F3, text='Login', font='arial 12 bold', padx=5, pady=5, bg='yellow',fg='red',width=7, command=lambda: controller.show_frame(FirstPage),takefocus=False,bd=7)#,command=total
          btn1.grid(row=0,column=0,padx=30,pady=7)
          # receipt button အစားထိုးရန်
          btn2 = tk.Button(F3, text='Total', font='arial 12 bold', padx=5, pady=5, bg='yellow',fg='red',width=7,command=total,takefocus=False,bd=7)
          btn2.grid(row=0,column=1,padx=30,pady=7)

          btn3 = tk.Button(F3, text='Print', font='arial 12 bold', padx=5, pady=5, bg='yellow',fg='red',width=7,command=print,takefocus=False,bd=7)
          btn3.grid(row=0,column=2,padx=30,pady=7)

          btn4 = tk.Button(F3, text='Reset', font='arial 12 bold', padx=5, pady=5, bg='yellow',fg='red',width=7,command=reset,takefocus=False,bd=7)
          btn4.grid(row=0,column=3,padx=30,pady=7)

          btn5 = tk.Button(F3, text='Exit', font='arial 12 bold', padx=5, pady=5, bg='yellow',fg='red',width=7,command=exit,takefocus=False,bd=7)
          btn5.grid(row=0,column=4,padx=30,pady=7)


          # ----------------------- ကြိုပေါ်ချင်လို့
          x=random.randint(100,10000)
          billnumber='BILL'+str(x)
          date=time.strftime('%d/%m/%Y')
          #အကုန်ပြီးမှ bill form ထဲ့ထည့်ရမဲ့အပိုင်း
          # textarea bold form change 
          textarea.insert("end",'\t'+"       အောင်ဘုန်းခန့် MiniMart"+ '\n')
          textarea.tag_add("here", "1.0", "2.5")
          #textarea.tag_add("start", "1.8", "1.13")
          textarea.tag_config("here", background="yellow", foreground="blue",font='arial 15 bold')
          #textarea.tag_config("start", background="black", foreground="green")

          textarea.insert("end",'**************     Ph:09777775706      *************\n')
          textarea.insert("end",'    ' 'Receipt Ref:\t\t'+billnumber+'\t'+date+'\n')
          textarea.insert("end",'******************************************************\n')
          textarea.insert("end",'    ''Items:\t Qty \t Price\t Amount\n')
          textarea.insert("end",'******************************************************\n')
          


# ---------------------------   Seven Page (Add User / Register form) -------------------------------------
        
class SevenPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        self.configure(bg='Tomato')
        def register():
            window = tk.Tk()
            window.resizable(0,0)
            window.configure(bg="deep sky blue")
            window.title("Register User")
            l1 = tk.Label(window, text="Username:", font=("Arial",15), bg="deep sky blue")
            l1.place(x=10, y=10)
            t1 = tk.Entry(window, width=30, bd=5)
            t1.place(x = 200, y=10)
            l2 = tk.Label(window, text="Password:", font=("Arial",15), bg="deep sky blue")
            l2.place(x=10, y=60)
            t2 = tk.Entry(window, width=30, show="*", bd=5)
            t2.place(x = 200, y=60)
            l3 = tk.Label(window, text="Confirm Password:", font=("Arial",15), bg="deep sky blue")
            l3.place(x=10, y=110)
            t3 = tk.Entry(window, width=30, show="*", bd=5)
            t3.place(x = 200, y=110)
        
            def check():
                 if t1.get()!="" or t2.get()!="" or t3.get()!="":
                      if t2.get()==t3.get():
                           with open("credential.txt", "a") as f:
                                f.write(t1.get()+","+t2.get()+"\n")
                                messagebox.showinfo("Welcome","You are registered successfully!!")
                      else:
                           messagebox.showinfo("Error","Your password didn't get match!!")
                 else:
                      messagebox.showinfo("Error", "Please fill the complete field!!")
                 
            b1 = tk.Button(window, text="Register", font=("Arial",15), bg="#ffc22a", command=check,width=10)
            b1.place(x=180, y=170)
        
            window.geometry("470x240+500+220")
            
            
        B2 = tk.Button(self, text="Register", bg = "dark orange", font=("Arial",15),width=15,height=1,relief="groove",border=1, command=register) # chagne code next
        B2.place(x=1250, y=20)
        #,command=lambda: controller.show_frame(SevenPage)
        B3 = tk.Button(self,text="Main Form", bg = "dark orange", font=("Arial",15),width=15,height=1,relief="groove",border=1,command=lambda: controller.show_frame(SecondPage))
        B3.place(x=50,y=20)
          

# ------------------------------------  WareHouse Contorl Form ----------------
#
#
class EightPage(tk.Frame):
     def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
          # set background color
          
            dborder = DatabaseOrder("real/Order.db")
            dbwarehouse = warehouse("real/Warehouse.db")
            

            ItemCode = tk.StringVar()
            Name = tk.StringVar()
            Quantity = tk.StringVar()
            Order_Date = tk.StringVar()
            Coming_Date = tk.StringVar()

            #------------------------------------warehouse-------------------------#
            ItemCode1 = tk.StringVar()
            Name1 = tk.StringVar()
            Catagory = tk.StringVar()
            Quantity1 = tk.StringVar()
            Cost = tk.StringVar()
            Price = tk.StringVar()
            Exp_date = tk.StringVar()
            Alert_date = tk.StringVar()





            # Entries Frame for Order
            entries_frame = tk.Frame(self, bg="#dff168")
            entries_frame.pack(side="top", fill="x")
            title = tk.Label(entries_frame, text="Adding Items Store", font=("Calibri", 25, "bold"), bg="#dff168", fg="black")
            title.grid(row=0,columnspan="6", padx=600, pady=20, sticky="w")
            B3 = tk.Button(self,text="Main Form", bg = "dark orange", font=("Arial",15),width=15,height=1,relief="groove",border=1,command=lambda: controller.show_frame(SecondPage))
            B3.place(x=20,y=20)

            # ----------------- Order form -------------
            orderFrame = tk.Frame(self,bg="#6bb4ff",width=662) # width 
            orderFrame.pack(side="left",fill="y")
            lblOrder = tk.Label(orderFrame, text="Order Form", font=("Calibri", 18,"bold"), bg="#6bb4ff", fg="yellow")
            lblOrder.grid(row=1, column=0, padx=10, pady=10, sticky="w")

            lblItemCode = tk.Label(orderFrame, text="Item Code", font=("Calibri", 16,"bold"), bg="#6bb4ff", fg="black")
            lblItemCode.grid(row=2, column=0, padx=10, pady=10, sticky="w")
            txtItemCode= tk.Entry(orderFrame, textvariable=ItemCode, font=("Calibri", 16), width=30)
            txtItemCode.grid(row=2, column=1, padx=10, pady=10, sticky="w")

            lblName = tk.Label(orderFrame, text="Name", font=("Calibri", 16,"bold"), bg="#6bb4ff", fg="black")
            lblName.grid(row=3, column=0, padx=10, pady=10, sticky="w")
            txtName = tk.Entry(orderFrame, textvariable=Name, font=("Calibri", 16), width=30)
            txtName.grid(row=3, column=1, padx=10, pady=10, sticky="w")

            lblQuantity = tk.Label(orderFrame, text="Quantity", font=("Calibri", 16,"bold"), bg="#6bb4ff", fg="black")
            lblQuantity.grid(row=4, column=0, padx=10, pady=10, sticky="w")
            txtQuantity = tk.Entry(orderFrame, textvariable=Quantity, font=("Calibri", 16), width=30)
            txtQuantity.grid(row=4, column=1, padx=10, pady=10, sticky="w")

            lblOrder = tk.Label(orderFrame, text="Order Date", font=("Calibri", 16,"bold"), bg="#6bb4ff", fg="black")
            lblOrder.grid(row=5, column=0, padx=10, pady=10, sticky="w")
            txtOrder = DateEntry(orderFrame, width=15,background='darkblue', foreground='white', borderwidth=2,font=('Helvetic',9, 'bold'),showweeknumbers=False,
                          showothermonthdays=False)
            txtOrder.grid(row=5, column=1, padx=10, pady=10, sticky="w")

            lblComing = tk.Label(orderFrame, text="Coming Date", font=("Calibri", 16,"bold"), bg="#6bb4ff", fg="black")
            lblComing.grid(row=6, column=0, padx=10, pady=10, sticky="w")
            txtComing = DateEntry(orderFrame, width=15,background='darkblue', foreground='white', borderwidth=2,font=('Helvetic',9, 'bold'),showweeknumbers=False,
                          showothermonthdays=False)
            txtComing.grid(row=6, column=1, padx=10, pady=10, sticky="w")

            #--------------------------warehouse frame------------#
            

            warehouseFrame = tk.Frame(self,bg="#5aefe9",width=1550) # width 
            warehouseFrame.pack(side="left",fill="y",padx=30)

            lblWareHouse = tk.Label(warehouseFrame, text="Warehouse Form", font=("Calibri", 15,"bold"), bg="#5aefe9", fg="yellow")
            lblWareHouse.grid(row=1, column=0, padx=10, pady=10, sticky="w")

            lblItemCode = tk.Label(warehouseFrame, text="Item Code", font=("Calibri", 16, "bold"), bg="#5aefe9", fg="Black")
            lblItemCode.grid(row=2, column=0, padx=10, pady=10, sticky="w")
            txtItemCode1= tk.Entry(warehouseFrame, textvariable=ItemCode1, font=("Calibri", 14), width=17)
            txtItemCode1.grid(row=2, column=1, padx=10, pady=10, sticky="w")

            lblName = tk.Label(warehouseFrame, text="Name", font=("Calibri", 16, "bold"), bg="#5aefe9", fg="Black")
            lblName.grid(row=3, column=0, padx=10, pady=10, sticky="w")
            txtName1 = tk.Entry(warehouseFrame, textvariable=Name1, font=("Calibri", 14), width=17)
            txtName1.grid(row=3, column=1, padx=10, pady=10, sticky="w")

            lblCatagory = tk.Label(warehouseFrame, text="Catagory", font=("Calibri", 16, "bold"), bg="#5aefe9", fg="black")
            lblCatagory.grid(row=2, column=2, padx=10, pady=10, sticky="w")
            txtCatagory= tk.Entry(warehouseFrame, textvariable=Catagory, font=("Calibri", 14), width=17)
            txtCatagory.grid(row=2, column=3, padx=10, pady=10, sticky="w")

            lblQuantity = tk.Label(warehouseFrame, text="Quantity", font=("Calibri", 16, "bold"), bg="#5aefe9", fg="black")
            lblQuantity.grid(row=3, column=2, padx=10, pady=10, sticky="w")
            txtQuantity1 = tk.Entry(warehouseFrame, textvariable=Quantity1, font=("Calibri", 14), width=17)
            txtQuantity1.grid(row=3, column=3, padx=10, pady=10, sticky="w")

            lblCost = tk.Label(warehouseFrame, text="Cost", font=("Calibri", 16, "bold"), bg="#5aefe9", fg="black")
            lblCost.grid(row=4, column=0, padx=10, pady=10, sticky="w")
            txtCost = tk.Entry(warehouseFrame, textvariable=Cost, font=("Calibri", 14), width=17)
            txtCost.grid(row=4, column=1, padx=10, pady=10, sticky="w")

            lblPrice = tk.Label(warehouseFrame, text="Price", font=("Calibri", 16, "bold"), bg="#5aefe9", fg="black")
            lblPrice.grid(row=4, column=2, padx=10, pady=10, sticky="w")
            txtPrice = tk.Entry(warehouseFrame, textvariable=Price, font=("Calibri", 14), width=17)
            txtPrice.grid(row=4, column=3, padx=10, pady=10, sticky="w")

            lblExp = tk.Label(warehouseFrame, text="Exp-Date", font=("Calibri", 16, "bold"), bg="#5aefe9", fg="black")
            lblExp.grid(row=5, column=0, padx=10, pady=10, sticky="w")

            

            txtExp = DateEntry(warehouseFrame, width=15,background='darkblue', foreground='white', borderwidth=2,font=('Helvetic',9),showweeknumbers=False,
                          showothermonthdays=False) # week တွေကို hide 
            txtExp.grid(row=5, column=1, padx=10, pady=10, sticky="w")

            lblAlert = tk.Label(warehouseFrame, text="Alert-Date", font=("Calibri", 16, "bold"), bg="#5aefe9", fg="black")
            lblAlert.grid(row=5, column=2, padx=10, pady=10, sticky="w")
            txtAlert = DateEntry(warehouseFrame, width=15,background='darkblue', foreground='white', borderwidth=2,font=('Helvetic',9),showweeknumbers=False,
                          showothermonthdays=False)
            txtAlert.grid(row=5, column=3, padx=10, pady=10, sticky="w")



            def getData(event):
                selected_row = tv.focus()
                data = tv.item(selected_row)
                global row
                row = data["values"]
                #print(row)
                ItemCode.set(row[1])
                Name.set(row[2])
                Quantity.set(row[3])
                Order_Date.set(row[4])
                Coming_Date.set(row[5])




            def dispalyAll():
                tv.delete(*tv.get_children())
                for row in dborder.fetch():
                    tv.insert("", "end", values=row)


            def add_employee():
                if txtItemCode.get() == "" or txtName.get() == "" or txtQuantity.get() == ""or txtOrder.get() == "" or txtComing.get() == "":
                    messagebox.showerror("Erorr in Input", "Please Fill All the Details")
                    return
                dborder.insert(txtItemCode.get(),txtName.get(),txtQuantity.get(),txtOrder.get(),txtComing.get())
                messagebox.showinfo("Success", "Record Inserted")
                clearAll()
                dispalyAll()



            def update_employee():
                if txtItemCode.get() == "" or txtName.get() == "" or txtQuantity.get() == ""or txtOrder.get() == "" or txtComing.get() == "":
                    messagebox.showerror("Erorr in Input", "Please Fill All the Details")
                    return
                dborder.update(row[0],txtItemCode.get(),txtName.get(),txtQuantity.get(),txtOrder.get(),txtComing.get())
                messagebox.showinfo("Success", "Record Update")
                clearAll()
                dispalyAll()


            def delete_employee():
                dborder.remove(row[0])
                clearAll()
                dispalyAll()


            def clearAll():
                ItemCode.set("")
                Name.set("")
                Quantity.set("")
                Order_Date.set("")
                Coming_Date.set("")

            def MoveOrder():
               # item changed to warehouse
               ItemCode1.set(txtItemCode.get())
               Name1.set(txtName.get())
               Quantity1.set(txtQuantity.get())


               dborder.remove(row[0])


            btn_frame = tk.Frame(orderFrame, bg="#6bb4ff")
            btn_frame.grid(row=7, column=0, columnspan=4, padx=10, pady=10, sticky="w")
            btnAdd = tk.Button(btn_frame, command=add_employee, text="Add", width=15, font=("Calibri", 16, "bold"), fg="white",
                            bg="#16a085", bd=0).grid(row=0, column=0, padx=10 )
            btnEdit = tk.Button(btn_frame, command=update_employee, text="Update Details", width=15, font=("Calibri", 16, "bold"),
                             fg="white", bg="#2980b9",
                             bd=0).grid(row=0, column=1,padx=10)
            btnDelete = tk.Button(btn_frame, command=delete_employee, text="Delete Details", width=15, font=("Calibri", 16, "bold"),
                               fg="white", bg="#c0392b",
                               bd=0).grid(row=0, column=2,padx=10)
            btnClear = tk.Button(btn_frame, command=clearAll, text="Clear Details", width=15, font=("Calibri", 16, "bold"), fg="white",
                              bg="#f39c12",
                              bd=0).grid(row=1, column=1,pady=10)
            btnMove = tk.Button(btn_frame, command=MoveOrder, text="Move Warehouse", width=15, font=("Calibri", 16, "bold"), fg="blue",
                              bg="#f39c12",
                              bd=0).grid(row=1, column=2,pady=10)

            # Table Frame
            tree_frame = tk.Frame(self, bg="#ecf0f1")
            tree_frame.place(x=0, y=520, width=593, height=270)
            style = ttk.Style()
            style.configure("mystyle.Treeview", font=('Calibri', 12),
                            rowheight=50)  # Modify the font of the body
            style.configure("mystyle.Treeview.Heading", font=('Calibri', 14))  # Modify the font of the headings
            tv = ttk.Treeview(tree_frame, columns=(1, 2, 3, 4, 5, 6), style="mystyle.Treeview")
            tv.heading("1", text="ID")
            tv.column("1", minwidth=60, width=60, stretch=False)
            tv.heading("2", text="Item Code")
            tv.column("2", minwidth=100, width=100, stretch=False)
            tv.heading("3", text="Name")
            tv.column("3", minwidth=110, width=110, stretch=False)
            tv.heading("4", text="Quantity")
            tv.column("4", minwidth=100, width=100, stretch=False)
            tv.heading("5", text="Order Date")
            tv.column("5", minwidth=100, width=100, stretch=False)
            tv.heading("6", text="Coming Date")
            tv.column("6", minwidth=120, width=120, stretch=False)

            tv['show'] = 'headings'
            tv.bind("<ButtonRelease-1>", getData)
            tv.pack(fill="x")

            #------------------------------AUDC----------------#WareHOuse
            def getData1(event):
                selected_row = tv1.focus()
                data = tv1.item(selected_row)
                global row1
                row1 = data["values"]
                #print(row)
                ItemCode1.set(row1[1])
                Name1.set(row1[2])
                Catagory.set(row1[3])
                Quantity1.set(row1[4])
                Cost.set(row1[5])
                Price.set(row1[6])
                Exp_date.set(row1[7])
                Alert_date.set(row1[8])




            def dispalyAll1():
                tv1.delete(*tv1.get_children())
                for row in dbwarehouse.fetch():
                    tv1.insert("", "end", values=row)


            def add_employee1():
                if txtItemCode1.get() == "" or txtName1.get() == "" or txtCatagory.get() =="" or txtQuantity1.get() == ""or txtCost.get() ==""or txtPrice.get() ==""or txtExp.get() == "" or txtAlert.get() == "":
                    messagebox.showerror("Erorr in Input", "Please Fill All the Details")
                    return
                dbwarehouse.insert(txtItemCode1.get(),txtName1.get(),txtCatagory.get(),txtQuantity1.get(),txtCost.get(),txtPrice.get(),txtExp.get(),txtAlert.get())
                messagebox.showinfo("Success", "Record Inserted")
                clearAll1()
                dispalyAll1()



            def update_employee1():
                if txtItemCode1.get() == "" or txtName1.get() == "" or txtCatagory.get() =="" or txtQuantity1.get() == ""or txtCost.get() ==""or txtPrice.get() ==""or txtExp.get() == "" or txtAlert.get() == "":
                    messagebox.showerror("Erorr in Input", "Please Fill All the Details")
                    return
                dbwarehouse.update(row1[0],txtItemCode1.get(),txtName1.get(),txtCatagory.get(),txtQuantity1.get(),txtCost.get(),txtPrice.get(),txtExp.get(),txtAlert.get())
                messagebox.showinfo("Success", "Record Update")
                clearAll1()
                dispalyAll1()


            def delete_employee1():
                dbwarehouse.remove(row1[0])
                clearAll1()
                dispalyAll1()


            def clearAll1():
                ItemCode1.set("")
                Name1.set("")
                Catagory.set("")
                Quantity1.set("")
                Cost.set("")
                Price.set("")
                Exp_date.set("")
                Alert_date.set("")



            btn_frame = tk.Frame(warehouseFrame, bg="#5aefe9")
            btn_frame.grid(row=7, column=0, columnspan=4, padx=5, pady=50, sticky="w")
            btnAdd = tk.Button(btn_frame, command=add_employee1, text="Add", width=15, font=("Calibri", 16, "bold"), fg="white",
                            bg="#16a085", bd=0).grid(row=0, column=0,padx=6)
            btnEdit = tk.Button(btn_frame, command=update_employee1, text="Update Details", width=15, font=("Calibri", 16, "bold"),
                             fg="white", bg="#2980b9",
                             bd=0).grid(row=0, column=1, padx=6)
            btnDelete = tk.Button(btn_frame, command=delete_employee1, text="Delete Details", width=15, font=("Calibri", 16, "bold"),
                               fg="white", bg="#c0392b",
                               bd=0).grid(row=0, column=2, padx=6)
            btnClear = tk.Button(btn_frame, command=clearAll1, text="Clear Details", width=15, font=("Calibri", 16, "bold"), fg="white",
                              bg="#f39c12",
                              bd=0).grid(row=0, column=3, padx=6)

            # Table Frame
            tree_frame1 = tk.Frame(self, bg="#ecf0f1")
            tree_frame1.place(x=623, y=520, width=833, height=280)
            style = ttk.Style()
            style.configure("mystyle.Treeview", font=('Calibri', 12),
                            rowheight=50)  # Modify the font of the body
            style.configure("mystyle.Treeview.Heading", font=('Calibri', 13))  # Modify the font of the headings
            tv1 = ttk.Treeview(tree_frame1, columns=(1, 2, 3, 4, 5, 6,7,8,9), style="mystyle.Treeview")
            tv1.heading("1", text="ID")
            tv1.column("1", minwidth=40, width=40, stretch=False)
            tv1.heading("2", text="Item Code")
            tv1.column("2", minwidth=90, width=90, stretch=False)
            tv1.heading("3", text="Name")
            tv1.column("3", minwidth=100, width=100, stretch=False)
            tv1.heading("4", text="Catagory")
            tv1.column("4", minwidth=80, width=80, stretch=False)
            tv1.heading("5", text="Quantity")
            tv1.column("5", minwidth=90, width=90, stretch=False)
            tv1.heading("6", text="Cost")
            tv1.column("6", minwidth=80, width=80, stretch=False)
            tv1.heading("7", text="Price")
            tv1.column("7", minwidth=80, width=80, stretch=False)
            tv1.heading("8", text="Exp-date")
            tv1.column("8", minwidth=85, width=85, stretch=False)
            tv1.heading("9", text="Alert-Date")
            tv1.column("9", minwidth=85, width=85, stretch=False)

            tv1['show'] = 'headings'
            tv1.bind("<ButtonRelease-1>", getData1)
            tv1.pack(fill="x")

            dispalyAll1()
#------------------------------------- Report Form --------------------------
class NinePage(tk.Frame):
     def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)

            self.configure(bg='Tomato')

            

            title = tk.Frame(self,bg="#60adcc",width=30,height=70,highlightbackground="#c3c7c7", highlightthickness=8)
            title.pack(fill=tk.X)
            lbltitle = tk.Label(title, text="Report From", font=("Calibri", 30, "bold"), fg="white", bg="#60adcc")
            lbltitle.place(x=580,y=20)
            B3 = tk.Button(self,text="Main Form", bg = "dark orange", font=("Arial",15),width=15,height=1,relief="groove",border=1,command=lambda: controller.show_frame(SecondPage))
            B3.place(x=20,y=20)

            #----------------------------------------------------------
            firstLine = tk.Frame(self,bg="yellow",width=30,height=370)
            firstLine.pack(fill=tk.X)

            DailyReport = tk.Frame(firstLine,width=500,height=250,bg="white",highlightbackground="#c3c7c7", highlightthickness=8)
            DailyReport.pack(side="left",fill="x",expand=False)
            lbltitle = tk.Label(DailyReport,text="Daily Report", font=("Calibri", 25, "bold"), fg="Black", bg="#c3c7c7")
            lbltitle.place(x=10,y=20)
            lblsales = tk.Button(firstLine, text="Sales", font=("Calibri", 15, "bold"), fg="white",bg="#e69719",width=15)
            lblsales.place(x=10,y=80)
            txtSales = tk.Entry(firstLine,font=("Calibri", 18, "bold"), fg="black",bg="white",width=15,relief="groove",border=5)
            txtSales.place(x=200,y=80)
            lblquantity = tk.Button(firstLine, text="Quantity", font=("Calibri", 15, "bold"), fg="white",bg="#09ad2c",width=15)
            lblquantity.place(x=10,y=140)
            txtquantity = tk.Entry(firstLine,font=("Calibri", 18, "bold"), fg="black",bg="white",width=15,relief="groove",border=5)
            txtquantity.place(x=200,y=140)


            RemaingItems = tk.Frame(firstLine,width=700,height=250,bg="white",highlightbackground="#c3c7c7", highlightthickness=8)
            RemaingItems.pack(side="left",fill="x",expand=True)
            lbltitle = tk.Label(RemaingItems,text="Remaing Item", font=("Calibri", 20, "bold"), fg="Black", bg="#c3c7c7")
            lbltitle.place(x=40,y=15)


            #----------------------------------------------------------
            SecondLine = tk.Frame(self,bg="#fae7e6",width=30,height=370,highlightbackground="#c3c7c7", highlightthickness=8)
            SecondLine.pack(fill=tk.X)

            MonthlyReport = tk.Frame(SecondLine,width=700,height=250,bg="white")
            MonthlyReport.pack(side="left",fill="x",expand=False)
            lbltitle = tk.Label(MonthlyReport,text="Monthly Report", font=("Calibri", 20, "bold"), fg="Black", bg="#c3c7c7")
            lbltitle.place(x=10,y=15)
            lblsales = tk.Button(MonthlyReport, text="Sales", font=("Calibri", 15, "bold"), fg="white",bg="#f54949",width=10)
            lblsales.place(x=10,y=70)
            txtSales = tk.Entry(MonthlyReport,font=("Calibri", 18, "bold"), fg="black",bg="white",width=15,relief="groove",border=5)
            txtSales.place(x=140,y=70)
            lblquantity = tk.Button(MonthlyReport, text="Quantity", font=("Calibri", 15, "bold"), fg="white",bg="#e69719",width=10)
            lblquantity.place(x=350,y=70)
            txtquantity = tk.Entry(MonthlyReport,font=("Calibri", 18, "bold"), fg="black",bg="white",width=15,relief="groove",border=5)
            txtquantity.place(x=500,y=70)
            lbltitle1 = tk.Label(MonthlyReport,text="Best Seller", font=("Calibri", 20, "bold"), fg="Black", bg="#c3c7c7")
            lbltitle1.place(x=10,y=140)
            lblname = tk.Button(MonthlyReport, text="Name", font=("Calibri", 15, "bold"), fg="white",bg="#7f57cf",width=10)
            lblname.place(x=10,y=200)
            txtname = tk.Entry(MonthlyReport,font=("Calibri", 18, "bold"), fg="black",bg="white",width=15,relief="groove",border=5)
            txtname.place(x=140,y=200)
            lblquantity = tk.Button(MonthlyReport, text="Quantity", font=("Calibri", 15, "bold"), fg="white",bg="#09ad2c",width=10)
            lblquantity.place(x=350,y=200)
            txtquantity = tk.Entry(MonthlyReport,font=("Calibri", 18, "bold"), fg="black",bg="white",width=15,relief="groove",border=5)
            txtquantity.place(x=500,y=200)

            ExpiredForm = tk.Frame(SecondLine,width=700,height=250,bg="white",highlightbackground="#c3c7c7", highlightthickness=8)
            ExpiredForm.pack(side="left",fill="x",expand=True)
            lbltitle = tk.Label(ExpiredForm,text="Expired Form", font=("Calibri", 20, "bold"), fg="Black", bg="#c3c7c7")
            lbltitle.place(x=70,y=20)

            #------------------------------------
            FinalLine = tk.Frame(self,bg="white",width=30,height=370,highlightbackground="#c3c7c7", highlightthickness=8)
            FinalLine.pack(fill=tk.X)
            lbltitle = tk.Label(FinalLine,text="Expired Form", font=("Calibri", 20, "bold"), fg="Black", bg="#c3c7c7")
            lbltitle.place(x=70,y=20)
            



# ------------------------------------ Main Rotate File ----------------------
# 
# 
# 
#


class Application(tk.Tk):
     def __init__(self, *args, **kwargs):
          tk.Tk.__init__(self, *args, **kwargs)

          #creating a window
          window = tk.Frame(self)
          window.pack()

          window.grid_rowconfigure(0, minsize =800) # 500
          window.grid_columnconfigure(0, minsize = 1280) # 800

          self.frames = {}
          for F in (FirstPage, SecondPage, ThirdPage, FourthPage, FivePage, SixPage,SevenPage,EightPage,NinePage):
               frame = F(window, self)
               self.frames[F] = frame
               frame.grid(row = 0, column=0, sticky="nsew")
            
          self.show_frame(FirstPage)
        
     def show_frame(self, page):
          frame = self.frames[page]
          frame.tkraise()
          w, h = self.winfo_screenwidth(), self.winfo_screenheight()
          #self.config("%dx%d+0+0" % (w, h))
          self.title("POS Application")
          self.geometry('1980x1050+0+0')
            
app = Application()
#app.maxsize(1500,800)
app.mainloop()