import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk  # pip install pillow
from tkinter.filedialog import asksaveasfilename


from tkinter import ttk



from dblogin import Registerdata

# --------------------------------- Login form ------------------------
class FirstPage(tk.Frame):
     def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        load = Image.open("login1.png")
        photo = ImageTk.PhotoImage(load)
        label = tk.Label(self, image=photo)
        label.image=photo
        label.place(x=0,y=0)
        lblTitle=tk.Label(self,text="Eg Login System", font=("arial", 45, 'bold'),fg="black",bg="#fcc203")
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

        dbRegister = Registerdata("real/Registerdatabase.db")
        register = []
        register = dbRegister.fetch()
        registerUser = dbRegister.fetchUser()
        def Insert():
          dbRegister(T1.get(),)
 
        def verify():
            if T1.get() == "admin" and T2.get() == "12345@":
               text1.set('')
               text2.set('')
               controller.show_frame(ThirdPage)
            else: 
               try:
                    #with open("credential.txt", "r") as f:
                    #     info = f.readlines()
                    #     i  = 0
                    #     for e in info:
                    #         u, p =e.split(",")
                    #         if u.strip() == T1.get() and p.strip() == T2.get():
                    #             # -------------------  Login Design --------------------
                    #             controller.show_frame(SixPage)
                    #             i = 1
                    #             break
                    #     if i==0:
                    #          messagebox.showinfo("Error", "Please provide correct username and password!!")
                    i = 0
                    for a in range(len(register)):
                        if register[a][1]==T1.get() and register[a][2]==T2.get():
                            controller.show_frame(ThirdPage)
                            i = 1
                            break
                    if i==0:
                        messagebox.showinfo("Error", "Please provide correct username and password!!")
               except:
                      
                    messagebox.showinfo("Error", "Please provide correct username and password!!")
            
        def Userverify():
            if T1.get() == "admin" and T2.get() == "12345@":
               text1.set('')
               text2.set('')
               controller.show_frame(ThirdPage)
            else: 
               try:
                    #with open("credential.txt", "r") as f:
                    #     info = f.readlines()
                    #     i  = 0
                    #     for e in info:
                    #         u, p =e.split(",")
                    #         if u.strip() == T1.get() and p.strip() == T2.get():
                    #             # -------------------  Login Design --------------------
                    #             controller.show_frame(SixPage)
                    #             i = 1
                    #             break
                    #     if i==0:
                    #          messagebox.showinfo("Error", "Please provide correct username and password!!")
                    i = 0
                    for a in range(len(register)):
                        if registerUser[a][1]==T1.get() and registerUser[a][2]==T2.get():
                            controller.show_frame(ThirdPage)
                            i = 1
                            break
                    if i==0:
                        messagebox.showinfo("Error", "Please provide correct username and password!!")
               except:
                      
                    messagebox.showinfo("Error", "Please provide correct username and password!!")
        
        tk.Button(mainframe,text="Admin",height="2",font=("Arial Bold", 15),width="14",bg="#ed3833",fg="white",bd=0,command=verify).place(x=70,y=250)
        tk.Button(mainframe,text="User",height="2",font=("Arial Bold", 15),width="14",bg="#f1f300",fg="white",bd=0,command=Userverify).place(x=290,y=250)

        tk.Button(mainframe,text="Exit",height="2",font=("Arial Bold", 15),width="14",bg="#00bd56",fg="white",bd=0,command=self.destroy).place(x=520,y=250)
          #B1 = tk.Button(border, text="Submit", font=("Arial", 15), command=verify)
          #B1.place(x=320, y=115)
  
# ------------------------------------  Main Form (second page)  ------------------------------------

#----------------------------------------- example form for tkinter ----------    
class ThirdPage(tk.Frame):
     def __init__(self, parent, controller):
          tk.Frame.__init__(self, parent)

          self.configure(bg='Tomato')

          Label = tk.Label(self, text="Store some content related to your \n project or what your application made for. \n All the best!!", bg = "orange", font=("Arial Bold", 25))
          Label.place(x=40, y=150)

          Button = tk.Button(self, text="Home", font=("Arial", 15), command=lambda: controller.show_frame(FirstPage))
          Button.place(x=650, y=450)

 

class Application(tk.Tk):
     def __init__(self, *args, **kwargs):
          tk.Tk.__init__(self, *args, **kwargs)

          #creating a window
          window = tk.Frame(self)
          window.pack()

          window.grid_rowconfigure(0, minsize =800) # 500
          window.grid_columnconfigure(0, minsize = 1280) # 800

          self.frames = {}
          for F in (FirstPage,ThirdPage):
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
          self.iconbitmap(bitmap="logoicon.ico")
            
app = Application()
#app.maxsize(1500,800)
app.mainloop()