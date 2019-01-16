#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Dec 22 15:50:48 2018

@author: danieltjondro
"""

from tkinter import ttk
from tkinter import *
import numpy as np
import tkinter as tk
import xlrd 
import random 
import PIL.Image, PIL.ImageTk 

dict =	{
  1: "The finest steel has to go through the hottest fire. - Richard Nixon",
  2: "Persistence is to the character of man as carbon is to steel. - Napoleon Hill",
  3: "You've got to have steel in you somewhere. - Alan Bates",
  4: "One of the reasons that I work in steel is durability. - Jeff Koons",
  5: "Steel is needed everywhere. - Rodrigo Duterte",
  6: "Don't get lost in the code! - Powell Draper",
}

login_dict = {
        "dtjondro": ("whitman","Whitman","WHITMAN","Daniel"),
        "declerck": ("forbes","Forbes","FORBES","Keiko"),
        "ngrossi": ("mathey","Mathey","MATHEY","Natalie"),
        "mhallee": ("forbes","Forbes","FORBES","Mitchell"),
        "solmazj": ("forbes","Forbes","FORBES","Solmaz"),
        "ekeim" : ("rocky","Rocky","ROCKY","rockefeller","Rockefeller","ROCKEFELLER","Elizabeth"),
        "kimikom": ("whitman","Whitman","WHITMAN","Kimiko"),
        "rachelsc": ("butler","Butler","BUTLER","Rachel"),
        "cheubner": ("forbes","Forbes","FORBES","Camille"),
        "ofoster": ("mathey","Mathey","MATHEY","Olivia"),
        "nyema.wesley": ("rocky","Rocky","ROCKY","rockefeller","Rockefeller","ROCKEFELLER","Nyema"),
        "rkn2": ("equad","EQuad","EQUAD","Equad","Becca"),
        "pdraper": ("equad","EQuad","EQUAD","Equad","Professor Draper"),
        "yzjin": ("mathey","Mathey","MATHEY","Yolanda"),
        "mcoar": ("equad","EQuad","EQUAD","Equad","Max"),
        "cdr3": ("butler","Butler","BUTLER","Chris"),
        "ameenm": ("rocky","Rocky","ROCKY","rockefeller","Rockefeller","ROCKEFELLER","Ameen")
    }  

# Give the loc1ation of the file 
loc1 = ("aisc-shapes-database-v15.0h.xlsx")
loc2 = ("Compression.xlsx")
loc3 = ("Zx_values.xlsx")

# To open Workbook 
wb1 = xlrd.open_workbook(loc1) 
sheet1 = wb1.sheet_by_index(0)

wb2 = xlrd.open_workbook(loc2) 
w12sheet = wb2.sheet_by_index(0)
w14sheet = wb2.sheet_by_index(1)

wb3 = xlrd.open_workbook(loc3)
Zxsheet = wb3.sheet_by_index(0)

root = tk.Tk()
root.geometry('500x500')
root.configure(bg = "#C1CDCD")
root.title("Login")

method = tk.StringVar()
component = tk.StringVar()
steel = tk.StringVar()
shape_of_member = tk.StringVar()
type_of_load_d = tk.StringVar()
type_of_load_l = tk.StringVar()
type_of_connection = tk.StringVar()
type_of_bolt = tk.StringVar()

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

error_x = (screen_width/2) - (600/2)
error_y = (screen_height/2) - (90/2)

x_cor = (screen_width/2) - (500/2)
y_cor = (screen_height/2) - (500/2)
root.geometry("%dx%d+%d+%d" % (500,500,x_cor,y_cor))

photo = PhotoImage(file = "Steelpreme.gif")
background_label = tk.Label(root, image=photo,bd = 0,fg = "#C1CDCD")
background_label.configure(bg = "#C1CDCD")
background_label.pack(pady=(70,15))    

Username_l = tk.Label(root, text = "Netid:",bg = "#C1CDCD")
Username_l.pack()

Username_e = Entry(root, bd = 2, bg = "#C1CDCD")
Username_e.pack(pady = (0,15))

Password_l = tk.Label(root,bd = 0,text = "Residential College:",bg = "#C1CDCD")
Password_l.configure(bg = "#C1CDCD")  
Password_l.pack()

Password_e = Entry(root, bd = 2, bg = "#C1CDCD")
Password_e.pack(pady = (0,15))
Password_e.config(show="*")

def login():
    
    username = str(Username_e.get())
    password = str(Password_e.get())
    
    print (username)
    print (username in login_dict)
    
    if username in login_dict:
        
        if password in login_dict.get(username) and username in login_dict:
        
            homepage = Toplevel(root)
            homepage.title("Homepage")
            homepage.geometry("%dx%d+%d+%d" % (500,500,x_cor,y_cor))
            homepage.configure(bg = "#C1CDCD")
            
            Title1 = tk.Label(homepage,image=photo, bd = 0,fg = "#C1CDCD",bg = "#C1CDCD")
            welcome = tk.Label(homepage,text = "Welcome " + login_dict.get(username)[-1] + "!", bd = 0,bg = "#C1CDCD", font = "Verdana 20")
                                         
            T1 = tk.Label(homepage,text="Which design philosophy do you follow?",bg = "#C1CDCD")
            R_LRFD = tk.Radiobutton(homepage, text="LRFD", variable=method, value="LRFD",bg = "#C1CDCD")
            R_ASD = tk.Radiobutton(homepage, text="ASD", variable=method, value="ASD",bg = "#C1CDCD")
            
            T2 = tk.Label(homepage,text="Which structural component are you interested in?",bg = "#C1CDCD")
            R_Tension_Members = tk.Radiobutton(homepage, text="Tension Members", variable=component, value="Tension Members",bg = "#C1CDCD")
            R_Compression_Members = tk.Radiobutton(homepage, text="Compression Members", variable=component, value="Compression Members",bg = "#C1CDCD")
            R_Bending_Members = tk.Radiobutton(homepage, text="Bending Members", variable=component, value="Bending Members",bg = "#C1CDCD")
            R_Connection = tk.Radiobutton(homepage, text="Connections", variable=component, value="Connections",bg = "#C1CDCD")
            
            Title1.pack(pady = (0,10))
                                    
            welcome.pack(anchor='w', pady = (0,10))
            T1.pack(anchor='w')
            R_LRFD.pack(anchor='w')
            R_ASD.pack(pady=(0,10), anchor='w')
            
            T2.pack(anchor='w')
            R_Tension_Members.pack(anchor='w')
            R_Compression_Members.pack(anchor='w')
            R_Bending_Members.pack(anchor='w')
            R_Connection.pack(pady=(0,15), anchor='w')
            
            button = ttk.Button(homepage, text="Next!",command=new_winF,style='TButton')
            ttk.Style().configure("TButton", padding=6, relief="flat",background="#000")
            button.pack(anchor='w')
        
        else:
            
            invalid_login = Toplevel(root)
            invalid_login.title("Invalid login")
            invalid_login.geometry("%dx%d+%d+%d" % (500,500,x_cor,y_cor))
            invalid_login.configure(bg = "#C1CDCD")
                               
            error1 = tk.Label(invalid_login, text="Invalid login", font = "Verdana 30",bg = "#C1CDCD")
            error1.pack(fill=X)
            error2 = tk.Label(invalid_login, text="Please try again!", font = "Verdana 18",bg = "#C1CDCD")
            error2.pack(fill=X)
            invalid_login.geometry('600x90')
            invalid_login.configure(bg = "#C1CDCD")
            invalid_login.geometry("%dx%d+%d+%d" % (600,90,error_x,error_y))
    else:
        
            invalid_login = Toplevel(root)
            invalid_login.title("Invalid login")
            invalid_login.geometry("%dx%d+%d+%d" % (500,500,x_cor,y_cor))
            invalid_login.configure(bg = "#C1CDCD")
                               
            error1 = tk.Label(invalid_login, text="Invalid login", font = "Verdana 30",bg = "#C1CDCD")
            error1.pack(fill=X)
            error2 = tk.Label(invalid_login, text="Please try again!", font = "Verdana 18",bg = "#C1CDCD")
            error2.pack(fill=X)
            invalid_login.geometry('600x90')
            invalid_login.configure(bg = "#C1CDCD")
            invalid_login.geometry("%dx%d+%d+%d" % (600,90,error_x,error_y))
            

Login = ttk.Button(root, text = "Login", command = login, style = 'TButton')
ttk.Style().configure("TButton", padding=6, relief="flat",background="#000")
Login.pack()
         
#Title1 = tk.Label(root,text="Welcome to Steel Shape Up",bg = "#C1CDCD", font = "Verdana 24")
#Title2 = tk.Label(root,text="an app that helps you design steel members!",bg = "#C1CDCD", font = "Verdana 20")
#                 
#T1 = tk.Label(root,text="Which design philosophy do you follow?",bg = "#C1CDCD")
#R_LRFD = tk.Radiobutton(root, text="LRFD", variable=method, value="LRFD",bg = "#C1CDCD")
#R_ASD = tk.Radiobutton(root, text="ASD", variable=method, value="ASD",bg = "#C1CDCD")
#
#T2 = tk.Label(root,text="Which structural component are you interested in?",bg = "#C1CDCD")
#R_Tension_Members = tk.Radiobutton(root, text="Tension Members", variable=component, value="Tension Members",bg = "#C1CDCD")
#R_Compression_Members = tk.Radiobutton(root, text="Compression Members", variable=component, value="Compression Members",bg = "#C1CDCD")
#R_Bending_Members = tk.Radiobutton(root, text="Bending Members", variable=component, value="Bending Members",bg = "#C1CDCD")
#R_Connection = tk.Radiobutton(root, text="Connections", variable=component, value="Connections",bg = "#C1CDCD")
#
#Title1.grid(row = 0)
#Title2.grid(row = 1, pady = (0,10))
#
#root.columnconfigure(0, weight=1)
#root.rowconfigure(0, weight=0)
#                              
#T1.grid(row=2, column=0, sticky="W")
#R_LRFD.grid(row=3, column=0, sticky="W")
#R_ASD.grid(row=4, column=0, sticky="W", pady=(0,10))
#
#T2.grid(row=5, column=0, sticky="W")
#R_Tension_Members.grid(row=6, column=0, sticky="W")
#R_Compression_Members.grid(row=7t, column=0, sticky="W")
#R_Bending_Members.grid(row=8, column=0, sticky="W")
#R_Connection.grid(row=9, column=0, sticky="W", pady=(0,10))
        
def new_winF():
    
    if component.get() == '' or method.get() == '':
        newwin = Toplevel(root)
        error1 = Label(newwin, text="Error! Please make sure you have filled ", font = "Verdana 20",bg = "#C1CDCD")
        error2 = Label(newwin, text="out all the necessary fields correctly!", font = "Verdana 20",bg = "#C1CDCD")
        error3 = Label(newwin, text="Hint: Make sure you specified the distribution of the loads!", font = "Verdana 12",bg = "#C1CDCD")
        
        error1.pack(fill=X)
        error2.pack(fill=X)
        error3.pack(fill=X)
        newwin.geometry('600x90')
        newwin.configure(bg = "#C1CDCD")
        newin.geometry("%dx%d+%d+%d" % (600,90,error_x,error_y))
    
    if component.get() == "Tension Members":
    
        newwin = Toplevel(root)
        newwin.title("Tension Member Design")
        newwin.geometry('500x500')
        newwin.configure(bg = "#C1CDCD")
        display = Label(newwin, text="Tension member design by " + method.get(), font = "Verdana 20",bg = "#C1CDCD")
        display.grid(row = 0, column = 0, sticky = "W")
        newwin.geometry("%dx%d+%d+%d" % (500,500,x_cor,y_cor))
                            
        Load_Instructions1 = tk.Label(newwin, text="Enter either both a dead and live load, or the required strength:",bg = "#C1CDCD")
        Load_Instructions1.grid(row = 1, column = 0, sticky = "W")
        Parameters_1 = tk.Label(newwin, text="Dead Load (in kips):",bg = "#C1CDCD")
        Parameters_1.grid(row=3, column=0, sticky="W")
        Entry_1 = tk.Entry(newwin, bd = 1,bg = "#C1CDCD")
        Entry_1.grid(row=4, column = 0, sticky="W")
        
        Parameters_2 = tk.Label(newwin, text="Live Load (in kips):",bg = "#C1CDCD")
        Parameters_2.grid(row=5, column=0, sticky="W")
        Entry_2 = tk.Entry(newwin, bd = 1,bg = "#C1CDCD")
        Entry_2.grid(row=6, column = 0, sticky="W", pady = (0,12))
        
        Parameters_4 = tk.Label(newwin, text="Required Strength (in kips):",bg = "#C1CDCD")
        Parameters_4.grid(row=7, column=0, sticky="W")
        Entry_4 = tk.Entry(newwin, bd = 1,bg = "#C1CDCD")
        Entry_4.grid(row=8,column=0,sticky="W", pady = (0,12))
        
        T4 = tk.Label(newwin,text="Which shape are you planning to use?",bg = "#C1CDCD")
        T4.grid(row=9, column=0, sticky="W")
        Shape1 = tk.Radiobutton(newwin, text="Double-Angle", variable=shape_of_member, value="Double-Angle",bg = "#C1CDCD")
        Shape1.grid(row=10, column=0, sticky="W")
        Shape2 = tk.Radiobutton(newwin, text="WT9", variable=shape_of_member, value="WT9",bg = "#C1CDCD")
        Shape2.grid(row=11, column=0, sticky="W")
        Shape3 = tk.Radiobutton(newwin, text="WT12", variable=shape_of_member, value="WT12",bg = "#C1CDCD")
        Shape3.grid(row=12, column=0, sticky="W")
        Shape4 = tk.Radiobutton(newwin, text="W14", variable=shape_of_member, value="W14",bg = "#C1CDCD")
        Shape4.grid(row=13, column=0, sticky="W",pady = (0,10))
        
        T3 = tk.Label(newwin,text="Which steel are you planning to use?",bg = "#C1CDCD")
        T3.grid(row=14, column=0, sticky="W")
        Type_of_steel1 = tk.Radiobutton(newwin, text="A36 Steel", variable=steel, value="A36 Steel",bg = "#C1CDCD")
        Type_of_steel1.grid(row=15, column=0, sticky="W")
        Type_of_steel2 = tk.Radiobutton(newwin, text="A992 Steel", variable=steel, value="A992 Steel",bg = "#C1CDCD")
        Type_of_steel2.grid(row=16, column=0, sticky="W", pady=(0,10))
    
    if component.get() == "Compression Members":
    
        newwin = Toplevel(root)
        newwin.title("Compression Member Design")
        newwin.geometry('500x500')
        newwin.configure(bg = "#C1CDCD")
        display = Label(newwin, text="Compression member design by " + method.get(), font = "Verdana 20",bg = "#C1CDCD")
        display.grid(row = 0, column = 0, sticky = "W",columnspan=2)
        newwin.geometry("%dx%d+%d+%d" % (500,500,x_cor,y_cor))        
                    
        Load_Instructions = tk.Label(newwin, text="Enter values in all three load types, or just the required strength:",bg = "#C1CDCD")
        Load_Instructions.grid(row = 1, column = 0, sticky = "W",columnspan=2)
        Parameters_1 = tk.Label(newwin, text="Dead Load (in kips):",bg = "#C1CDCD")
        Parameters_1.grid(row=2, column=0, sticky="W")
        Entry_1 = tk.Entry(newwin, bd = 1,bg = "#C1CDCD")
        Entry_1.grid(row=3, column = 0, sticky="W")
        
        Parameters_2 = tk.Label(newwin, text="Live Load (in kips):",bg = "#C1CDCD")
        Parameters_2.grid(row=4, column=0, sticky="W")
        Entry_2 = tk.Entry(newwin, bd = 1,bg = "#C1CDCD")
        Entry_2.grid(row=5, column = 0, sticky="W")
        
        Parameters_3 = tk.Label(newwin, text="Wind Load (in kips):",bg = "#C1CDCD")
        Parameters_3.grid(row=6, column=0, sticky="W",columnspan=2)
        Entry_3 = tk.Entry(newwin, bd = 1,bg = "#C1CDCD")
        Entry_3.grid(row=7, column = 0, sticky="W",pady=(0,15))
        
        Parameters_4 = tk.Label(newwin, text="Required Strength (in kips):",bg = "#C1CDCD")
        Parameters_4.grid(row=2, column=1, sticky="W",columnspan=2)
        Entry_4 = tk.Entry(newwin, bd = 1,bg = "#C1CDCD")
        Entry_4.grid(row=3,column=1,sticky="W")
        
        Length_Instructions = tk.Label(newwin, text="Enter the effective lengths about the x- and y-axes:",bg = "#C1CDCD")
        Length_Instructions.grid(row = 8, column = 0, sticky = "W",columnspan=2)
        Parameters_5 = tk.Label(newwin, text="Effective length about y-axis:",bg = "#C1CDCD")
        Parameters_5.grid(row=9, column=0, sticky="W")
        Entry_5 = tk.Entry(newwin, bd = 1,bg = "#C1CDCD")
        Entry_5.grid(row=10,column=0,sticky="W")
        
        Parameters_6 = tk.Label(newwin, text="Effective length about x-axis:",bg = "#C1CDCD")
        Parameters_6.grid(row=11, column=0, sticky="W",columnspan=2)
        Entry_6 = tk.Entry(newwin, bd = 1,bg = "#C1CDCD")
        Entry_6.grid(row=12,column=0,sticky="W",pady=(0,15))
        
        Shape_Instructions = tk.Label(newwin, text="Which shape would you like to use?",bg = "#C1CDCD")
        Shape_Instructions.grid(row = 13, column = 0, sticky = "W")
        Shape1 = tk.Radiobutton(newwin, text="W12", variable=shape_of_member, value="W12",bg = "#C1CDCD")
        Shape1.grid(row=14, column=0, sticky="W")
        Shape2 = tk.Radiobutton(newwin, text="W14", variable=shape_of_member, value="W14",bg = "#C1CDCD")
        Shape2.grid(row=15, column=0, sticky="W",pady = (0,15))
        
    if component.get() == "Bending Members":
        
        newwin = Toplevel(root)
        newwin.geometry('500x500')
        newwin.title("Bending Member Design")
        newwin.configure(bg = "#C1CDCD")
        display = Label(newwin, text="Bending member design by " + method.get(), font = "Verdana 20",bg = "#C1CDCD")
        display.grid(row = 0, column = 0, sticky = "W")
        newwin.geometry("%dx%d+%d+%d" % (500,500,x_cor,y_cor))
        
        Parameters_1 = tk.Label(newwin, text="Span length (in feet):",bg = "#C1CDCD")
        Parameters_1.grid(row = 1, column = 0, sticky = "W")
        Entry_1 = tk.Entry(newwin, bd = 1,bg = "#C1CDCD")
        Entry_1.grid(row=2, column = 0, sticky="W", pady = (0,10))
        
        Parameters_2 = tk.Label(newwin, text="Dead load (in kips):",bg = "#C1CDCD")
        Parameters_2.grid(row = 3, column = 0, sticky = "W")
        Entry_2 = tk.Entry(newwin, bd = 1,bg = "#C1CDCD")
        Entry_2.grid(row=4, column = 0, sticky="W")
        point_load1 = tk.Radiobutton(newwin, text="Point load at midspan", variable=type_of_load_d, value="point_load",bg = "#C1CDCD")
        point_load1.grid(row=5, column=0, sticky="W")
        dist_load1 = tk.Radiobutton(newwin, text="Distributed load along entire span", variable=type_of_load_d, value="dist_load",bg = "#C1CDCD")
        dist_load1.grid(row=6, column=0, sticky="W", pady = (0,10))
        
        Parameters_3 = tk.Label(newwin, text="Live load (in kips):",bg = "#C1CDCD")
        Parameters_3.grid(row = 8, column = 0, sticky = "W")
        Entry_3 = tk.Entry(newwin, bd = 1,bg = "#C1CDCD")
        Entry_3.grid(row=9, column = 0, sticky="W")
        point_load2 = tk.Radiobutton(newwin, text="Point load at midspan", variable=type_of_load_l, value="point_load",bg = "#C1CDCD")
        point_load2.grid(row=10, column=0, sticky="W")
        dist_load2 = tk.Radiobutton(newwin, text="Distributed load along entire span", variable=type_of_load_l, value="dist_load",bg = "#C1CDCD")
        dist_load2.grid(row=11, column=0, sticky="W", pady = (0,10))
        
    if component.get() == "Connections":
    
        newwin = Toplevel(root)
        newwin.geometry('500x500')
        newwin.title("Connection Member Design")
        newwin.configure(bg = "#C1CDCD")
        display = Label(newwin, text="Connection member design by " + method.get(), font = "Verdana 20",bg = "#C1CDCD")
        display.grid(row = 0, column = 0, sticky = "W")
        newwin.geometry("%dx%d+%d+%d" % (500,500,x_cor,y_cor))           
         
        Load_Instructions = tk.Label(newwin, text="Would you like to determine the design strength for a bolt or c-shaped weld?",bg = "#C1CDCD")
        Load_Instructions.grid(row = 1, column = 0, sticky = "W")
        Choice_1 = tk.Radiobutton(newwin, text="Bolt",variable=type_of_connection,value="Bolt",bg = "#C1CDCD")
        Choice_1.grid(row=2, column=0, sticky="W")
        Choice_2 = tk.Radiobutton(newwin, text="C-Shaped weld",variable=type_of_connection,value = "C-Shaped weld", bg = "#C1CDCD")
        Choice_2.grid(row=3, column = 0, sticky="W",pady=(0,12))
        
        Bolt_Instructions1 = tk.Label(newwin, text="If you chose bolt, which type of bolt?",bg = "#C1CDCD")
        Bolt_Instructions1.grid(row = 4, column = 0, sticky = "W")
        Bolt_1 = tk.Radiobutton(newwin, text="A325-N",variable=type_of_bolt,value="A325-N",bg = "#C1CDCD")
        Bolt_1.grid(row = 5, column = 0, sticky = "W")                        
        Bolt_2 = tk.Radiobutton(newwin, text="A490-X",variable=type_of_bolt,value="A490-X",bg = "#C1CDCD")
        Bolt_2.grid(row = 6, column = 0, sticky = "W") 
        Bolt_Instructions2 = tk.Label(newwin, text="What is the diameter (in inches) of the bolt? Please enter as a decimal",bg = "#C1CDCD")
        Bolt_Instructions2.grid(row = 7, column = 0, sticky = "W")
        Bolt_Diameter = tk.Entry(newwin, bd = 1, bg = "#C1CDCD")             
        Bolt_Diameter.grid(row = 8, column = 0, sticky = "W", pady=(0,12))
                                 
        Weld_Instructions = tk.Label(newwin, text="If you chose c-shaped weld, what are the dimensions?",bg = "#C1CDCD")
        Weld_Instructions.grid(row = 9, column = 0, sticky = "W")
        
        Parallel_length_Instructions = tk.Label(newwin, text="Parallel length (in inches):",bg = "#C1CDCD")
        Parallel_length_Instructions.grid(row = 10, column = 0, sticky = "W")
        Parallel_length = tk.Entry(newwin, bd = 1, bg = "#C1CDCD")
        Parallel_length.grid(row = 11, column = 0, sticky = "W")  
                      
        Transverse_length_Instructions = tk.Label(newwin, text="Transverse length (in inches):",bg = "#C1CDCD")
        Transverse_length_Instructions.grid(row = 12, column = 0, sticky = "W")
        Transverse_length = tk.Entry(newwin, bd = 1, bg = "#C1CDCD")
        Transverse_length.grid(row = 13, column = 0, sticky = "W") 
        
        length_of_weld_Instructions = tk.Label(newwin, text="Length of weld (in inches), please enter as a decimal:",bg = "#C1CDCD")
        length_of_weld_Instructions.grid(row = 14, column = 0, sticky = "W")
        length_of_weld = tk.Entry(newwin, bd = 1, bg = "#C1CDCD")
        length_of_weld.grid(row = 15, column = 0, sticky = "W", pady=(0,12))
                                               
    def Bending_Design():
        
        bendwin = Toplevel(newwin)
        if (Entry_1.get() == '' or
        Entry_2.get() == '' or
        Entry_3.get() == '' or
        type_of_load_d.get() == '' or
        type_of_load_l.get() == '' 
            ):
            bendwin.title("Error")
            error1 = Label(bendwin, text="Error! Please make sure you have filled ", font = "Verdana 20",bg = "#C1CDCD")
            error2 = Label(bendwin, text="out all the necessary fields correctly!", font = "Verdana 20",bg = "#C1CDCD")
            error3 = Label(bendwin, text="Hint: Make sure you specified the distribution of the loads", font = "Verdana 12",bg = "#C1CDCD")
            
            error1.pack(fill=X)
            error2.pack(fill=X)
            error3.pack(fill=X)
            bendwin.geometry('600x90')
            bendwin.configure(bg = "#C1CDCD")
            bendwin.geometry("%dx%d+%d+%d" % (600,90,error_x,error_y))
        
        else:
        
            bendwin.title("Bending Member Design Results")
            bendwin_display = Label(bendwin, text="Bending member design by " + method.get(), font = "Verdana 24",bg = "#C1CDCD")
            bendwin_display.grid(row = 0, column = 0, sticky = "W",pady = (0,12))
            bendwin.geometry('500x500')
            bendwin.configure(bg = "#C1CDCD")
            bendwin.geometry("%dx%d+%d+%d" % (500,500,x_cor,y_cor))
            
            bending_load = tk.Label(bendwin,bg = "#C1CDCD")
            bending_modulus = tk.Label(bendwin,bg = "#C1CDCD")
            bending_shape1 = tk.Label(bendwin,bg = "#C1CDCD")
            bending_shape2 = tk.Label(bendwin,bg = "#C1CDCD") 
            bending_self1 = tk.Label(bendwin,bg = "#C1CDCD")
            bending_self2 = tk.Label(bendwin,bg = "#C1CDCD")
            bending_modulus_new = tk.Label(bendwin,bg = "#C1CDCD")
            bending_shape_final1 = tk.Label(bendwin,bg = "#C1CDCD")
            bending_shape_final2 = tk.Label(bendwin,bg = "#C1CDCD")
            quote = tk.Label(bendwin,bg = "#C1CDCD")
            
            if Entry_2.get() == '':
                Entry2 = 0
            else:
                Entry2 = float(Entry_2.get()) 
                
            if Entry_3.get() == '':
                Entry3 = 0
            else:
                Entry3 = float(Entry_3.get())
                
            span_length = float(Entry_1.get())
            
            if method.get() == "LRFD":
                if type_of_load_d.get() == "point_load":
                    dead_load = 1.2 * Entry2
                    dead_moment = 1.2 * Entry2 * span_length / 4
                else:
                    dead_load = 1.2 * Entry2 * span_length
                    dead_moment = 1.2 * Entry2 * span_length * span_length / 8
                if type_of_load_l.get() == "point_load":
                    live_load = 1.6 * Entry3
                    live_moment = 1.6 * Entry3 * span_length / 4
                else:
                    live_load = 1.6 * Entry3 * span_length
                    live_moment = 1.6 * Entry3 * span_length * span_length / 8
                    
            if method.get() == "ASD":
                if type_of_load_d.get() == "point_load":
                    dead_load = Entry2
                    dead_moment = Entry2 * span_length / 4
                else:
                    dead_load = Entry2 * span_length
                    dead_moment = Entry2 * span_length * span_length / 8
                if type_of_load_l.get() == "point_load":
                    live_load = Entry3
                    live_moment = Entry3 * span_length / 4
                else:
                    live_load = Entry3 * span_length
                    live_moment = Entry3 * span_length * span_length / 8
            
            moment = dead_moment + live_moment
            total_load = dead_load + live_load
            bending_load["text"] = "The required strengths are Pu = " + str("%.2f" % total_load) + " kips and Mu = " + str("%.2f" % moment) + " ft-kips."
            bending_load.grid(row=1, column=0, sticky="W", pady = (0,10))
            
            if method.get() == "LRFD":
                Z_req = moment * 12 / (0.90 * 50)
            if method.get() == "ASD":
                Z_req = moment * 12 * 1.67 / 50
            bending_modulus["text"] = "The required plastic section modulus = " + str("%.2f" % Z_req) + " inches cubed."
            bending_modulus.grid(row=2, column=0, sticky="W", pady = (0,10))
            
            z_value = 0
            for i in range (1,36):
                if Zxsheet.cell_value(36-i,1) > Z_req:
                    best_shape = Zxsheet.cell_value(36-i,0)
                    weight = Zxsheet.cell_value(36-i,0)[4:]
                    z_value = Zxsheet.cell_value(36-i,1)
                    break
                
            if z_value != 0:
                bending_shape1["text"] = "Using this required plastic section modulus, we recommend the "
                bending_shape2["text"] = "minimum-weight W-shape: " + str(best_shape) + " with Zx = " + str(z_value) + "."
                bending_shape1.grid(row=3, column=0, sticky="W")
                bending_shape2.grid(row=4, column=0, sticky="W",pady = (0,10))
            
            else:
                bending_shape1["text"] = "Sorry, I was not able to find a shape that satisfies that section modulus!"
                bending_shape1.grid(row=3, column=0, sticky="W")
            
            if z_value != 0:
                if method.get() == "LRFD":
                    moment_self_weight = 1.2 * float(weight)/1000 * .125 * span_length * span_length
                if method.get() == "ASD":
                    moment_self_weight = float(weight)/1000 * .125 * span_length * span_length
                
                bending_self1["text"] = "The additional required strength based on the actual weight of "
                bending_self2["text"] = "this beam = " + str("%.2f" % moment_self_weight) + " ft-kips."
                bending_self1.grid(row = 5, column=0, sticky="W")
                bending_self2.grid(row = 6, column=0, sticky="W",pady = (0,10))
                
                if method.get() == "LRFD":
                    new_Z_req = (moment + moment_self_weight) * 12 / (0.90 * 50)
                if method.get() == "ASD":
                    new_Z_req = (moment + moment_self_weight) * 12 * 1.67 / 50
        
                bending_modulus_new["text"] = "The new required plastic section modulus = " + str("%.2f" % new_Z_req) + " inches cubed."
                bending_modulus_new.grid(row = 7, column=0, sticky="W",pady = (0,10))
                
                if z_value > new_Z_req:
                    bending_shape_final1["text"] = "This required plastic section modulus is less than that provided by the "
                    bending_shape_final2["text"] = str(best_shape) + " we have already chosen."
                else:
                    for i in range (1,36):
                        if Zxsheet.cell_value(36-i,1) > new_Z_req:
                            best_shape = Zxsheet.cell_value(36-i,0)
                            weight = Zxsheet.cell_value(36-i,0)[4:]
                            z_value = Zxsheet.cell_value(36-i,1)
                            break
                    bending_shape_final["text"] = "Using the new required plastic section modulus, we have selected the minimum-weight W-shape: " + str(best_shape) + "."
                bending_shape_final1.grid(row = 8, column=0, sticky="W")
                bending_shape_final2.grid(row = 9, column=0, sticky="W")
            
            quote["text"] = dict[random.randint(1,6)]
            quote.place(rely=1, relx=0, x=0, y=0, anchor=SW)
    
    def Compression_Design():
        
        compwin = Toplevel(newwin)
        
        if (Entry_1.get() == '' and Entry_2.get() == '' and Entry_3.get() == '' and Entry_4.get() == '' or
        Entry_5.get() == '' and Entry_6.get() == '' or
        shape_of_member.get() == '' or
        Entry_4.get() != '' and (Entry_1.get() != '' or Entry_2.get() != '' or Entry_3.get() != '')):
            compwin.title("Error")
            error1 = Label(compwin, text="Error! Please make sure you have filled ", font = "Verdana 20",bg = "#C1CDCD")
            error2 = Label(compwin, text="out all the necessary fields correctly!", font = "Verdana 20",bg = "#C1CDCD")
            error3 = Label(compwin, text="Hint: Make sure you only either fill in the D/L/W loads OR the required strength!", font = "Verdana 12",bg = "#C1CDCD")

            error1.pack(fill=X)
            error2.pack(fill=X)
            error3.pack(fill=X)
            compwin.geometry('600x90')
            compwin.configure(bg = "#C1CDCD")
            compwin.geometry("%dx%d+%d+%d" % (600,90,error_x,error_y))
        else:
            
            compwin.title("Compression Member Design Results")
            compwin_display = Label(compwin, text="Compression member design by " + method.get(), font = "Verdana 24",bg = "#C1CDCD")
            compwin_display.grid(row = 0, column = 0, sticky = "W", pady = (0,12))
            compwin.geometry('500x500')
            compwin.configure(bg = "#C1CDCD")
            compwin.geometry("%dx%d+%d+%d" % (500,500,x_cor,y_cor))
            
            compression_label = tk.Label(compwin,bg = "#C1CDCD")
            compression_length1 = tk.Label(compwin,bg = "#C1CDCD")
            compression_length2 = tk.Label(compwin,bg = "#C1CDCD")
            compression_shape1 = tk.Label(compwin,bg = "#C1CDCD")
            compression_shape2 = tk.Label(compwin,bg = "#C1CDCD")
            compression_new_length = tk.Label(compwin,bg = "#C1CDCD")
            compression_new_controlling_length1 = tk.Label(compwin,bg = "#C1CDCD")
            compression_new_controlling_length2 = tk.Label(compwin,bg = "#C1CDCD")
            compression_final_shape1 = tk.Label(compwin,bg = "#C1CDCD")
            compression_final_shape2 = tk.Label(compwin,bg = "#C1CDCD")
            quote = tk.Label(compwin,bg = "#C1CDCD")
            
            if Entry_3.get() == '':
                Entry3 = 0
            else:
                Entry3 = float(Entry_3.get())  
            if Entry_2.get() == '':
                Entry2 = 0
            else:
                Entry2 = float(Entry_2.get())
            if Entry_1.get() == '':
                Entry1 = 0
            else:
                Entry1 = float(Entry_1.get())
            if Entry_4.get() != '':
                total_load = float(Entry_4.get())
            else:
                if method.get() == "LRFD":
                    total_load = max(1.4 * Entry1, 1.2 * Entry1 + 1.6 * Entry2, 1.2 * Entry1 + 0.5 * Entry2 + Entry3, 0.9 * Entry1 + Entry3)
                if method.get() == "ASD":
                    total_load = max(Entry1, Entry1 + Entry2, Entry1 + 0.6 * Entry3, Entry1 + 0.75 * Entry2 + 0.75 * 0.6 * Entry3, 0.6 * Entry1 + 0.6 * Entry3)
            
            if method.get() == "LRFD":
                increment = 2
            if method.get() == "ASD":
                increment = 1
                
            if shape_of_member.get() == "W12":
                sheet = w12sheet
                n = 11
            if shape_of_member.get() == "W14":
                sheet = w14sheet
                n = 15
            
            display = str("%.2f" % total_load)
            compression_label["text"] = "The column must carry a load of " + display + " kips."
            compression_label.grid(row=1, column=0, sticky="W",pady = (0,10))
            
            effective_length = float(Entry_6.get())/(2.44)
            controlling_effective_length = max(float(Entry_5.get()), effective_length)
            compression_length1["text"] = "Starting with an assumption of rx/ry = 2.44, the tentative controlling"
            compression_length2["text"] = "effective length = " + str("%.2f" % controlling_effective_length) + " feet."
            compression_length1.grid(row = 2, column = 0, sticky = "W")
            compression_length2.grid(row = 3, column = 0, sticky = "W", pady = (0,10))
            raw_length = controlling_effective_length
            
            controlling_effective_length = round(controlling_effective_length)
            if controlling_effective_length % 2 != 0 and controlling_effective_length > 20:
                controlling_effective_length = controlling_effective_length + 1
            best_value = float(sheet.cell_value(controlling_effective_length+2,2))
            shape = float(sheet.cell_value(0,2))
            r = 2.44
            for i in range(1,n):
                if sheet.cell_value(controlling_effective_length+2,increment+(2*i)) > total_load and sheet.cell_value(controlling_effective_length+2,increment+(2*i)) < best_value:
                    best_value = float(sheet.cell_value(controlling_effective_length+2,increment+(2*i)))
                    shape = float(sheet.cell_value(0,increment+(2*i)))
                    r = float(sheet.cell_value(34,increment+(2*i)))
            if shape_of_member.get() == "W12":        
                compression_shape1["text"] = "Try shape W12x" + str(shape) + ". It can carry up to " + str(best_value) + " kips "
                compression_shape2["text"] = "and has a rx/ry value = " + str(r)  + "."
            if shape_of_member.get() == "W14":
                compression_shape1["text"] = "Try shape W14x" + str(shape) + ". It can carry up to " + str(best_value) + " kips "
                compression_shape2["text"] = "and has a rx/ry value = " + str(r)  + "."
            compression_shape1.grid(row = 4, column = 0, sticky = "W")
            compression_shape2.grid(row = 5, column = 0, sticky = "W",pady = (0,10))
            
            new_controlling_effective_length = max(float(Entry_5.get()), float(float(Entry_6.get())/r))
            if new_controlling_effective_length == raw_length:
                compression_new_controlling_length1["text"] = "With that rx/ry value, our new controlling effective length "
                compression_new_controlling_length2["text"] = "remains at " + str("%.2f" % new_controlling_effective_length) + " feet."
            else:
                compression_new_controlling_length1["text"] = "With that rx/ry value, our new controlling effective length "
                compression_new_controlling_length2["text"] = "is " + str("%.2f" % new_controlling_effective_length) + " feet."
            compression_new_controlling_length1.grid(row = 6, column = 0, sticky = "W")
            compression_new_controlling_length2.grid(row = 7, column = 0, sticky = "W",pady = (0,10))
            
            new_controlling_effective_length = round(new_controlling_effective_length)
            if new_controlling_effective_length % 2 != 0 and new_controlling_effective_length > 20:
                new_controlling_effective_length = new_controlling_effective_length + 1
            best_value = float(sheet.cell_value(new_controlling_effective_length+2,2))
            shape = float(sheet.cell_value(0,2))
            r = 2.44
            for i in range(1,n):
                if sheet.cell_value(new_controlling_effective_length+2,increment+(2*i)) > total_load and sheet.cell_value(new_controlling_effective_length+2,increment+(2*i)) < best_value:
                    best_value = float(sheet.cell_value(new_controlling_effective_length+2,increment+(2*i)))
                    shape = float(sheet.cell_value(0,increment+(2*i)))
                    r = float(sheet.cell_value(34,increment+(2*i)))
            
            if shape_of_member.get() == "W12":
                compression_final_shape1["text"] = "Finally, we recommend you use a W12x" + str(shape) + "."
                compression_final_shape2["text"] = "It can carry up to " + str(best_value) + " kips at this effective length."
            if shape_of_member.get() == "W14":
                compression_final_shape1["text"] = "Finally, we recommend you use a W14x" + str(shape) + "."
                compression_final_shape2["text"] = "It can carry up to " + str(best_value) + " kips at this effective length."
            compression_final_shape1.grid(row = 8, column = 0, sticky = "W")
            compression_final_shape2.grid(row = 9, column = 0, sticky = "W")
            
            quote["text"] = dict[random.randint(1,6)]
            quote.place(rely=1, relx=0, x=0, y=0, anchor=SW)
    
    def Connections_Design():
        
        conwin = Toplevel(newwin)
        if (Parallel_length.get() == '' and Transverse_length.get() == '' and length_of_weld.get() == '' and type_of_connection.get() == "C-Shaped weld" or 
        Bolt_Diameter.get() == '' and type_of_connection.get() == "Bolt" or
        type_of_bolt.get() == '' and type_of_connection.get() == "Bolt" or 
        type_of_connection.get() == '' or
        Bolt_Diameter.get() != '' and (Parallel_length.get() != '' or Transverse_length.get() != '')):
            conwin.title("Error")
            error1 = Label(conwin, text="Error! Please make sure you have filled ", font = "Verdana 20",bg = "#C1CDCD")
            error2 = Label(conwin, text="out all the necessary fields correctly!", font = "Verdana 20",bg = "#C1CDCD")
            error3 = Label(conwin, text="Hint: Only fill in the information for the type of connection you chose!", font = "Verdana 12",bg = "#C1CDCD")

            error1.pack(fill=X)
            error2.pack(fill=X)
            error3.pack(fill=X)
            conwin.geometry('600x90')
            conwin.configure(bg = "#C1CDCD")
            conwin.geometry("%dx%d+%d+%d" % (600,90,error_x,error_y))
        else:
            conwin.title("Connection Member Design Results")
            conwin_display = Label(conwin, text="Connections design by " + method.get(), font = "Verdana 20",bg = "#C1CDCD")
            conwin_display.grid(row = 0, column = 0, sticky = "W")
            conwin.geometry('500x500')
            conwin.configure(bg = "#C1CDCD")
            conwin.geometry("%dx%d+%d+%d" % (500,500,x_cor,y_cor))
                             
            area = tk.Label(conwin,bg = "#C1CDCD")
            strength = tk.Label(conwin,bg = "#C1CDCD")
            p_strength = tk.Label(conwin,bg = "#C1CDCD")
            t_strength = tk.Label(conwin,bg = "#C1CDCD")
            quote = tk.Label(conwin,bg = "#C1CDCD")
            assumption = tk.Label(conwin,bg = "#C1CDCD")
            candidate1a = tk.Label(conwin,bg = "#C1CDCD")
            candidate1b = tk.Label(conwin,bg = "#C1CDCD")                       
            candidate2a = tk.Label(conwin,bg = "#C1CDCD")
            candidate2b = tk.Label(conwin,bg = "#C1CDCD")
            candidate2c = tk.Label(conwin,bg = "#C1CDCD")
            bolt_info1 = tk.Label(conwin,bg = "#C1CDCD")
            bolt_info2 = tk.Label(conwin,bg = "#C1CDCD")
            strength1 = tk.Label(conwin,bg = "#C1CDCD")
            strength2 = tk.Label(conwin,bg = "#C1CDCD")
        
            if type_of_connection.get() == "Bolt":
                                
                bolt_shank_area = float(Bolt_Diameter.get())**2 * np.pi * .25
                if type_of_bolt.get() == "A325-N":
                    nominal_shear_stress = 0.45 * 120
                    bolt_info1["text"] = "For an A325 bolt, Fu = 120 ksi and for the threads included," 
                    bolt_info2["text"] = "Fnv = (0.45)Fu = " + str("%.2f" % nominal_shear_stress) + " kips."

                if type_of_bolt.get() == "A490-X":
                    nominal_shear_stress = 0.563 * 150
                    bolt_info1["text"] = "For an A325 bolt, Fu = 120 ksi and for the threads included," 
                    bolt_info2["text"] = "Fnv = (0.563)Fu = " + str("%.2f" % nominal_shear_stress) + " kips."
                
                bolt_info1.grid(row=1, column=0, sticky="W")
                bolt_info2.grid(row=2, column=0, sticky="W",pady = (0,10))
                
                if method.get() == "LRFD":
                    strength1["text"] = "For shear, the design strength = (0.75)(" + str("%.2f" % nominal_shear_stress) + ")" + "(" +str("%.3f" % bolt_shank_area) + ") "
                    design_strength = 0.75 * nominal_shear_stress * bolt_shank_area
                if method.get() == "ASD":
                    strength1["text"] = "For shear, the design strength = (0.5)(" + str("%.2f" % nominal_shear_stress) + ")" + "(" +str("%.3f" % bolt_shank_area) + ") "
                    design_strength = nominal_shear_stress * bolt_shank_area * 0.5
                    
                area["text"] = "The bolt shank area = " + str("%.3f" % bolt_shank_area) + " inches squared."
                area.grid(row=3, column=0, sticky="W",pady = (0,10))
                strength2["text"] = "= " + str("%.3f" % design_strength) + " kips."
                strength1.grid(row=4, column=0, sticky="W")
                strength2.grid(row=5, column=0, sticky="W",pady = (0,10))
            
            elif type_of_connection.get() == "C-Shaped weld":
                
                assumption["text"] = "Assuming E70 electrodes:"
                assumption.grid(row=1, column=0, sticky="W")
                
                if method.get() == "LRFD":
                    p_str = float(Parallel_length.get()) * 2 * float(length_of_weld.get()) * 16 * 1.392
                    t_str = float(Transverse_length.get()) * float(length_of_weld.get()) * 16 * 1.392
                    
                if method.get() == "ASD":
                    p_str = float(Parallel_length.get()) * 2 * float(length_of_weld.get()) * 16 * .928
                    t_str = float(Transverse_length.get()) * float(length_of_weld.get()) * 16 * .928
                
                p_strength["text"] = "The parallel weld strength = " + str("%.2f" % p_str) + " kips."
                p_strength.grid(row=2, column=0, sticky="W",pady = (0,10))
                t_strength["text"] = "The transverse weld strength = " + str("%.2f" % t_str) + " kips."
                t_strength.grid(row=3, column=0, sticky="W",pady = (0,10))
                
                weld_str = max(p_str + t_str, 0.85 * p_str + 1.5 * t_str)
                candidate1a["text"] = "The connection design strength is obtained by adding the two strengths"
                candidate1b["text"] = "together: " + str("%.2f" % p_str) + " + " + str("%.2f" % t_str) + " = " + str("%.2f" % (p_str + t_str)) + " kips."
                candidate2a["text"] = "The design strength considering the added contribution of the transverse "
                candidate2b["text"] = "welds while reducing the contribution of the longitudinal welds:"
                candidate2c["text"] = "(0.85)(" + str("%.2f" % p_str) + ") + (1.5)(" + str("%.2f" % t_str) + ") = " + str("%.2f" % (0.85 * p_str + 1.5 * t_str)) + " kips."
                
                candidate1a.grid(row=4, column=0, sticky="W")
                candidate1b.grid(row=5, column=0, sticky="W", pady = (0,10))
                candidate2a.grid(row=6, column=0, sticky="W")
                candidate2b.grid(row=7, column=0, sticky="W")
                candidate2c.grid(row=8, column=0, sticky="W", pady = (0,10))
                
                strength["text"] = "The weld strength is the greater of the two values = " + str("%.2f" % weld_str) + " kips."
                strength.grid(row=9, column=0, sticky="W")
            
            quote["text"] = dict[random.randint(1,6)]
            quote.place(rely=1, relx=0, x=0, y=0, anchor=SW)
            
    def Tension_Design():   
        
        if steel.get() == "A36 Steel":
            Fy = 36
            Fu = 58
        elif steel.get() == "A992 Steel":
            Fy = 50
            Fu = 65
        
        tenwin = Toplevel(newwin)
        if (Entry_1.get() == '' and Entry_2.get() == '' and Entry_4.get() == '' or
        shape_of_member.get() == '' or steel.get() == '' or
        Entry_1.get() != '' and Entry_2.get() != '' and Entry_4.get() != ''):
            tenwin.title("Error")
            error1 = Label(tenwin, text="Error! Please make sure you have filled ", font = "Verdana 20",bg = "#C1CDCD")
            error2 = Label(tenwin, text="out all the necessary fields correctly!", font = "Verdana 20",bg = "#C1CDCD")
            error3 = Label(tenwin, text="Hint: Make sure you only either fill in the D/L loads OR the required strength!", font = "Verdana 12",bg = "#C1CDCD")

            error1.pack(fill=X)
            error2.pack(fill=X)
            error3.pack(fill=X)
            tenwin.geometry('600x90')
            tenwin.configure(bg = "#C1CDCD")
            tenwin.geometry("%dx%d+%d+%d" % (600,90,error_x,error_y))
        else:
            tenwin.title("Tension Member Design Results")
            tenwin_display = Label(tenwin, text="Tension member design by " + method.get(), font = "Verdana 24",bg = "#C1CDCD")
            tenwin_display.grid(row = 0, column = 0, sticky = "W", pady = (0,12))
            tenwin.geometry('500x500')
            tenwin.configure(bg = "#C1CDCD")
            tenwin.geometry("%dx%d+%d+%d" % (500,500,x_cor,y_cor))
        
            label2 = tk.Label(tenwin,bg = "#C1CDCD")
            tension_l1 = tk.Label(tenwin,bg = "#C1CDCD")
            tension_shape1 = tk.Label(tenwin,bg = "#C1CDCD")
            tension_shape2 = tk.Label(tenwin,bg = "#C1CDCD")
            tension_l2 = tk.Label(tenwin,bg = "#C1CDCD")
            tension_l3 = tk.Label(tenwin,bg = "#C1CDCD")
            tension_l4 = tk.Label(tenwin,bg = "#C1CDCD")
            quote = tk.Label(tenwin,bg = "#C1CDCD")
        
            if method.get() == "LRFD":
                if Entry_4.get() != '':
                    total_load = float(Entry_4.get())
                else: 
                    total_load = (float(1.2) * float(Entry_1.get())) + (float(1.6) * float(Entry_2.get()))
                display = str("%.2f" % total_load)
                label2["text"] = "Factored load = " + display + " kips"
                
                Min_area = total_load/(.90 * Fy)
                if shape_of_member.get() == "Double-Angle":
                    best_area = sheet1.cell_value(855, 10)
                    shape = sheet1.cell_value(855, 3)
                    for i in range(855,1464):
                        if sheet1.cell_value(i, 10) > Min_area and sheet1.cell_value(i, 10) < best_area:
                            best_area = sheet1.cell_value(i, 10)
                            shape = sheet1.cell_value(i, 3)
                elif shape_of_member.get() == "WT9":
                    best_area = sheet1.cell_value(673, 10)
                    shape = sheet1.cell_value(673, 3)
                    for i in range(673,695):
                        if sheet1.cell_value(i, 10) > Min_area and sheet1.cell_value(i, 10) < best_area:
                            best_area = sheet1.cell_value(i, 10)
                            shape = sheet1.cell_value(i, 3)
                elif shape_of_member.get() == "WT12":
                    best_area = sheet1.cell_value(673, 10)
                    shape = sheet1.cell_value(673, 3)
                    for i in range(634,654):
                        if sheet1.cell_value(i, 10) > Min_area and sheet1.cell_value(i, 10) < best_area:
                            best_area = sheet1.cell_value(i, 10)
                            shape = sheet1.cell_value(i, 3)
                elif shape_of_member.get() == "W14":
                    best_area = sheet1.cell_value(673, 10)
                    shape = sheet1.cell_value(673, 3)
                    for i in range(168,203):
                        if sheet1.cell_value(i, 10) > Min_area and sheet1.cell_value(i, 10) < best_area:
                            best_area = sheet1.cell_value(i, 10)
                            shape = sheet1.cell_value(i, 3)
                tension_l1["text"] = "Minimum required gross area (based on yielding) = " + str("%.2f" % (total_load/(.90 * Fy))) + " square inches."
                tension_shape1["text"] = "Based on this minimum gross area, select " + str(shape)
                tension_shape2["text"] = "with area = " + str(best_area) + " inches squared."
                tension_l2["text"] = "Minimum effective net area (based on rupture) = " + str("%.2f" % (total_load/(.75 * Fu))) + " square inches."
                tension_l3["text"] = "The combination of holes and shear lag may not reduce the area of this "
                tension_l4["text"] = "member by more than " + str("%.3f" % (total_load/(.75 * Fu * best_area)))
                
                label2.grid(row=10, column=0, sticky="W", pady = (0,10))
                tension_l1.grid(row=11, column=0, sticky="W", pady = (0,10))
                tension_shape1.grid(row=12, column=0, sticky="W")
                tension_shape2.grid(row=13, column=0, sticky="W", pady = (0,10))
                tension_l2.grid(row=14, column=0, sticky="W", pady = (0,10))
                tension_l3.grid(row = 15, column = 0, sticky = "W")
                tension_l4.grid(row = 16, column = 0, sticky = "W")
    
            if method.get() == "ASD":
                if Entry_4.get() != '':
                    total_load = float(Entry_4.get())
                else: 
                    total_load = float(Entry_1.get()) + float(Entry_2.get())
                display = str("%.2f" % total_load)
                label2["text"] = "Factored load = " + display + " kips"
                
                Min_area = total_load * 1.67 / Fy
                if shape_of_member.get() == "Double-Angle":
                    best_area = sheet1.cell_value(855, 10)
                    shape = sheet1.cell_value(855, 3)
                    for i in range(855,1464):
                        if sheet1.cell_value(i, 10) > Min_area and sheet1.cell_value(i, 10) < best_area:
                            best_area = sheet1.cell_value(i, 10)
                            shape = sheet1.cell_value(i, 3)
                elif shape_of_member.get() == "WT9":
                    best_area = sheet1.cell_value(673, 10)
                    shape = sheet1.cell_value(673, 3)
                    for i in range(673,695):
                        if sheet1.cell_value(i, 10) > Min_area and sheet1.cell_value(i, 10) < best_area:
                            best_area = sheet1.cell_value(i, 10)
                            shape = sheet1.cell_value(i, 3)
                elif shape_of_member.get() == "WT12":
                    best_area = sheet1.cell_value(673, 10)
                    shape = sheet1.cell_value(673, 3)
                    for i in range(634,654):
                        if sheet1.cell_value(i, 10) > Min_area and sheet1.cell_value(i, 10) < best_area:
                            best_area = sheet1.cell_value(i, 10)
                            shape = sheet1.cell_value(i, 3)
                elif shape_of_member.get() == "W14":
                    best_area = sheet1.cell_value(673, 10)
                    shape = sheet1.cell_value(673, 3)
                    for i in range(168,203):
                        if sheet1.cell_value(i, 10) > Min_area and sheet1.cell_value(i, 10) < best_area:
                            best_area = sheet1.cell_value(i, 10)
                            shape = sheet1.cell_value(i, 3)
                tension_l1["text"] = "Minimum required gross area (based on yielding) = " + str("%.2f" % (total_load * 1.67 /Fy)) + " square inches."
                tension_shape1["text"] = "Based on this minimum gross area, select " + str(shape)
                tension_shape2["text"] = "with area = " + str(best_area) + " inches squared."
                tension_l2["text"] = "Minimum effective net area (based on rupture) = " + str("%.2f" % (total_load * 2 /Fu)) + " square inches."
                tension_l3["text"] = "The combination of holes and shear lag may not reduce the area of this "
                tension_l4["text"] = "member by more than " + str("%.3f" % ((total_load * 2 /Fu)/best_area))
                
                label2.grid(row=10, column=0, sticky="W", pady = (0,10))
                tension_l1.grid(row=11, column=0, sticky="W", pady = (0,10))
                tension_shape1.grid(row=12, column=0, sticky="W")
                tension_shape2.grid(row=13, column=0, sticky="W", pady = (0,10))
                tension_l2.grid(row=14, column=0, sticky="W", pady = (0,10))
                tension_l3.grid(row = 15, column = 0, sticky = "W")
                tension_l4.grid(row = 16, column = 0, sticky = "W", )
    
            quote["text"] = dict[random.randint(1,6)]
            quote.place(rely=1, relx=0, x=0, y=0, anchor=SW)
            
    if component.get() == "Tension Members":
        design_button = ttk.Button(newwin, text="Design!",command=Tension_Design,style='TButton')
        design_button.grid(row=18, column=0, sticky = "W")
    elif component.get() == "Compression Members":
        design_button = ttk.Button(newwin, text="Design!", command=Compression_Design,style='TButton')
        design_button.grid(row=18, column=0, sticky = "W")
    elif component.get() == "Bending Members":
        design_button = ttk.Button(newwin, text="Design!", command=Bending_Design,style='TButton')
        design_button.grid(row=14, column=0, sticky = "W")
    elif component.get() == "Connections":
        design_button = ttk.Button(newwin, text="Design!", command=Connections_Design,style='TButton')
        design_button.grid(row=18, column=0, sticky = "W")

#button = ttk.Button(root, text="Next!",command=new_winF,style='TButton')
#ttk.Style().configure("TButton", padding=6, relief="flat",background="#000")
#button.pack()

root.mainloop()
