'''NAME : ANURAG KAVI SHANKAR TAHKUR
CLAS: XI A,  ROLL. NO : 6, 
DAV INTERNATIONAL SCHOOL
COMPUTER SCIENCE PROJECT (2021-2022)'''

'''PROCESS TO INTALL A LIBRARY
1. RUN CMD AS ADMINISTRATOR
2. TYPE : "pip install <name of library>"
3. PRESS ENTER, THEN WAIT FOR IT TO INSTALL.'''

'''Note: IF YOU WANT TO RN THE CODE YOU WILL HAVE 
TO DOWNLOAD AND KEEP THE FILE NAMED ‘Sales data.xlsx’ 
PRESENT IN THE GIVEN LINK.
Link: https://github.com/anuragthakur2102/CS-PROJECT-CLASS-11'''

from tkinter import *
import pandas as pd
import matplotlib.pyplot as p
import os


l=[]
l_2=[]

root = Tk()
root.title("Excel Sheet")
root.geometry("600x350")

#Warning
label = Label(text = "**IF YOU DON'T HAVE ANY DATA, WE CAN PROVIDE YOU WITH THE AVAILABLE DATA**" ,bg= "white",padx = 50,font="arial 10 bold")
label.pack(side = BOTTOM)

#Question

quest= Label(text = "DO YOU HAVE ANY DATA OF YOUR OWN?", fg = "black", font = "arial 15 bold")
quest.pack(side=TOP, pady=10)

def click(event):
    text=event.widget.cget("text")
    l.append(text)
    if (l[0]=='YES'):
        print('''Your Excel Sheet must be of this form
    ----------------------------------------    
    ROW    TITLE ''')
        data=(pd.read_excel('Sales data.xlsx'))
        print(pd.DataFrame(data))
        

    elif (l[0]=="NO"):
        print('We have Sales Data of three comapnies: MICROSOFT, META, ALPHABET')
        print()
        comp=int(input('''Which one do you want
        1. MICROSOFT- *(TYPE 1)*
        1. ALPHABET- *(TYPE 2)*
        1. META- *(TYPE 3)*
        INPUT HERE : '''))

def yes():
    print(" SO, YOU HAVE CHOSEN TO INPUT YOUR DATA")
    print()
    print('''Your Excel Sheet must be of this form
-----------------------------------------    
ROW    TITLE ''')
    data=(pd.read_excel('Sales data.xlsx'))
    print(pd.DataFrame(data))
    root.destroy()
    print()
    print("***NOTE: THE PROGRAM CAN COMPARE ONLY 5 COMPANIES AT A TIME***")
    print()
    confirm=int(input("Enter 0 to continue after reading the instructions :"))
    if confirm!=0:
        print("Oops! You will have to run the program again and follow thw steps properly")
        print()
        print()
    else:
        path_root=Tk()
        path_root.title("PATH")
        path_root.geometry("600x400")

        #instruction
        inst= Label(path_root, text = '''***SAVE YOUR FILE IN THE GIVEN DIRECTORY IN SAME FOLDER IN WHICH PROGRAM IS SAVED
        PLEASE MAKE SURE THAT YOUR SHEET NAME AND HEADER AT 0 IS SAME AND HEADER AT 1 IS ANALYSIS ''', bg="white", fg="red", font="areal 9 bold")
        inst.pack(side=BOTTOM, pady=20)
        quest_2= Label(path_root, text="Enter name of your file with .xlsx at the end ", fg="black", font="areal 13")
        quest_2.pack(side=TOP,pady=20)
        print()

        
        
        
    #taking all the inputs from user
    def okay():
        print(" Path Taken. Now searching for data......")
        path1=p_entry.get()
        path_root.destroy()
        assert os.path.exists(path1), " I did not find your file at, "+str(path1)
        f=open(path1)
        print("Hooray! We found your file.")
        
        #Now reading the file
        file=pd.ExcelFile(path1)
        sheets_1=file.sheet_names
        if len(sheets_1)==1:
            sheet_root=Tk()
            sheet_root.title("SHEET NAME")
            sheet_root.geometry("700x500")

            quest_4= Label(text = "IN WHICH COMAPNIES REVENUE ARE YOU INTERESTED IN?", fg = "black", font = "arial 15 bold")
            quest_4.pack(side=TOP, pady=10)
            
            def micro():
                print("Looks like you are interested in", sheets_1[0])
                print()
                
                #graph
                #using pandas to get specific rows now
                #x axis
                y_1=pd.read_excel(path1, sheet_name=sheets_1[0])
                year=y_1[sheets_1[0]].tolist()
                year=year[::-1]
                year.pop()
                
                #y axis
                r_1=pd.read_excel(path1, sheet_name=sheets_1[0])
                revenue=r_1["ANALYSIS"].tolist()
                revenue=revenue[::-1]
                revenue.pop()
                
                #making graph
                x=year
                y=revenue
                p.plot(x,y, color="deepskyblue", linewidth=2, marker='.',markersize=15, label=sheets_1[0])
                p.xlabel("YEAR")
                p.ylabel("REVENUE (BILLIONS OF US $)")
                p.title("REVENUE TREND")
                p.legend()
                p.show()
                

            def ex():
                sheet_root.destroy()
                print('''THANK YOU
        VISIT AGAIN''')

            m= Button(sheet_root, text = sheets_1[0], padx= 40 , pady =20 ,command =micro , font="areal 15", bg="deepskyblue")
            m.pack(side=LEFT,padx=25)
            e= Button(sheet_root, text = " EXIT ",padx= 25 , pady =20 ,command =ex, font="areal 20", fg="white", bg="black")
            e.pack(side=BOTTOM, pady=15)
        
        if len(sheets_1)==2:
            sheet_root=Tk()
            sheet_root.title("SHEET NAME")
            sheet_root.geometry("700x500")

            quest_4= Label(text = "IN WHICH COMAPNIES REVENUE ARE YOU INTERESTED IN?", fg = "black", font = "arial 15 bold")
            quest_4.pack(side=TOP, pady=10)
            
            def micro():
                print("Looks like you are interested in", sheets_1[0])
                print()
                
                #graph
                #using pandas to get specific rows now
                #x axis
                y_1=pd.read_excel(path1, sheet_name=sheets_1[0])
                year=y_1[sheets_1[0]].tolist()
                year=year[::-1]
                year.pop()
                
                #y axis
                r_1=pd.read_excel(path1, sheet_name=sheets_1[0])
                revenue=r_1["ANALYSIS"].tolist()
                revenue=revenue[::-1]
                revenue.pop()
                
                #making graph
                x=year
                y=revenue
                p.plot(x,y, color="deepskyblue", linewidth=2, marker='.',markersize=15, label=sheets_1[0])
                p.xlabel("YEAR")
                p.ylabel("REVENUE (BILLIONS OF US $)")
                p.title("REVENUE TREND")
                p.legend()
                p.show()
                

            def alpha():
                print("Looks like you are interested in", sheets_1[1])
                print()
                
                #graph
                #using pandas to get specific rows now
                #x axis
                y_1=pd.read_excel(path1, sheet_name=sheets_1[1])
                year=y_1[sheets_1[1]].tolist()
                year=year[::-1]
                year.pop()
                #y axis
                r_1=pd.read_excel(path1, sheet_name=sheets_1[1])
                revenue=r_1["ANALYSIS"].tolist()
                revenue=revenue[::-1]
                revenue.pop()
                
                #making graph
                x=year
                y=revenue
                p.plot(x,y, color="greenyellow",linewidth=2, marker='.',markersize=15, label=sheets_1[1])
                p.xlabel("YEAR")
                p.ylabel("REVENUE (BILLIONS OF US $)")
                p.title("REVENUE TREND")
                p.legend()
                p.show()
                

            def ex():
                sheet_root.destroy()
                print('''THANK YOU
        VISIT AGAIN''')

            m= Button(sheet_root, text = sheets_1[0], padx= 40 , pady =20 ,command =micro , font="areal 15", bg="deepskyblue")
            m.pack(side=LEFT,padx=25)
            a= Button(sheet_root, text = sheets_1[1], padx= 40 , pady =20 ,command =alpha, font="areal 15", bg="yellowgreen")
            a.pack(side=RIGHT,padx=25)
            e= Button(sheet_root, text = " EXIT ",padx= 25 , pady =20 ,command =ex, font="areal 20", fg="white", bg="black")
            e.pack(side=BOTTOM, pady=15)

        if len(sheets_1)==3:
            sheet_root=Tk()
            sheet_root.title("SHEET NAME")
            sheet_root.geometry("700x500")

            quest_4= Label(text = "IN WHICH COMAPNIES REVENUE ARE YOU INTERESTED IN?", fg = "black", font = "arial 15 bold")
            quest_4.pack(side=TOP, pady=10)
            
            def micro():
                print("Looks like you are interested in", sheets_1[0])
                print()
                
                #graph
                #using pandas to get specific rows now
                #x axis
                y_1=pd.read_excel(path1, sheet_name=sheets_1[0])
                year=y_1[sheets_1[0]].tolist()
                year=year[::-1]
                year.pop()
                
                #y axis
                r_1=pd.read_excel(path1, sheet_name=sheets_1[0])
                revenue=r_1["ANALYSIS"].tolist()
                revenue=revenue[::-1]
                revenue.pop()
                
                #making graph
                x=year
                y=revenue
                p.plot(x,y, color="deepskyblue",linewidth=2, marker='.',markersize=15, label=sheets_1[0])
                p.xlabel("YEAR")
                p.ylabel("REVENUE (BILLIONS OF US $)")
                p.title("REVENUE TREND")
                p.legend()
                p.show()
                

            def alpha():
                print("Looks like you are interested in", sheets_1[1])
                print()
                
                #graph
                #using pandas to get specific rows now
                #x axis
                y_1=pd.read_excel(path1, sheet_name=sheets_1[1])
                year=y_1[sheets_1[1]].tolist()
                year=year[::-1]
                year.pop()
                #y axis
                r_1=pd.read_excel(path1, sheet_name=sheets_1[1])
                revenue=r_1["ANALYSIS"].tolist()
                revenue=revenue[::-1]
                revenue.pop()
                
                #making graph
                x=year
                y=revenue
                p.plot(x,y, color="greenyellow",linewidth=2, marker='.',markersize=15,label=sheets_1[1])
                p.xlabel("YEAR")
                p.ylabel("REVENUE (BILLIONS OF US $)")
                p.title("REVENUE TREND")
                p.legend()
                p.show()
                

            def met():
                print("Looks like you are interested in", sheets_1[2])
                print()
                
                #graph
                #using openpyxl to get specific rows now
                y_1=pd.read_excel(path1, sheet_name=sheets_1[2])
                year=y_1[sheets_1[2]].tolist()
                year=year[::-1]
                year.pop()
                #y axis
                r_1=pd.read_excel(path1, sheet_name=sheets_1[2])
                revenue=r_1["ANALYSIS"].tolist()
                revenue=revenue[::-1]
                revenue.pop()
                
                #making graph
                x=year
                y=revenue
                p.plot(x,y, color="coral",linewidth=2, marker='.',markersize=15,label=sheets_1[2])
                p.xlabel("YEAR")
                p.ylabel("REVENUE (BILLIONS OF US $)")
                p.title("REVENUE TREND")
                p.legend()
                p.show()
                
            
            def ex():
                sheet_root.destroy()
                print('''THANK YOU
        VISIT AGAIN''')

            m= Button(sheet_root, text = sheets_1[0], padx= 40 , pady =20 ,command =micro , font="areal 15", bg="deepskyblue")
            m.pack(side=LEFT,padx=25)
            a= Button(sheet_root, text = sheets_1[1], padx= 40 , pady =20 ,command =alpha, font="areal 15", bg="yellowgreen")
            a.pack(side=RIGHT,padx=25)
            e= Button(sheet_root, text = " EXIT ",padx= 25 , pady =20 ,command =ex, font="areal 20", fg="white", bg="black")
            e.pack(side=BOTTOM, pady=15)
            me= Button(sheet_root, text = sheets_1[2], padx= 40 , pady= 20 ,command =met, font="areal 15", bg="coral")
            me.pack(side=BOTTOM, pady=62)

        if len(sheets_1)==4:
            sheet_root=Tk()
            sheet_root.title("SHEET NAME")
            sheet_root.geometry("700x500")

            quest_4= Label(text = "IN WHICH COMAPNIES REVENUE ARE YOU INTERESTED IN?", fg = "black", font = "arial 15 bold")
            quest_4.pack(side=TOP, pady=10)
            
            def micro():
                print("Looks like you are interested in", sheets_1[0])
                print()
                
                #graph
                #using pandas to get specific rows now
                #x axis
                y_1=pd.read_excel(path1, sheet_name=sheets_1[0])
                year=y_1[sheets_1[0]].tolist()
                year=year[::-1]
                year.pop()
                
                #y axis
                r_1=pd.read_excel(path1, sheet_name=sheets_1[0])
                revenue=r_1["ANALYSIS"].tolist()
                revenue=revenue[::-1]
                revenue.pop()
                
                #making graph
                x=year
                y=revenue
                p.plot(x,y, color="deepskyblue",linewidth=2, marker='.',markersize=15, label=sheets_1[0])
                p.xlabel("YEAR")
                p.ylabel("REVENUE (BILLIONS OF US $)")
                p.title("REVENUE TREND")
                p.legend()
                p.show()
                

            def alpha():
                print("Looks like you are interested in", sheets_1[1])
                print()
                
                #graph
                #using pandas to get specific rows now
                #x axis
                y_1=pd.read_excel(path1, sheet_name=sheets_1[1])
                year=y_1[sheets_1[1]].tolist()
                year=year[::-1]
                year.pop()
                #y axis
                r_1=pd.read_excel(path1, sheet_name=sheets_1[1])
                revenue=r_1["ANALYSIS"].tolist()
                revenue=revenue[::-1]
                revenue.pop()
                
                #making graph
                x=year
                y=revenue
                p.plot(x,y, color="greenyellow",linewidth=2, marker='.',markersize=15,label=sheets_1[1])
                p.xlabel("YEAR")
                p.ylabel("REVENUE (BILLIONS OF US $)")
                p.title("REVENUE TREND")
                p.legend()
                p.show()
                

            def met():
                print("Looks like you are interested in", sheets_1[2])
                print()
                
                #graph
                #using openpyxl to get specific rows now
                y_1=pd.read_excel(path1, sheet_name=sheets_1[2])
                year=y_1[sheets_1[2]].tolist()
                year=year[::-1]
                year.pop()
                #y axis
                r_1=pd.read_excel(path1, sheet_name=sheets_1[2])
                revenue=r_1["ANALYSIS"].tolist()
                revenue=revenue[::-1]
                revenue.pop()
                
                #making graph
                x=year
                y=revenue
                p.plot(x,y, color="coral",linewidth=2, marker='.',markersize=15,label=sheets_1[2])
                p.xlabel("YEAR")
                p.ylabel("REVENUE (BILLIONS OF US $)")
                p.title("REVENUE TREND")
                p.legend()
                p.show()
                

            def beta():
                print("Looks like you are interested in", sheets_1[3])
                print()
                
                #graph
                #using openpyxl to get specific rows now
                y_1=pd.read_excel(path1, sheet_name=sheets_1[3])
                year=y_1[sheets_1[3]].tolist()
                year=year[::-1]
                year.pop()
                #y axis
                r_1=pd.read_excel(path1, sheet_name=sheets_1[3])
                revenue=r_1["ANALYSIS"].tolist()
                revenue=revenue[::-1]
                revenue.pop()
                
                #making graph
                x=year
                y=revenue
                p.plot(x,y, color="slategray",linewidth=2, marker='.',markersize=15,label=sheets_1[3])
                p.xlabel("YEAR")
                p.ylabel("REVENUE (BILLIONS OF US $)")
                p.title("REVENUE TREND")
                p.legend()
                p.show()
                

            def ex():
                sheet_root.destroy()
                print('''THANK YOU
        VISIT AGAIN''')
            
            m= Button(sheet_root, text = sheets_1[0], padx= 40 , pady =20 ,command =micro , font="areal 15", bg="deepskyblue")
            m.pack(side=LEFT,padx=25)
            a= Button(sheet_root, text = sheets_1[1], padx= 40 , pady =20 ,command =alpha, font="areal 15", bg="yellowgreen")
            a.pack(side=RIGHT,padx=25)
            e= Button(sheet_root, text = " EXIT ",padx= 25 , pady =20 ,command =ex, font="areal 20", fg="white", bg="black")
            e.pack(side=BOTTOM, pady=15)
            me= Button(sheet_root, text = sheets_1[2], padx= 40 , pady= 20 ,command =met, font="areal 15", bg="coral")
            me.pack(side=BOTTOM, pady=62)
            n= Button(sheet_root, text = sheets_1[3], padx= 40 , pady =20 ,command =beta, font="areal 15", bg="slategray")
            n.pack(side=RIGHT,padx=25)

        if len(sheets_1)==5:
            sheet_root=Tk()
            sheet_root.title("SHEET NAME")
            sheet_root.geometry("700x600")

            quest_4= Label(text = "IN WHICH COMAPNIES REVENUE ARE YOU INTERESTED IN?", fg = "black", font = "arial 15 bold")
            quest_4.pack(side=TOP, pady=10)
            
            def micro():
                print("Looks like you are interested in", sheets_1[0])
                print()
                
                #graph
                #using pandas to get specific rows now
                #x axis
                y_1=pd.read_excel(path1, sheet_name=sheets_1[0])
                year=y_1[sheets_1[0]].tolist()
                year=year[::-1]
                year.pop()
                
                #y axis
                r_1=pd.read_excel(path1, sheet_name=sheets_1[0])
                revenue=r_1["ANALYSIS"].tolist()
                revenue=revenue[::-1]
                revenue.pop()
                
                #making graph
                x=year
                y=revenue
                p.plot(x,y, color="deepskyblue",linewidth=2, marker='.',markersize=15, label=sheets_1[0])
                p.xlabel("YEAR")
                p.ylabel("REVENUE (BILLIONS OF US $)")
                p.title("REVENUE TREND")
                p.legend()
                p.show()
                

            def alpha():
                print("Looks like you are interested in", sheets_1[1])
                print()
                
                #graph
                #using pandas to get specific rows now
                #x axis
                y_1=pd.read_excel(path1, sheet_name=sheets_1[1])
                year=y_1[sheets_1[1]].tolist()
                year=year[::-1]
                year.pop()
                #y axis
                r_1=pd.read_excel(path1, sheet_name=sheets_1[1])
                revenue=r_1["ANALYSIS"].tolist()
                revenue=revenue[::-1]
                revenue.pop()
                
                #making graph
                x=year
                y=revenue
                p.plot(x,y, color="greenyellow",linewidth=2, marker='.',markersize=15, label=sheets_1[1])
                p.xlabel("YEAR")
                p.ylabel("REVENUE (BILLIONS OF US $)")
                p.title("REVENUE TREND")
                p.legend()
                p.show()
                

            def met():
                print("Looks like you are interested in", sheets_1[2])
                print()
                
                #graph
                #using openpyxl to get specific rows now
                y_1=pd.read_excel(path1, sheet_name=sheets_1[2])
                year=y_1[sheets_1[2]].tolist()
                year=year[::-1]
                year.pop()
                #y axis
                r_1=pd.read_excel(path1, sheet_name=sheets_1[2])
                revenue=r_1["ANALYSIS"].tolist()
                revenue=revenue[::-1]
                revenue.pop()
                
                #making graph
                x=year
                y=revenue
                p.plot(x,y, color="coral",linewidth=2, marker='.',markersize=15,label=sheets_1[2])
                p.xlabel("YEAR")
                p.ylabel("REVENUE (BILLIONS OF US $)")
                p.title("REVENUE TREND")
                p.legend()
                p.show()
                

            def beta():
                print("Looks like you are interested in", sheets_1[3])
                print()
                
                #graph
                #using openpyxl to get specific rows now
                y_1=pd.read_excel(path1, sheet_name=sheets_1[3])
                year=y_1[sheets_1[3]].tolist()
                year=year[::-1]
                year.pop()
                #y axis
                r_1=pd.read_excel(path1, sheet_name=sheets_1[3])
                revenue=r_1["ANALYSIS"].tolist()
                revenue=revenue[::-1]
                revenue.pop()
                
                #making graph
                x=year
                y=revenue
                p.plot(x,y, color="slategray",linewidth=2, marker='.',markersize=15, label=sheets_1[3])
                p.xlabel("YEAR")
                p.ylabel("REVENUE (BILLIONS OF US $)")
                p.title("REVENUE TREND")
                p.legend()
                p.show()
                

            def gama():
                print("Looks like you are interested in", sheets_1[4])
                print()
                
                #graph
                #using openpyxl to get specific rows now
                y_1=pd.read_excel(path1, sheet_name=sheets_1[4])
                year=y_1[sheets_1[4]].tolist()
                year=year[::-1]
                year.pop()
                #y axis
                r_1=pd.read_excel(path1, sheet_name=sheets_1[4])
                revenue=r_1["ANALYSIS"].tolist()
                revenue=revenue[::-1]
                revenue.pop()
                
                #making graph
                x=year
                y=revenue
                p.plot(x,y, color="magenta",linewidth=2, marker='.',markersize=15,label=sheets_1[4])
                p.xlabel("YEAR")
                p.ylabel("REVENUE (BILLIONS OF US $)")
                p.title("REVENUE TREND")
                p.legend()
                p.show()
                
            def ex():
                sheet_root.destroy()
                print('''THANK YOU
        VISIT AGAIN''')

            
            e= Button(sheet_root, text = " EXIT ",padx= 30 , pady =20 ,command =ex, font="areal 20", fg="white", bg="black")
            e.pack(side=BOTTOM, padx=25, pady=15)
            me= Button(sheet_root, text = sheets_1[2], padx= 40 , pady= 20 ,command =met, font="areal 15", bg="coral")
            me.pack(side=RIGHT,padx=20, pady=20)
            n= Button(sheet_root, text = sheets_1[3], padx= 40 , pady =20 ,command =beta, font="areal 15", bg="slategray")
            n.pack(side=LEFT,padx=25, pady=30)
            n2= Button(sheet_root, text = sheets_1[4], padx= 40 , pady =20 ,command =gama, font="areal 15", bg="magenta")
            n2.pack(side=BOTTOM,padx=25, pady=30)
            m= Button(sheet_root, text = sheets_1[0], padx= 40 , pady =20 ,command =micro , font="areal 15", bg="deepskyblue")
            m.pack(side=TOP,pady=30)
            a= Button(sheet_root, text = sheets_1[1], padx= 40 , pady =20 ,command =alpha, font="areal 15", bg="yellowgreen")
            a.pack(side=BOTTOM,padx=35, pady=33)
        
            
    p_entry=Entry(path_root, width=50)
    p_entry.pack()
    okay_1=Button(path_root, text="DONE", fg="black", bg="white", command=okay, font="areal 12")
    okay_1.pack(pady=30)

def no():
    print(" TAKING AVAILABLE DATA TO PROCEED FROM HERE")
    print()
    root.destroy()
    
    #button for company
    comp_root=Tk()
    comp_root.title("COMPANIES")
    comp_root.geometry("700x500")

    quest_3= Label(text = "IN WHICH COMAPNIES REVENUE ARE YOU INTERESTED IN?", fg = "black", font = "arial 15 bold")
    quest_3.pack(side=TOP, pady=10)

    def micro():
        print("Looks like you are interested in MICROSOFT")
        print()
        
        #graph
        #using pandas to get specific rows now
        #x axis
        y_1=pd.read_excel('Sales data.xlsx', sheet_name="MICROSOFT")
        year=y_1["MICROSOFT"].tolist()
        year=year[::-1]
        year.pop()
        #y axis
        r_1=pd.read_excel('Sales data.xlsx', sheet_name="MICROSOFT")
        revenue=r_1["ANALYSIS"].tolist()
        revenue=revenue[::-1]
        revenue.pop()
        
        #making graph
        x=year
        y=revenue
        p.plot(x,y, color="deepskyblue",linewidth=2, marker='.',markersize=15, label="MICROSOFT")
        p.xlabel("YEAR")
        p.ylabel("REVENUE (BILLIONS OF US $)")
        p.title("REVENUE TREND")
        p.legend()
        p.show()

    def alpha():
        print("Looks like you are interested in ALPHABET")
        print()
        
        #graph
        #using pandas to get specific rows now
        #x axis
        y_1=pd.read_excel('Sales data.xlsx', sheet_name="ALPHABET")
        year=y_1["ALPHABET"].tolist()
        year=year[::-1]
        year.pop()
        #y axis
        r_1=pd.read_excel('Sales data.xlsx', sheet_name="ALPHABET")
        revenue=r_1["ANALYSIS"].tolist()
        revenue=revenue[::-1]
        revenue.pop()
        
        #making graph
        x=year
        y=revenue
        p.plot(x,y, color="greenyellow",linewidth=2, marker='.',markersize=15,label="ALPHABET")
        p.xlabel("YEAR")
        p.ylabel("REVENUE (BILLIONS OF US $)")
        p.title("REVENUE TREND")
        p.legend()
        p.show()
    
    def met():
        print("Looks like you are interested in META")
        print()
        
        #graph
        #using openpyxl to get specific rows now
        y_1=pd.read_excel('Sales data.xlsx', sheet_name="META")
        year=y_1["META"].tolist()
        year=year[::-1]
        year.pop()
        #y axis
        r_1=pd.read_excel('Sales data.xlsx', sheet_name="META")
        revenue=r_1["ANALYSIS"].tolist()
        revenue=revenue[::-1]
        revenue.pop()
        
        #making graph
        x=year
        y=revenue
        p.plot(x,y, color="coral",linewidth=2, marker='.',markersize=15, label="META")
        p.xlabel("YEAR")
        p.ylabel("REVENUE (BILLIONS OF US $)")
        p.title("REVENUE TREND")
        p.legend()
        p.show()
    
    def ex():
        comp_root.destroy()
        print('''THANK YOU
        VISIT AGAIN''')

    m= Button(comp_root, text = " MICROSOFT ",padx= 40 , pady =20 ,command =micro , font="areal 15", bg="deepskyblue")
    m.pack(side=LEFT,padx=25)
    a= Button(comp_root, text = " ALPHABET ",padx= 40 , pady =20 ,command =alpha, font="areal 15", bg="yellowgreen")
    a.pack(side=RIGHT,padx=25)
    e= Button(comp_root, text = " EXIT ",padx= 25 , pady =20 ,command =ex, font="areal 20", fg="white", bg="black")
    e.pack(side=BOTTOM, pady=15)
    me= Button(comp_root, text = " META ",padx= 40 , pady= 20 ,command =met, font="areal 15", bg="coral")
    me.pack(side=BOTTOM, pady=62)
    
b1= Button(root, text = " YES ",padx= 50 , pady =30 ,command =yes, font="areal 20", bg="cyan")
b1.pack(side=LEFT,padx=40)
b1.bind("<Button-1>", click)

b2 = Button(root, text = " NO ",padx= 50 , pady = 30,command = no, font="areal 20", bg="orange")
b2.pack(side=RIGHT,padx=40)
b2.bind("<Button-1>", click)

root.mainloop()

print("DONE")
