from tkinter import *
import backend
import pandas as pd
import xlsxwriter
from tkinter import messagebox
from PIL import Image, ImageTk
from datetime import datetime
import difflib
from difflib import get_close_matches
import numpy

date=datetime.today().strftime('%Y-%m-%d')
data = backend.view()
print(data)


def get_selected_row(event):
    try:
        global selected_tuple
        index=list1.curselection()[0]
        selected_tuple=list1.get(index)
        e1.delete(0,END)
        e1.insert(END,selected_tuple[1])
        e2.delete(0,END)
        e2.insert(END,selected_tuple[2])
        e3.delete(0,END)
        e3.insert(END,selected_tuple[3])
        e4.delete(0,END)
        e4.insert(END,selected_tuple[4])
    except IndexError:
        pass

def clear_command():
    e1.delete(0,END)
    e2.delete(0,END)
    e3.delete(0,END)
    e4.delete(0,END)

def view_command():
    list1.delete(0,END)
    for row in backend.view():
        list1.insert(END,row)

def export_command():
    try:
        df = pd.DataFrame(backend.view(),columns=['Book ID','Title','Author','Year','ISBN'])
        writer = pd.ExcelWriter('Book Store report, %s.xlsx' % date, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='1st edition')
        writer.save()
        return messagebox.showinfo("Success","Report has been saved")
    except xlsxwriter.exceptions.FileCreateError:
        return messagebox.showerror("Error","Close the existing Excel file")

def search_command():
    df2 = pd.DataFrame(backend.view(),columns=['Book ID','Title','Author','Year','ISBN'])
    search_result=[]
    for row in backend.search(title_text.get(),author_text.get(),year_text.get(),isbn_text.get()):
        search_result=row
    if len(search_result) > 0:
        list1.delete(0,END)
        for row in backend.search(title_text.get(),author_text.get(),year_text.get(),isbn_text.get()):
            list1.insert(END,row)
    if  len(get_close_matches(title_text.get(), df2.iloc[:,1])) > 0:
        new_title_word = get_close_matches(title_text.get(), df2.iloc[:,1])[0]
        MsgBox = messagebox.askquestion ('Typo error','Did you mean {''%s''} instead?' % new_title_word, icon = 'warning')
        if MsgBox == 'yes':
            list1.delete(0,END)
            for row in backend.search(new_title_word,author_text.get(),year_text.get(),isbn_text.get()):
                list1.insert(END,row)
                clear_command()
        else:
            clear_command()
            return messagebox.showinfo("Return","Try again something different")
    else:
        if len(search_result) == 0:
            list1.delete(0,END)
            clear_command()
            return messagebox.showinfo("No results","Try again something different")


def add_command():
    if len(title_text.get()) == 0 or len(author_text.get()) == 0 or len(year_text.get()) == 0 or len(isbn_text.get()) == 0:
        return messagebox.showerror("Error","Some data is missing!")
    else:
        backend.insert(title_text.get(),author_text.get(),year_text.get(),isbn_text.get())
        list1.delete(0,END)
        list1.insert(END,(title_text.get(),author_text.get(),year_text.get(),isbn_text.get()))
        return messagebox.showinfo("Success","Inserted new row")


def delete_command():
    backend.delete(selected_tuple[0])
    list1.delete(0,END)
    for row in backend.view():
        list1.insert(END,row)
    return messagebox.showinfo("Success","Removed")

def update_command():
    backend.update(selected_tuple[0], title_text.get(),author_text.get(),year_text.get(),isbn_text.get())
    return messagebox.showinfo("Success","Updated!")


window=Tk()
window.configure(background='#6F6D6B')
window.geometry("400x300")
# i = Image.open("image.jpg")
# image1= ImageTk.PhotoImage(i)
# label_for_image= Label(window, image=image1)
# label_for_image.pack()

l1=Label(window,text = "Title", background='#6F6D6B', fg="white", font='Arial 9 bold')
l1.place(relx=0.04, rely=0.05)

l2=Label(window,text = "Author", background='#6F6D6B', fg="white", font='Arial 9 bold')
l2.place(relx=0.04, rely=0.15)

l3=Label(window,text = "Year", background='#6F6D6B', fg="white", font='Arial 9 bold')
l3.place(relx=0.53, rely=0.05)

l4=Label(window,text = "ISBN", background='#6F6D6B', fg="white", font='Arial 9 bold')
l4.place(relx=0.53, rely=0.15)

title_text=StringVar()
e1=Entry(window,textvariable=title_text)
e1.place(relx=0.155, rely=0.05)

author_text=StringVar()
e2=Entry(window,textvariable=author_text)
e2.place(relx=0.156, rely=0.15)

year_text=StringVar()
e3=Entry(window,textvariable=year_text)
e3.place(relx=0.625, rely=0.05)

isbn_text=StringVar()
e4=Entry(window,textvariable=isbn_text)
e4.place(relx=0.625, rely=0.15)

list1=Listbox(window)
list1.place(relx=0.05, rely=0.27, relheight=0.682, relwidth=0.55)

sb1=Scrollbar(window, orient="vertical", activebackground='#6F6D6B', troughcolor='#6F6D6B')
sb1.place(relx=0.56, rely=0.27, relheight=0.682, relwidth=0.05)

list1.configure(yscrollcommand=sb1.set)
sb1.configure(command=list1.yview)

list1.bind('<<ListboxSelect>>',get_selected_row)

b1=Button(window,text="View all", width=13,command=view_command)
b1.place(relx=0.682, rely=0.27)

b2=Button(window,text="Search entry", width=13,command=search_command)
b2.place(relx=0.682, rely=0.35)

b3=Button(window,text="Add entry", width=13,command=add_command)
b3.place(relx=0.682, rely=0.43)

b4=Button(window,text="Update selected", width=13,command=update_command)
b4.place(relx=0.682, rely=0.51)

b5=Button(window,text="Delete selected", width=13,command=delete_command)
b5.place(relx=0.682, rely=0.591)

b6=Button(window,text="Export to Excel", width=13,command=export_command)
b6.place(relx=0.682, rely=0.671)

b7=Button(window,text="Clear", width=13,command=clear_command)
b7.place(relx=0.682, rely=0.751)

b8=Button(window,text="Close", width=13,command=window.destroy)
b8.place(relx=0.682, rely=0.8315)
#
# entry_list = [child for child in window.winfo_children()
#               if isinstance(child, Entry)]



# entries = []
# for i in range(4):
#     entry = Entry()
#     entries.append(entry)
# print(entries)

# print(entry_list)
window.mainloop()
