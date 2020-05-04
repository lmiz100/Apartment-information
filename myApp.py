#make exe file command: pyinstaller.exe --onefile -w --icon=apartmentIcon.ico myApp.py

import pandas as pd
from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
import os
import sys



def quit_program():
    sys.exit()
	

# find file name includes "סטאטוס דירות" with latest version
def get_file_name():
    max_file = ""	#max lexicographic name (latest version)

    for fname in os.listdir(dir):
        if "סטאטוס דירות" in fname:
            if fname > max_file: max_file = fname

    return max_file

   

dir = 'enter your dicrectory path'
try:
    file_name = get_file_name()
except:
    file_name = ""
file_path = path = os.path.join(dir, file_name)
input_entry = None
output_label = None
historyList = None
history_search = []

#select file manualy when can't find path
if not os.path.isfile(file_path):
    file_path = askopenfilename(title = "Select file", filetypes = [("Excel file","*.xlsx"),("Excel file 97-2003","*.xls")])

try:
	apartments = pd.read_excel(file_path, usecols='A:J')
except FileNotFoundError:
	messagebox.showerror(title="Error", message="File doesn't exist!!")
	quit_program()
	
#remove whitespaces
apartments.rename(str.strip, axis='columns', inplace=True)
apartments.fillna('פרט חסר', inplace=True)

    
def write2screen(err, msg):
    if err == 1:
        output_label.config(font=(1), fg='#ff0000')
    else: output_label.config(font=(1), fg='#0000ff')

    output_label["text"] = msg

    return


def getInfo():
    number = input_entry.get()

    if number.strip() == "":
        write2screen(1, "הכנס מספר דירה לפני החיפוש")
    else:
        try:
            number = int(number)
        except ValueError:
            number = input_entry.get()

        filter_rec = apartments[apartments['מספר דירה'].isin([number])]

        dicts = filter_rec.to_dict('records')

        if len(dicts) == 0: 
            write2screen(1, "לא קיימת דירה מספר {} במאגר".format(number))
            return

        to_print =""
        apartment_dict = dicts[0]
        if "ריקה" == apartment_dict['סטאטוס איכלוס']:
            to_print = "דירה {} ריקה, בדוק את קובץ הסטאטוס".format(number)
            write2screen(1, to_print)
        else:
            owner = apartment_dict['בעלות']
            name = apartment_dict['שם הדייר']
            management_deal = apartment_dict['מסלול דמי ניהול']

            to_print = "דירה מספר {} של הדייר {} נמצאת בבעלות {}, תחת מסלול ניהול {}".format(number, name,owner,management_deal)

            write2screen(0, to_print)
            
        update_history(number, to_print)

    return  


def update_history(number, str):
    global history_search, historyList

    #keep every record apear only once in the list
    if number in history_search:
        ind = history_search.index(number)

        history_search.remove(number)
        history_search.insert(0, number)
        historyList.delete(ind)
        historyList.insert(0, str)
    else:
        history_search.insert(0, number)
        historyList.insert(0, str)

    return    


def clear_history():
    global historyList, history_search

    if historyList.size() > 0:
        historyList.delete(0, historyList.size()-1)
        history_search = []  

    return



class Window(Frame):

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master = master
        self.init_window()

    #Creation of init_window
    def init_window(self):

        # changing the title of our master widget
        self.master.title("Apartment information")

        # allowing the widget to take the full space of the root window
        self.pack(fill=BOTH, expand=1)

        # creating a button instance
        quitButton = Button(self, text="יציאה", command=quit_program, bg="light blue", height="2", width="10", activebackground="light yellow", font=("Helvetica", "12","bold"))
        # placing the button on my window
        quitButton.place(x=555, y=5)

        check_apartment = Button(self, text="בדוק דירה", command=getInfo, bg="light blue", height="2", width="10", activebackground="light yellow", font=("Helvetica", "12","bold"))
        check_apartment.place(x=443, y=5)

        clearButton = Button(self, text="מחק היסטוריה", command=clear_history, bg="light blue", height="2", width="10", activebackground="light yellow", font=("Helvetica", "12","bold"))
        clearButton.place(x=330, y=5)

        global input_entry, output_label, historyList
        output_label = Label(self, font=(1), fg='#0000ff')
        output_label.place(relx=0.5, y=85, anchor=CENTER)

        Label(self, text="הכנס מספר דירה", font=("Helvetica", "12")).place(relx=0.5, y=115)
        input_entry = Entry(self, width=25)
        input_entry.place(relx=0.3, y=115)
        input_entry.focus()

        Label(self, text="היסטוריית חיפוש", font=("Helvetica", "12")).place(relx=0.5, y=150)
        scrollbar = Scrollbar(self)
        scrollbar.pack(side=RIGHT, fill=Y)
        historyList = Listbox(self, yscrollcommand=scrollbar.set, width=90, fg='#ff8000')
        historyList.place(relx=0.1, y=170)
        scrollbar.config(command=historyList.yview)
        

		
        
root = Tk()

#size of the window
root.geometry("850x350")

app = Window(root)
root.mainloop()
