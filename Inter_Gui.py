from tkinter import *
from tkinter import ttk, filedialog, messagebox as mb
import mysql.connector as dbcon
import csv
import os
from encodings import utf_8

def connect_db():
    host = 'localhost'
    user = 'root'
    password = ""

    global myDBcon
    try:
        myDBcon = dbcon.connect(
            host = host,
            user = user,
            password = password
            )

    except dbcon.Error as err:
        print("Something went wrong: {}".format(err))
        try:
            print( "MySQL Error [%d]: %s" % (err.args[0], err.args[1]))
            return None
        except IndexError:
            print ("MySQL Error: %s" % str(err))
            return None
    except TypeError:
        print(err)
        return None
    except ValueError:
        print(err)
        return None
    return myDBcon

def createtab():
    
    def get_all_data():
        connect_db()
        mycursor = myDBcon.cursor()
        zumDB = "USE interessenten_db;" 
        mycursor.execute(zumDB)
        get_data = "SELECT * FROM interessenten;"
        mycursor.execute(get_data)
        global all_data
        all_data = mycursor.fetchall()
        
        for data in all_data:
            tab.insert("", "end", values=(data[0], data[1], data[2], data[3], data[4], data[5], data[6], data[7], data[8], data[9], data[10]))

    global tab_frame
    tab_frame = Frame(main, relief= "flat", width= 1450)
    tab_frame.pack()
    global vsb
    vsb = Scrollbar(tab_frame, orient="vertical")
    vsb.pack(side="right", anchor = "e", fill="y", expand=True)
    global tab
    tab = ttk.Treeview(tab_frame, columns=(1,2,3,4,5,6,7,8,9,10,11), height=34, show="headings", yscrollcommand= vsb.set, )
    
    tab.column(1, width=40)
    tab.column(2, width=100)
    tab.column(3, width=80)
    tab.column(4, width=140)
    tab.column(5, width=80)
    tab.column(6, width=100)
    tab.column(7, width=100)
    tab.column(8, width=100)
    tab.column(9, width=140)
    tab.column(10, width=300)
    tab.column(11, width=200)

    tab.heading(1, text = "ID")
    tab.heading(2, text = "Standort")
    tab.heading(3, text = "PLZ")
    tab.heading(4, text = "Adresse_Standort")
    tab.heading(5, text = "KFZ")
    tab.heading(6, text = "Abteilung")
    tab.heading(7, text = "Anrede")
    tab.heading(8, text = "Vorname")
    tab.heading(9, text = "Nachname")
    tab.heading(10, text = "EMail")
    tab.heading(11, text = "Benutzeranmeldename")
    
    get_all_data()
    
    tab.pack()
    tab.bind("<Button-1>", onselect)
    tab.bind("<Key> <Button-1>", pressed)
    
    vsb.configure(command=tab.yview)
   
global focus_list
global index_list
global checker
focus_list = []
index_list = []
checker = 0

def converter(flst, tab):
    global checker
    fl = flst
    t = tab
    index_min = 0
    index_max = 0

    if checker == 0:
        for x in fl:
            value = t.item(x)['values']
            index_list.append(value[0])
    elif checker == 1:
        index_list.clear()
        for x in fl:
            value = t.item(x)['values']
            index_list.append(value[0])
    elif checker == 2:
        index_list.clear()
        for x in fl: 
            value = t.item(x)['values']
            index_list.append(value[0])
    else: pass
    
    if checker == 0:
        pass
    elif checker == 1:
        index_list.sort()
        index_min = index_list[0]
        index_max = index_list[-1]
        summe = index_max - index_min
        i = 0
        index_list.clear()
        for i in range(summe + 1):
            index_list.append(index_min + i)        
    elif checker == 2:
        index_list.sort()
    else: pass

def pressed(evt):
    global checker
    if evt.state == 8: # MouseClick
        index_list.clear()
        focus_list.clear()
        checker = 0
        _iid = tab.identify_row(evt.y)
        focus_list.append(_iid)
    elif evt.state == 9: # Shift
        checker = 1
        _iid = tab.identify_row(evt.y)
        focus_list.append(_iid)
    elif evt.state == 12: # Control
        checker = 2
        _iid = tab.identify_row(evt.y)
        focus_list.append(_iid)
    else: print("Not work")

    converter(focus_list, evt.widget)

def onselect(evt):
    global checker
    index_list.clear()
    focus_list.clear()
    checker = 0
    _iid = tab.identify_row(evt.y)
    focus_list.append(_iid)
    
    converter(focus_list, evt.widget)

def load_buttons():
    buttonpfad = Button(main, text = "Suchen", command = openfile, font = "Arial 9 bold", relief='flat')
    buttonload = Button(main, text = "Laden", command = import_to_db, font = "Arial 9 bold", relief='flat')
    buttonselect = Button(main, text = "Select", command = selected, font = "Arial 9 bold", relief='flat')
    buttondeselect = Button(main, text = "Deselect", command = deselected, font = "Arial 9 bold", relief='flat')
    buttontocsv = Button(main, text = "Send to CSV", command = send_to_csv, font = "Arial 9 bold", relief='flat')
    buttontops = Button(main, text = "Send to PS1", command = send_to_ps, font = "Arial 9 bold", relief='flat')

    buttonpfad.place(x = 20, y = 724, height= 40, width=80)
    buttonload.place(x = 120, y = 724, height= 40, width=80)
    buttonselect.place(x = 220, y = 724, height= 40, width=80)
    buttondeselect.place(x = 320, y = 724, height= 40, width=80)
    buttontocsv.place(x = 420, y = 724, height= 40, width=80)
    buttontops.place(x = 520, y = 724, height= 40, width=80)
    
def openfile():
    global filename
    filename = filedialog.askopenfilename(filetypes=(("CSV files","*.csv"),("All files","*.*")))
    return filename

def import_to_db():
    
    def send_data():
        connect_db()
        mycursor = myDBcon.cursor()
        zumDB = "USE interessenten_db;"
        mycursor.execute(zumDB)
        
        check_bereitschaft = mb.askyesno(title="Bereitschaft", message="Sind Sie sicher?", parent=main)
        
        if check_bereitschaft and (len(new_items) > 0):
            x = 0
            for x in range(len(new_items[0])):
                send_d = "INSERT INTO `interessenten` (`Standort`, `PLZ`, `Adresse_Standort`, `KFZ`, `Abteilung`, `Anrede`, `Vorname`, `Nachname`, `EMail`, `Benutzeranmeldename`) VALUES ('" + str(new_items[x][0]) + "','" + str(new_items[x][1]) + "','" + str(new_items[x][2]) + "','" + str(new_items[x][3]) + "','" + str(new_items[x][4]) + "','" + str(new_items[x][5]) + "','" + str(new_items[x][6]) + "','" + str(new_items[x][7]) + "','" + str(new_items[x][8]) + "','" + str(new_items[x][9]) + "');"
                try:
                    mycursor.execute(send_d)
                    myDBcon.commit()
                except:
                    myDBcon.rollback()
                x += 1
                
        tab_frame.forget()        
        tab.forget()
        vsb.forget()
        createtab()

    if filename:
        file = open(filename, "r", encoding='utf-8')
        csvdata = list(csv.reader(file, delimiter=","))
        file.close()

        new_items = []
        
        if len(csvdata[0]) == (len(all_data[0])-1):
            for data in csvdata:
                if data[0] == "Standort": continue
                else: 
                    new_items.append(data)
        else:
            mb.showerror(message = "Kann nicht vergleichen! Bitte andere Datei wählen.")
        
        send_data()

    else:
        mb.showerror(message = "Bitte die Datei auswählen.")
        
def selected():
    global checker
    global tab
    w = tab
    index_list.clear()
    focus_list.clear()
    for item_id in w.get_children():
        w.selection_add(item_id)
    focus_list.append(w.selection()[0])
    focus_list.append(w.selection()[-1])
    checker = 1
    converter(focus_list, tab)

def deselected():
    index_list.clear()
    focus_list.clear()
    for item_id in tab.get_children():
        tab.selection_remove(item_id)
    
def send_to_csv():
    try:
        file_path = str((filedialog.asksaveasfile(mode='w', filetypes=(("CSV files","*.csv"),("All files","*.*")))).name)
        with open(file_path, mode='a', newline='', encoding='utf-8') as file:
            csvwriter = csv.writer(file, dialect="excel", delimiter=",", quotechar='|')
            for row_id in tab.get_children():
                row_data = tab.item(row_id)["values"]
                csvwriter.writerow(row_data)
    except AttributeError:
        pass
        
def send_to_ps():
    try:
        file_path = str((filedialog.asksaveasfile(mode='w', filetypes=(("PS1 files","*.ps1"),("All files","*.*")))).name)
        with open(file_path, mode='a', newline='', encoding='utf-8') as file:
            file.write("Import-Module ActiveDirectory\n")
            file.write("\n")
            file.write("$Kennwort = ConvertTo-SecureString -String 'Passwort123' -AsPlainText -Force")
            file.write("\n")

            for each in tab.get_children():
                user = (tab.item(each)["values"])
                for x in index_list:
                    if user[0] == x:
                        file.write("\nif (-not (Get-ADUser -Filter {SamAccountName -eq '" + (user[7] + ' ' + user[8]) + "'})) {\nNew-ADUser -Name '" + (user[7] + ' ' + user[8]) + "' -GivenName '" + user[7] + "' -Surname '" + user[8] + "' -SamAccountName '" + user[10] + "' -UserPrincipalName '" + user[9] + "' -AccountPassword $Kennwort -Enabled $true \n} else {\n} \ncontinue \n")
    except AttributeError:
        pass

main = Tk()
main.title("Interessenten Liste")
main.geometry(f"1400x780")
main.config(background="#e3dcdc")
main.focus()

createtab()
load_buttons()

main.mainloop()