import tkinter as tk
from tkinter import ttk
import numpy as np
import pandas as pd
from ics import Calendar, Event

def format_time(DateIn, TimeIn): #Helper function for calendar
    #format date to YYYYMMDD or YYYY-MM-DD or YYYY/M/D
    #excel default date format is DDMMYYYY (For UK, make sure to change in vba SaveAs by using Local:=True)
    Date_1 = DateIn.split("/")
    Date_new = str(Date_1[2] + "-" + Date_1[1] + "-" + Date_1[0])
    # Time_1 = TimeIn.split(" ") #Need to deal with the stupid AM PM time format
    #
    # if Time_1[1] == 'AM':
    #     Time_new = Time_1[0]
    # else:
    #     Time_2 = Time_1[0].split(":")
    #     Time_2[0] = int(Time_2[0]) + 12
    #     Time_new = str(str(Time_2[0]) + ":" + Time_2[1] + ":" + Time_2[2])
    #     #To do: Account for timezone? Lmao

    #formatted_date_time = Date_new + " " + Time_new
    formatted_date_time = Date_new + " " + TimeIn
    return formatted_date_time

#create main window
root= tk.Tk()
root.title('Generate ICS Files from Deadlines')
root.geometry('600x200')
root.columnconfigure(0, weight=4)
root.columnconfigure(1, weight=1)


#Create Instructions Pane
instruction_frame = ttk.Frame(root)
ttk.Label(
    instruction_frame,
    text='This tool is to create a calendar file (.ics) for your module'
         '\n You can generate calendars individually, or for all staff'
         '\n Make sure you have the DataForStaff csv ',
    anchor = "w"
    ).grid(column=0, row=0, sticky=tk.W, pady=5)


#Create Left Hand Input Pane
input_frame = ttk.Frame(root)

ttk.Label(input_frame, text='Choose Module: ').grid(column=0, row=0, sticky=tk.W, padx=20)


#Frame within Input Pane that holds file input fields
filename_input_frame = ttk.Frame(root)

dataforstaff_file = tk.StringVar()
filename1 = "DataForStaff.csv"
dataforstaff_file.set(filename1) #Default value to initialize with
ttk.Label(filename_input_frame, text='Data for Staff filename: ').grid(column=0, row=0, sticky=tk.W)
module_data_file = tk.Entry(filename_input_frame, textvariable=dataforstaff_file)
module_data_file.grid(column=1, row=0, sticky=tk.W)

commondeadlines_file = tk.StringVar()
commondeadlines_file.set("CommonDeadlines.csv") #Default value to initialize with
ttk.Label(filename_input_frame, text='Common Deadlines filename: ').grid(column=0, row=1, sticky=tk.W)
common_data_file = tk.Entry(filename_input_frame, textvariable=commondeadlines_file)
common_data_file.grid(column=1, row=1, sticky=tk.W)

include_common = tk.BooleanVar()
common_button = tk.Checkbutton(filename_input_frame, text='Include Common Deadlines for Staff',
                               variable=include_common, onvalue=True, offvalue=False)
common_button.grid(column=0, row=2, sticky=tk.W)


instruction_frame.grid(column=0, row=0)
filename_input_frame.grid(column=0,row=1)
input_frame.grid(column=0, row=2)
#input_frame.grid(column=0, row=2, sticky=tk.W, padx=25)

#canvas1 = tk.Canvas(root, width = 300, height = 300)
#canvas1.pack()

#Create Right Hand Button Pane
button_frame = ttk.Frame(root)
button_frame.grid(column=1, row=0)

#Create Notif Frame
notif_frame = ttk.Frame(root)
bottom_notif = ttk.Label(notif_frame, text='Ready', anchor="w")
bottom_notif.grid(column=0, row=0, sticky=tk.W)

notif_frame.grid(column=0, row=3, sticky=tk.W, padx=70)

#Initial Data Load
df = pd.read_csv(filename1)
print(dataforstaff_file)
mods_list = df["Module Code "].drop_duplicates().tolist()
mods_list_sorted = sorted(mods_list, reverse=False)

mod_menu = tk.StringVar()
mod_menu.set(mods_list[3])

def generate_ICS():
    option = mod_menu.get()
    print(option)
    df_selected = df.loc[df['Module Code '] == option]
    df_selected.reset_index()
    #print(df_selected)

    c = Calendar()

    for i in range(df_selected.shape[0]):

        profname = str(df_selected.iloc[i,6])
        name = str(df_selected.iloc[i,3] + " " + df_selected.iloc[i,7])
        DeadlineTime = df_selected.iloc[i,11]
        DeadlineDate = df_selected.iloc[i,12]
        DeadlineDateTime = format_time(DeadlineDate,DeadlineTime)
        print(name)
        print(DeadlineTime)
        print(DeadlineDateTime)

        e = Event() #Add an event
        e.name = name
        e.begin = DeadlineDateTime
        c.events.add(e)

    #print(c.events)
    ics_filename = str(option + "_" + profname + "_calendar.ics")
    with open(ics_filename, 'w') as f:
        f.writelines(c.serialize_iter())

    #label1 = tk.Label(button_frame, text= str(option + ' icsfile generated!'), fg='black', font=('helvetica', 10))
    label1 = tk.Label(notif_frame, text= str(option + ' icsfile generated!'), fg='black', font=('helvetica', 10))
    label1.grid(column=0, row=0, sticky=tk.W)


    #canvas1.create_window(150, 200, window=label1)

def option_changed(option):
    option = mod_menu.get()
    label2 = tk.Label(input_frame, text=str("You chose: " + option), fg='blue', font=('helvetica', 10))
    label2.grid(column=0, row=3)
    #bottom_notif = tk.Label(notif_frame, text=str('You chose: ' + option), fg="blue")
    #bottom_notif.grid(column=0, row=0, sticky=tk.W)
    #canvas1.create_window(150, 200, window=label2)

def generate_ICS_all():
    all_mods = df['Module Code '].unique()

    # option = mod_menu.get()
    # print(option)
    # df_selected = df.loc[df['Module Code '] == option]
    # df_selected.reset_index()
    # #print(df_selected)

    for module in all_mods:
        c = Calendar()
        print("Generating ICS for ",module)
        df_selected = df.loc[df['Module Code '] == module]
        df_selected.reset_index()
    #
        for i in range(df_selected.shape[0]):
            profname = str(df_selected.iloc[i,6])
            name = str(df_selected.iloc[i,3] + " " + df_selected.iloc[i,7])
            DeadlineTime = df_selected.iloc[i,11]
            DeadlineDate = df_selected.iloc[i,12]
            DeadlineDateTime = format_time(DeadlineDate,DeadlineTime)

            e = Event() #Add an event
            e.name = name
            e.begin = DeadlineDateTime
            c.events.add(e)

        ics_filename = str(module + "_" + profname + "_calendar.ics")
        with open(ics_filename, 'w') as f:
            f.writelines(c.serialize_iter())
        print("Done ICS for ",module)
    #
    # #label1 = tk.Label(button_frame, text= str(option + ' icsfile generated!'), fg='black', font=('helvetica', 10))
    # label1 = tk.Label(notif_frame, text= str(option + ' icsfile generated!'), fg='black', font=('helvetica', 10))
    # label1.grid(column=0, row=0, sticky=tk.W)


#Adding buttions to the Right Hand Pane
button1 = ttk.Button(button_frame, text='Generate ICS File', command=generate_ICS)
button1.grid(column=0, row=0)
button2 = ttk.Button(button_frame, text='Generate for All Mods', command=generate_ICS_all)
button2.grid(column=0, row=1)



#To be replaced with Listbox? or restrict to Mecheng mods only
choose_mod_menu = tk.OptionMenu(
    input_frame,
    mod_menu,
    *mods_list_sorted,
    command=option_changed
    )

#choose_mod_menu.pack(expand=True)
choose_mod_menu.grid(column=1, row=0)




root.mainloop()