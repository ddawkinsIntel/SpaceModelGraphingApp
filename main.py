import numpy.random.common
import numpy.random.bounded_integers
import numpy.random.entropy
import tkinter as tk
from tkinter import messagebox
import os
import database_funtions as db_func
import primary_functions as pf
import xlwings as xw

########################################################################################################################



# After values are verified, this starts the program
def fire_up(id_version, id_wif, s_trk_wb, s_trk_sheet):
    # Convert id_version and id_wif to string
    id_version = str(id_version)
    id_wif = str(id_wif)

    # Intiates connection to the sql server
    cnxn = db_func.initiate_conn()

    if s_trk_wb == "None":
        node_sheet = None
        pf.main_graph_func(id_version, id_wif, cnxn, node_sheet)

    else:
        # Get values from "wb_from" & set as capacity_sheet
        s_trk_wb = xw.books[s_trk_wb]
        node_sheet = s_trk_wb.sheets[s_trk_sheet]

        # Make the space graph workbook
        pf.main_graph_func(id_version, id_wif, cnxn, node_sheet)

    current_directory = os.getcwd()
    success_msg = "Graphs are located in the following folder.... \n \n" + current_directory
    messagebox.showinfo("Success!", success_msg)

# First function to be called gets and checks value user enters
def get_value():
    # Gets values from each field
    version_value = version_entry.get()
    wif_value = wif_entry.get()

    scr_trk_wb = scr_trk_workbook.get()
    # scr_trk_sheet = scr_trk_sheet_entry.get()
    scr_trk_sheet = "Node rollups"

    try:
        int(version_value) # checks if version_value is an int

        # Checks if wif_value NOT exists...if not set to 0...else check if its an int 
        if not wif_value:
            wif_value = 0
        else:
            int(wif_value)

        print("version value is: ", version_value)
        print("wif value is: ", wif_value)

        # fire_up(version_value, wif_value, scr_trk_wb, scr_trk_sheet)

    except ValueError:
        messagebox.showerror("Error", "Value Must Be a Number")
        root.update()

    except IndexError:
        messagebox.showerror("Error", "Choose a Valid Version Number")
        root.update()

    except Exception as e:
        print(e)
        messagebox.showerror("Error", "Unexpected Error")
        root.update()

    # # For testing
    # int(version_value)
    # print(version_value)
    # fire_up(version_value)
    fire_up(version_value, wif_value, scr_trk_wb, scr_trk_sheet)


# Terinates program when user closes the window
def on_closing():
    root.destroy()
    quit()


#######################################################################################################################
# Create the root
root = tk.Tk()

# Initiate versionValue and wifValue
versionValue = tk.StringVar()
wifValue = tk.StringVar()
scr_trk_workbook = tk.StringVar()
scr_trk_sheet = tk.StringVar()

# Window title
root.title('Space Model Graphing App')
root.geometry('550x300') # window size

book_name_list = ["None"]
for i in range(len(xw.books)):
    book_name_list.append(xw.books[i].name)

# Workbook dropdown menu options
wb_options = book_name_list

# Initial Workbook dropdown menu text
scr_trk_workbook.set(book_name_list[0])


# Creates version label and entry field 
version_label = tk.Label(root, text="Enter Version ID:").grid(row=1, column=1, sticky=tk.W, pady=15)
version_entry = tk.Entry(root, width=30, textvariable=versionValue)
version_entry.grid(row=1, column=2, sticky=tk.W, pady=15)

# Creates wif label and entry field 
wif_label = tk.Label(root, text="Enter WIF ID:").grid(row=2, column=1, sticky=tk.W, pady=15)
wif_entry = tk.Entry(root, width=30, textvariable=wifValue)
wif_entry.grid(row=2, column=2, sticky=tk.W, pady=15)

# Create "Cap Workbook" Dropdown menu
scr_trk_workbook_label = tk.Label(root, text="S.Tracker Workbook:").grid(row=3, column=1, sticky=tk.W, pady=15)
scr_trk_workbook_op = tk.OptionMenu( root , scr_trk_workbook , *wb_options )
scr_trk_workbook_op.grid(row=3, column=2, sticky=tk.W, pady=15)

# # Creates wif label and entry field 
# scr_trk_sheet_label = tk.Label(root, text="S.Tracker Sheet Name:").grid(row=4, column=1, sticky=tk.W, pady=15)
# scr_trk_sheet_entry = tk.Entry(root, width=30, textvariable=scr_trk_sheet)
# scr_trk_sheet_entry.grid(row=4, column=2, sticky=tk.W, pady=15)

# Creates "submit" button and calls get_value function when clicked
submit_btn = tk.Button(root, text="Submit", width=10, command=get_value).grid(row=5, column=1, columnspan=5, pady=25)

# Calls get_value() function when the user hits the "enter" key while in the version_entry or wif_entry field
version_entry.bind("<Return>", lambda e: get_value())
wif_entry.bind("<Return>", lambda e: get_value())
# scr_trk_sheet_entry.bind("<Return>", lambda e: get_value())

# Quits program on close
root.protocol("WM_DELETE_WINDOW", on_closing)

# Keeps root open, until user terminates it
root.mainloop()

version_id = int(versionValue.get())
wif_id = int(wifValue.get())

########################################################################################################################
