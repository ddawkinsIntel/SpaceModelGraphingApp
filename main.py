import numpy.random.common
import numpy.random.bounded_integers
import numpy.random.entropy
import tkinter as tk
from tkinter import messagebox
import os
import pathlib
import database_funtions as db_func
import graph_write_functions as gw_func

########################################################################################################################

# Begins to make the space graph workbook
def main_graph_func(id_version, id_wif, conn):

    space_alloc_table, bldg_space_table, version_name = db_func.get_db_data(conn, id_version, id_wif)

    # Set excel file name
    excel_file_name = r"pySpaceGraph v" + id_version + "w" + id_wif + "_" +version_name + ".xlsx"

    # Check if file exists
    folder = r".\graphs\\"

    # Create graph directory if it does not exist
    if not os.path.exists(folder):
        os.makedirs(folder)

    file_path = folder + excel_file_name
    file = pathlib.Path(file_path)
    if file.exists():
        print("File exist")
        gw_func.edit_space_graphs(space_alloc_table, bldg_space_table, file_path)

    else:
        print("File not exist")
        gw_func.create_space_graph_wb(file_path, space_alloc_table, bldg_space_table)

# After values are verified, this fires up the program
def fire_up(id_version, id_wif):
    # Convert id_version and id_wif to string
    id_version = str(id_version)
    id_wif = str(id_wif)

    # Intiates connection to the sql server
    cnxn = db_func.initiate_conn()

    # Make the space graph workbook
    main_graph_func(id_version, id_wif, cnxn)

    current_directory = os.getcwd()
    success_msg = "Graphs are located in the following folder.... \n \n" + current_directory
    messagebox.showinfo("Success!", success_msg)

# Gets and checks value user enters
def get_value():
    # Gets values from each field
    version_value = version_entry.get()
    wif_value = wif_entry.get()

    try:
        int(version_value) # checks if version_value is an int

        # Checks if wif_value NOT exists...if not set to 0...else check if its an int 
        if not wif_value:
            wif_value = 0
        else:
            int(wif_value)

        print("version value is: ", version_value)
        print("wif value is: ", wif_value)

        fire_up(version_value, wif_value)

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

# Window title
root.title('Space Model Graphs')
root.geometry('325x165') # window size

# Creates version label and entry field 
version_label = tk.Label(root, text="Enter Version ID:").grid(row=1, column=1, sticky=tk.W, pady=15)
version_entry = tk.Entry(root, width=30, textvariable=versionValue)
version_entry.grid(row=1, column=2, sticky=tk.W, pady=15)

# Creates wif label and entry field 
wif_label = tk.Label(root, text="Enter WIF ID:").grid(row=2, column=1, sticky=tk.W, pady=15)
wif_entry = tk.Entry(root, width=30, textvariable=wifValue)
wif_entry.grid(row=2, column=2, sticky=tk.W, pady=15)

# Creates "submit" button and calls get_value function when clicked
submit_btn = tk.Button(root, text="Submit", width=10, command=get_value).grid(row=3, column=1, columnspan=5, pady=25)

# Calls get_value() function when the user hits the "enter" key while in the version_entry or wif_entry field
version_entry.bind("<Return>", lambda e: get_value())
wif_entry.bind("<Return>", lambda e: get_value())

# Quits program on close
root.protocol("WM_DELETE_WINDOW", on_closing)

# Keeps root open, until user terminates it
root.mainloop()

version_id = int(versionValue.get())
wif_id = int(wifValue.get())

########################################################################################################################
