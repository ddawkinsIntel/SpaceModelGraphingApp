import os
import pathlib
import database_funtions as db_func
import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'

import create_space_functions as csf
import edit_space_functions as esf
import universal_functions as uf
import xlwings as xw

# Begins to make the space graph workbook
def main_graph_func(id_version, id_wif, conn, node_sheet):

    # Get data from database
    space_alloc_table, bldg_space_table, version_name, wif_name = db_func.get_db_data(conn, id_version, id_wif)

    # Set excel file name
    excel_file_name = f'{wif_name}w{id_wif}v{id_version}_{version_name}.xlsx'

    # folder location for the graph output
    folder = '.\graphs\\'

    # Create the 'graphs' folder, if it does not exist
    if not os.path.exists(folder):
        os.makedirs(folder)

    # Creating the 'graphs' folders file path
    file_path = folder + excel_file_name
    file = pathlib.Path(file_path) # creates an concrete path for the platform the code is running on

    # Check if node_sheet exists, if so get data
    if node_sheet is None:
        node_rollup_df = None
    
    else:
        node_rollup_df = node_sheet.range('A1').options(pd.DataFrame, index=False, expand='table').value
        node_rollup_df = uf.format_node_rollup(node_rollup_df)

    # Check if file exists and edits the file, else creates new file
    if file.exists():
        print("File exists, editing the existing file... \n")

        # Intiate an 'xw' object that contains all the active(open) excel workbooks
        try:
            book_name_list = []
            active_xl_books = xw.books

        except AttributeError:
            print('No active excel workbooks....')
            print('....opening Excel \n')
            xw.Book()

        # Loop through active excel workbooks and adds there name to the book_name_list
        for i in range(len(active_xl_books)):
            book_name_list.append(active_xl_books[i].name)

        # Check if the intended file is active by checking if it's in the book_name_list 
        if excel_file_name in book_name_list:
            wb = xw.Book(excel_file_name) # stores value of the active sheet, since it's active file_path not needed

        else:
            # write df to excel sheet
            wb = xw.books.open(file_path) # stores path to the intended file

        esf.edit_space_graphs(space_alloc_table, bldg_space_table, wb, node_rollup_df)

    else:
        print("File does not exist, createing a new file... \n")
        csf.create_space_graph_wb(file_path, space_alloc_table, bldg_space_table, node_rollup_df)


