import os
import pathlib
import database_funtions as db_func
import pandas as pd
import create_space_functions as csf
import edit_space_functions as esf
import universal_functions as uf
import xlwings as xw

# Begins to make the space graph workbook
def main_graph_func(id_version, id_wif, conn, node_sheet):

    space_alloc_table, bldg_space_table, version_name, wif_name = db_func.get_db_data(conn, id_version, id_wif)

    # Set excel file name
    excel_file_name = f'{wif_name}w{id_wif}v{id_version}_{version_name}.xlsx'
    # Check if file exists
    folder = '.\graphs\\'

    # Create graph directory if it does not exist
    if not os.path.exists(folder):
        os.makedirs(folder)

    file_path = folder + excel_file_name
    file = pathlib.Path(file_path)

    # Check if node_sheet exists, if so get data
    if node_sheet is None:
        node_rollup_df = None
    
    else:
        node_rollup_df = node_sheet.range('A1').options(pd.DataFrame, index=False, expand='table').value
        node_rollup_df = uf.format_node_rollup(node_rollup_df)

    # Check if file exists
    if file.exists():
        print("File exist")

        active_xl_books = xw.books
        book_name_list = []
        for i in range(len(active_xl_books)):
            book_name_list.append(active_xl_books[i].name)
        # print(book_name_list)

        # Check if file is open
        if excel_file_name in book_name_list:
            wb = xw.Book(excel_file_name)
            print("used book name")
        else:
            # write df to excel sheet
            wb = xw.books.open(file_path)
            print("used file path")

        esf.edit_space_graphs(space_alloc_table, bldg_space_table, wb, node_rollup_df)

    else:
        print("File not exist")
        csf.create_space_graph_wb(file_path, space_alloc_table, bldg_space_table, node_rollup_df)


