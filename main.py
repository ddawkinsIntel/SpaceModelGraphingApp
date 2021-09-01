import pyodbc
# import numpy.random.common
# import numpy.random.bounded_integers
# import numpy.random.entropy
import pandas as pd
import tkinter as tk
from tkinter import messagebox
import os
import xlwings as xw
import pathlib

# Select what site to iterate through
site_list = ['AZ', 'IR', 'IS', 'LTD', 'HVM1', 'HVM2', 'HVM3', 'HVM4', 'HVM5', 'NM']

# Process Color Dictionary
proc_color_dict = {1222: '#ED7D31', 1270: '#A5A5A5', 1272: '#FFC000', 1274: '#5B9BD5', 1276: 'purple', 1277: '#B3A2C7',
                   1278: '#663300', 1280: '#71FFBF'}


# Chart stuff
def add_area_chart_series(chart, chart_color, sheet, row, ncols, start_row):
    chart.add_series({
        'name': [sheet, row + 1 + start_row, 0],
        'categories': [sheet, start_row, 1, start_row, ncols],
        'values': [sheet, row + 1 + start_row, 1, row + 1 + start_row, ncols],
        'fill': {'color': chart_color}
    })


def add_line_chart_series(chart, chart_color, sheet, row, ncols, start_row):
    chart.add_series({
        'name': [sheet, row + 1 + start_row, 0],
        'categories': [sheet, start_row, 1, start_row, ncols],
        'values': [sheet, row + 1 + start_row, 1, row + 1 + start_row, ncols],
        'line': {'color': chart_color}
    })


# Sorts df so excel graphs always show install and demo at the top
def sort_df_rows(df):
    # Get only production rows
    production_df = df[df["CapacityType"].str.contains("Production")]

    # Get only install demo and mcrsf rows
    install_demo_mcrsf_df = df[~(df["CapacityType"].str.contains("Production") | df["CapacityType"].str.contains("WSVolume"))]

    # Get only wafer volume rows
    wafer_df = df[df["CapacityType"].str.contains("WSVolume")]

    # concatenating df1 and df2 along rows
    sorted_df = pd.concat([production_df, install_demo_mcrsf_df, wafer_df], axis=0)

    sorted_df = sorted_df.fillna(0)

    return sorted_df


def make_chart(df, site_val, pd_writer, start_row=0, chart_cell="G12"):

    # Use site value as sheet name
    sheet_name = site_val

    # Sort df rows for excel graph aesthetic purposes
    sorted_df = sort_df_rows(df)

    # write df to excel sheet
    sorted_df.to_excel(pd_writer, sheet_name=sheet_name, startrow=start_row, index=False)

    # Access the XlsxWriter workbook and worksheet objects from the dataframe.
    workbook = pd_writer.book
    worksheet = pd_writer.sheets[sheet_name]

    # Create a chart object.
    area_chart = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})

    # Get shape without including wafer volume rows
    non_wafer_df = df[~df["CapacityType"].str.contains("WSVolume")]

    # Number of columns
    ncols = non_wafer_df.shape[1]-1

    # Configure the series of the chart from the dataframe data.
    max_row = len(non_wafer_df)
    for i in range(max_row):

        row = i
        colname = sorted_df.iloc[row, 0]

        if colname == "MCRSF":
            # Add line to chart
            line_chart1 = workbook.add_chart({'type': 'line'})

            color = 'black'

            # Line chart series
            add_line_chart_series(line_chart1, color, sheet_name, row, ncols, start_row)

            # Combine Charts
            area_chart.combine(line_chart1)
        elif colname == "Demo":
            color = '#C00000'

            # Area chart series
            add_area_chart_series(area_chart, color, sheet_name, row, ncols, start_row)

        elif colname == "Install":
            color = "#00B050"

            # Area chart series
            add_area_chart_series(area_chart, color, sheet_name, row, ncols, start_row)

        else:
            process = int(colname[:4])
            color = proc_color_dict.get(process, "#fc03df")

            # Area chart series
            add_area_chart_series(area_chart, color, sheet_name, row, ncols, start_row)


    # Place legend at bottom of chart
    area_chart.set_legend({'position': 'bottom'})

    # Configure the chart axes.
    area_chart.set_x_axis({'name': 'Month'})
    area_chart.set_y_axis({'name': 'MCRSF', 'major_gridlines': {'visible': False}})

    # Set chart size
    area_chart.set_size({'width': 1100, 'height': 500})

    # Insert the chart into the worksheet @ cell G12
    worksheet.insert_chart(chart_cell, area_chart)


def make_alloc_chart(df, site_val, pd_writer):
    # Number of columns
    ncols = df.shape[1]-1

    # Use site value as sheet name
    sheet_name = site_val

    # write df to excel sheet
    df.to_excel(pd_writer, sheet_name=sheet_name, index=False)

    # Access the XlsxWriter workbook and worksheet objects from the dataframe.
    workbook = pd_writer.book
    worksheet = pd_writer.sheets[sheet_name]

    # # Create a chart object.
    # line_chart = workbook.add_chart({'type': 'line'})
    #
    # # Process Color Dictionary
    # proc_color_dict = {1222: '#ED7D31', 1270: '#A5A5A5', 1272: '#FFC000', 1274: '#5B9BD5', 1276: 'purple', 1277: '#B3A2C7', 1278: '#663300', 1280: '#71FFBF'}
    #
    # line_chart = workbook.add_chart({'type': 'line'})
    #
    # # Configure the series of the chart from the dataframe data.
    # max_row = len(df)
    # for i in range(max_row):
    #     row = i
    #     colname = df.iloc[row, 0]
    #
    #     color = 'black'
    #
    #     # Line chart series
    #     line_chart.add_series({
    #         'name': [sheet_name, row + 1, 1, row + 1, 1],
    #         'categories': [sheet_name, 0, row + 2, 0, ncols],
    #         'values': [sheet_name, row + 1, 2, row + 1, ncols],
    #         'line': {'color': color}
    #     })
    #
    #
    #
    # # Place legend at bottom of chart
    # line_chart.set_legend({'position': 'bottom'})
    #
    # # Configure the chart axes.
    # line_chart.set_x_axis({'name': 'Month'})
    # line_chart.set_y_axis({'name': 'WSPW', 'major_gridlines': {'visible': False}})
    #
    # # Set chart size
    # line_chart.set_size({'width': 1100, 'height': 500})
    #
    # # Insert the chart into the worksheet @ cell G12
    # worksheet.insert_chart('G12', line_chart)


def xw_write(df, file_name, sheet_name, df_cell='A1'):
    # write df to excel sheet
    wb = xw.Book(file_name)
    sorted_df = sort_df_rows(df)
    sheet = wb.sheets[sheet_name]
    sheet.range(df_cell).options(index=False).value = sorted_df


########################################################################################################################
# Monthly functions below
# Makes, formats and returns tmgsp monthly df
def make_tmgsp_mntly_df(tmgsp_list):
    # Concatenate list of dfs
    tmgsp_df = pd.concat(tmgsp_list)
    tmgsp_df = tmgsp_df.groupby(['CapacityType']).sum().reset_index()
    # tmgsp_df = tmgsp_df.append(tmgsp_df.sum(numeric_only=True).rename('Total')).reset_index()
    return tmgsp_df


# Xlwings makes tmgsp monthly graph - nothing is returned
def make_xw_tmgsp_mntly_table(tmgsp_list, file_name):
    # Makes and formats tmgsp monthly df
    tmgsp_df = make_tmgsp_mntly_df(tmgsp_list)

    sheet_name = "TMGSP"
    xw_write(tmgsp_df, file_name, sheet_name)


# Pandas makes tmgsp monthly graph - nothing is returned
def make_pd_tmgsp_mntly_graphs(tmgsp_list, writer):
    # Makes and formats tmgsp monthly df
    tmgsp_df = make_tmgsp_mntly_df(tmgsp_list)
    # Makes and saves chart
    make_chart(tmgsp_df, "TMGSP", writer)


########################################################################################################################
# Quarterly Functions below
# Makes, formats and returns tmgsp qtrly df
def make_site_qtrly_df(tmgsp_list_qtrly):
    # Concatenate list of dfs, fillna with 0 and groupby
    tmgsp_df_qtrly_long = tmgsp_list_qtrly.fillna(0)
    tmgsp_df_qtrly_long = tmgsp_df_qtrly_long.groupby(['CapacityType']).sum().reset_index()

    # converting the integers to datetime format
    tmgsp_df_qtrly_long['Date'] = pd.to_datetime(tmgsp_df_qtrly_long['CapacityType'], format='%Y%m')
    tmgsp_df_qtrly_long['qtr'] = pd.PeriodIndex(tmgsp_df_qtrly_long['Date'], freq='Q')
    tmgsp_df_qtrly_long['qtr'] = tmgsp_df_qtrly_long['qtr'].astype(str)

    return tmgsp_df_qtrly_long


# Cleans and returns tmgsp qtrly df
def clean_site_qtrly_df(tmgsp_df_qtrly_long):
    # Drop un-wanted columns
    tmgsp_df_qtrly_long.drop(['CapacityType'], axis='columns', inplace=True)
    tmgsp_df_qtrly_long.drop(['Date'], axis='columns', inplace=True)

    # Group by and get means for quarters
    tmgsp_df_qtrly_long = tmgsp_df_qtrly_long.groupby(['qtr']).mean()

    # Transpose "full" df then reset index
    tmgsp_df_qtrly = tmgsp_df_qtrly_long.T.reset_index()

    # changing column name
    tmgsp_df_qtrly.rename({'index': 'CapacityType'}, axis=1, inplace=True)

    return tmgsp_df_qtrly


# Makes, formats and returns tmgsp qtrly df
def make_tmgsp_qtrly_df(tmgsp_list_qtrly):
    # Concatenate list of dfs, fillna with 0 and groupby
    tmgsp_df_qtrly_long = pd.concat(tmgsp_list_qtrly)
    tmgsp_df_qtrly_long = tmgsp_df_qtrly_long.fillna(0)
    tmgsp_df_qtrly_long = tmgsp_df_qtrly_long.groupby(['CapacityType']).sum().reset_index()

    # converting the integers to datetime format
    tmgsp_df_qtrly_long['Date'] = pd.to_datetime(tmgsp_df_qtrly_long['CapacityType'], format='%Y%m')
    tmgsp_df_qtrly_long['qtr'] = pd.PeriodIndex(tmgsp_df_qtrly_long['Date'], freq='Q')
    tmgsp_df_qtrly_long['qtr'] = tmgsp_df_qtrly_long['qtr'].astype(str)

    return tmgsp_df_qtrly_long


# Cleans and returns tmgsp qtrly df
def clean_tmgsp_qtrly_df(tmgsp_df_qtrly_long):
    # Drop un-wanted columns
    tmgsp_df_qtrly_long.drop(['CapacityType'], axis='columns', inplace=True)
    tmgsp_df_qtrly_long.drop(['Date'], axis='columns', inplace=True)

    # Group by and get means for quarters
    tmgsp_df_qtrly_long = tmgsp_df_qtrly_long.groupby(['qtr']).mean()

    # Transpose "full" df then reset index
    tmgsp_df_qtrly = tmgsp_df_qtrly_long.T.reset_index()

    # changing column name
    tmgsp_df_qtrly.rename({'index': 'CapacityType'}, axis=1, inplace=True)

    return tmgsp_df_qtrly


# Xlwings makes tmgsp qtrly graph - nothing is returned
def make_xw_tmgsp_qtrly_graph(tmgsp_list_qtrly, file_name):
    # Make and clean qtrly df
    tmgsp_df_qtrly_long = make_tmgsp_qtrly_df(tmgsp_list_qtrly)
    tmgsp_df_qtrly = clean_tmgsp_qtrly_df(tmgsp_df_qtrly_long)

    sheet_name = "TMGSP_QTR"

    xw_write(tmgsp_df_qtrly, file_name, sheet_name)


# Pandas makes tmgsp qtrly graph - nothing is returned
def make_pd_tmgsp_qtrly_graph(tmgsp_list_qtrly, writer):
    # Make and clean qtrly df
    tmgsp_df_qtrly_long = make_tmgsp_qtrly_df(tmgsp_list_qtrly)
    tmgsp_df_qtrly = clean_tmgsp_qtrly_df(tmgsp_df_qtrly_long)

    # Makes and saves chart
    make_chart(tmgsp_df_qtrly, "TMGSP_QTR", writer)


# Xlwings makes tmgsp qtrly graph - nothing is returned
def make_xw_site_qtrly_graph(site_qtrly, site, file_name):
    # Make and clean qtrly df
    site_qtrly_df_long = make_site_qtrly_df(site_qtrly)
    site_qtrly_df = clean_tmgsp_qtrly_df(site_qtrly_df_long)

    xw_write(site_qtrly_df, file_name, site, df_cell='A66')


# Pandas makes site qtrly graph - nothing is returned
def make_pd_site_qtrly_graph(site_qtrly, site, writer):
    # Make and clean qtrly df
    site_qtrly_df = make_site_qtrly_df(site_qtrly)
    site_qtrly_df = clean_tmgsp_qtrly_df(site_qtrly_df)

    # Makes and saves chart
    make_chart(site_qtrly_df, site, writer, 65, 'G40')
########################################################################################################################


# Pandas creates tmgsp monthly and quarterly graphs
def tmgsp_graphs(tmgsp_list, tmgsp_list_qtrly, writer):
    make_pd_tmgsp_mntly_graphs(tmgsp_list, writer)
    make_pd_tmgsp_qtrly_graph(tmgsp_list_qtrly, writer)


########################################################################################################################
# DB functions below
# Initiates DB connection - returns a "conn"
def initiate_conn():

    # Initiate Connection to the db
    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=devpofco100.amr.corp.intel.com, 3180;'
                          'Database=SpaceModel_PROD_Bridge;'
                          'UID=AMR\ddawkins;'
                          'Trusted_Connection=yes;'
                          )
    return conn


# Returns version info table
def get_version_info(conn):
    sql = """\
    EXEC sml.pr_Version_Get
    """
    # test = cursor.execute(sql, params)
    table_version = pd.read_sql(sql, conn)

    return table_version


# From version table, returns the version name
def get_version_name(conn, id_version):
    # Get version info
    version_info = get_version_info(conn)

    # Get version name from version info, using the versionId
    version = version_info[version_info["VersionID"] == int(id_version)]

    version_name = version["VersionName"].values
    version_name = version_name[0]

    return version_name


# HF - get_building_space: Returns building space table
def read_bldg_space(conn, id_version):
    sql = """\
    EXEC sml.pr_BuildingSpace_Get @VersionID = ?
    """
    # test = cursor.execute(sql, params)
    table_building_space = pd.read_sql(sql, conn, params=[id_version])

    return table_building_space


# HF - get_building_space: Modifies and returns building table_building_space
def space_modify_df(table_building_space):

    # Map SiteID to Site
    table_building_space['SiteID'] = table_building_space['SiteID'].map({2506: 'AZ', 2507: 'IR', 2508: 'IS'})

    # Transpose df
    table_building_space = table_building_space.T

    new_header = table_building_space.iloc[0]  # grab the first row for the header
    table_building_space = table_building_space[3:]  # take the data less the header row
    table_building_space.columns = new_header  # set the header row as the df header

    # Reset index
    table_building_space.reset_index(level=0, inplace=True)

    # changing column name
    table_building_space.rename({'index': 'Month'}, axis=1, inplace=True)

    # Cast Month column values as int
    table_building_space["Month"] = table_building_space["Month"].astype(int)

    return table_building_space


# Gets, formats and returns building space table
def get_building_space(conn, id_version):

    table_building_space = read_bldg_space(conn, id_version)

    table_building_space = space_modify_df(table_building_space)

    return table_building_space


# Returns space allocation table
def get_rpt_solver(conn, id_version):
    sql = """\
    EXEC sml.pr_Rpt_SolverSpaceAlloc @VersionID = ?
    """
    # test = cursor.execute(sql, params)
    table_space_alloc = pd.read_sql(sql, conn, params=[id_version])

    return table_space_alloc


# Gets and Returns all needed DB data
def get_db_data(conn, id_version):

    # Get space alloc table
    space_alloc_table = get_rpt_solver(conn, id_version)

    # Get building space table
    bldg_space_table = get_building_space(conn, id_version)

    # Get the versions name(i.e. 2021Q3LRP Deramp)
    version_name = get_version_name(conn, id_version)

    return space_alloc_table, bldg_space_table, version_name
########################################################################################################################


########################################################################################################################
# create_full_df and helper functions below
# HF proc_capacity_col_piv: Pivots and returns capacity dfs
def pivot_df(df, capacity_type):
    # Pivot using the capacity type columns as the values
    df = df.pivot_table(index=["Month"], columns="Process", values=capacity_type)

    return df


# HF make_cap_type_dfs: Formats and returns capacity dfs
def proc_capacity_col_piv(df_type, capacity_type, col_list="foo"):
    # If col_list param is passed it is used
    if col_list != "foo":
        df_type[capacity_type] = df_type[col_list].sum(axis=1)
    else:
        df_type[capacity_type] = df_type[capacity_type].astype(int)

    if capacity_type == "Production":
        # Add the capacity type to the end of the process
        df_type["Process"] = df_type["Process"] + " " + capacity_type

        # Pivot df
        df_type = pivot_df(df_type, capacity_type)

    elif capacity_type == "WSVolume":
        # Add the capacity type to the end of the process
        df_type["Process"] = df_type["Process"] + " " + capacity_type

        # Pivot df
        df_type = pivot_df(df_type, capacity_type)

    else:
        df_type = df_type[["Month", capacity_type]]
        df_type = df_type.groupby(['Month']).sum()

    return df_type


# Creates and returns individual capacity type dfs
def make_cap_type_dfs(df_space, site):

    # Filter for site
    df_space = df_space[df_space.Site == site]

    prod_df = df_space.copy()
    install_df = df_space.copy()
    demo_df = df_space.copy()
    df_wafer = df_space.copy()

    prod_col_list = ['FEProduction', 'BEProduction', 'FEFabTPT', 'BEFabTPT']
    install_col_list = ['FEInstall', 'BEInstall']

    df_prod = proc_capacity_col_piv(prod_df, "Production", prod_col_list)
    df_install = proc_capacity_col_piv(install_df, "Install", install_col_list)
    df_demo = proc_capacity_col_piv(demo_df, "Demo")
    df_wafer = proc_capacity_col_piv(df_wafer, "WSVolume")

    return df_prod, df_install, df_demo, df_wafer


# Creates returns full_df by merging dataframes
def make_full_df(df_prod, df_install, df_demo, df_wafer, df_bldg, site):
    df_full = df_prod.merge(df_install, how="outer", on="Month")
    df_full = df_full.merge(df_demo, how="outer", on="Month")
    df_full = df_full.merge(df_bldg, how="outer", on="Month")
    df_full = df_full.merge(df_wafer, how="outer", on="Month")

    # changing column name
    df_full.rename({'Month': 'CapacityType'}, axis=1, inplace=True)

    df_full = df_full.rename(columns={site: 'MCRSF'})  # combine with functoin above

    return df_full


# Transpose, modfiy and return full df
def transpose_and_modify_df(full):
    # Transpose "full" df then reset index
    full = full.T.reset_index()

    # Use the row values where the column labeled "index" equals "Month" as column headers
    header_df = full.loc[full['index'] == "CapacityType"]  # selects the "Month" row
    header_list = header_df.values.tolist()  # Converts df to a list
    full.columns = header_list[0]

    # Exclude column header from the data frame
    full = full.loc[full["CapacityType"] != "CapacityType"]

    return full


# Reduces number of columns and returns bldg_site_df
def bldg_space_df_reducer(bldg_space_table, site):
    try:
        # Filter for site
        bldg_site_df = bldg_space_table[["Month", site]]
    except KeyError:
        bldg_site_df = bldg_space_table[["Month"]]

    return bldg_site_df


# Where the formating and stitching fun begins
def create_full_df(space_alloc_table, bldg_space_table, site, tmgsp_list, tmgsp_list_qtrly):
    bldg_site_df = bldg_space_df_reducer(bldg_space_table, site)

    # Makes a prod, install and demo df from the space df
    prod_df, install_df, demo_df, wafer_df = make_cap_type_dfs(space_alloc_table, site)

    # Merges all data frames
    full_long = make_full_df(prod_df, install_df, demo_df, wafer_df, bldg_site_df, site)

    # Append the full df to the tmgsp_list
    tmgsp_list_qtrly.append(full_long)

    full_wide = transpose_and_modify_df(full_long)

    tmgsp_list.append(full_wide)

    return full_wide, full_long
########################################################################################################################


########################################################################################################################
# When the space graph workbook already exists....
# HF - edit_space_graphs: Use XLWings to write full df to excel sheet
def xw_write_full(full, site, file_name):
    # write df to excel sheet
    wb = xw.Book(file_name)

    sheet = wb.sheets[site]
    sheet.range('A1').options(index=False).value = full


# Called if space graph workbook does exist
def edit_space_graphs(space_alloc_table, bldg_space_table, file_name):

    tmgsp_list_qtrly = []
    tmgsp_list = []
    for site in site_list:

        monthly_df, qtrly_df = create_full_df(space_alloc_table, bldg_space_table, site, tmgsp_list, tmgsp_list_qtrly)
        xw_write_full(monthly_df, site, file_name)
        make_xw_site_qtrly_graph(qtrly_df, site, file_name)


    make_xw_tmgsp_mntly_table(tmgsp_list, file_name)
    make_xw_tmgsp_qtrly_graph(tmgsp_list_qtrly, file_name)
########################################################################################################################


########################################################################################################################
# When the space graph workbook does not exist....
# HF - create_space_graph_wb: Creates space graphs
def create_space_graphs(writer, space_alloc_table, bldg_space_table):

    tmgsp_list_qtrly = []
    tmgsp_list = []
    ####################################################################################################################
    for site in site_list:
        full_wide, full_long = create_full_df(space_alloc_table, bldg_space_table, site, tmgsp_list, tmgsp_list_qtrly)

        # Makes and saves chart
        make_chart(full_wide, site, writer)
        make_pd_site_qtrly_graph(full_long, site, writer)

    tmgsp_graphs(tmgsp_list, tmgsp_list_qtrly, writer)


# Called if space graph workbook does not exist
def create_space_graph_wb(excel_file_name, space_alloc_table, bldg_space_table):
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(excel_file_name, engine='xlsxwriter')
    create_space_graphs(writer, space_alloc_table, bldg_space_table)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
########################################################################################################################


# Begins to make the space graph workbook
def main_graph_func(id_version, conn):

    space_alloc_table, bldg_space_table, version_name = get_db_data(conn, id_version)

    # Set excel file name
    excel_file_name = r"pySpaceGraph v" + id_version + "_" + version_name + ".xlsx"

    # Check if file exists
    file = pathlib.Path(excel_file_name)
    if file.exists():
        print("File exist")
        edit_space_graphs(space_alloc_table, bldg_space_table, excel_file_name)

    else:
        print("File not exist")
        create_space_graph_wb(excel_file_name, space_alloc_table, bldg_space_table)


def fire_up(id_version):

    id_version = str(id_version)

    # Intiates connection to the sql server
    cnxn = initiate_conn()

    # Make the space graph workbook
    main_graph_func(id_version, cnxn)

    current_directory = os.getcwd()
    success_msg = "Graphs are located in the following folder.... \n \n" + current_directory
    messagebox.showinfo("Success!", success_msg)


def get_value():

    version_value = version_entry.get()

    try:
        int(version_value)
        print(version_value)
        fire_up(version_value)

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


def on_closing():
    root.destroy()
    quit()


#######################################################################################################################
root = tk.Tk()


versionValue = tk.StringVar()

root.title('Space Model Graphs')
root.geometry('325x150')

version_label = tk.Label(root, text="Enter Version:").grid(row=1, column=1, sticky=tk.W, pady=15)

version_entry = tk.Entry(root, width=30, textvariable=versionValue)
version_entry.grid(row=1, column=2, sticky=tk.W, pady=15)

# Create "submit" button and destroy root on click
submit_btn = tk.Button(root, text="Submit", width=10, command=get_value).grid(row=2, column=1, columnspan=5, pady=25)

# Destroy root on "Enter" key press
version_entry.bind("<Return>", lambda e: get_value())

# Quits program on close
root.protocol("WM_DELETE_WINDOW", on_closing)


root.mainloop()

version_id = int(versionValue.get())

########################################################################################################################
