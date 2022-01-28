

import pandas as pd
import xlwings as xw
import create_df_functions as df_func

# Select what site to iterate through
site_list = ['AZ', 'IR', 'IS', 'LTD', 'HVM1', 'HVM2', 'HVM3', 'HVM4', 'HVM5', 'NM']

# Process Color Dictionary
proc_color_dict = {1222: '#ED7D31', 1270: '#A5A5A5', 1272: '#FFC000', 1274: '#5B9BD5', 1276: 'purple', 1277: '#403152',
                   1278: '#663300', 1280: '#71FFBF', 1282: '#FAC090', 1284:'#B3A2C7', 1286:'#D99694'}

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
# Xlwings makes tmgsp qtrly graph - nothing is returned
def make_xw_tmgsp_qtrly_graph(tmgsp_list_qtrly, file_name):
    # Make and clean qtrly df
    tmgsp_df_qtrly_long = df_func.make_tmgsp_qtrly_df(tmgsp_list_qtrly)
    tmgsp_df_qtrly = df_func.clean_tmgsp_qtrly_df(tmgsp_df_qtrly_long)

    sheet_name = "TMGSP_QTR"

    xw_write(tmgsp_df_qtrly, file_name, sheet_name)


# Pandas makes tmgsp qtrly graph - nothing is returned
def make_pd_tmgsp_qtrly_graph(tmgsp_list_qtrly, writer):
    # Make and clean qtrly df
    tmgsp_df_qtrly_long = df_func.make_tmgsp_qtrly_df(tmgsp_list_qtrly)
    tmgsp_df_qtrly = df_func.clean_tmgsp_qtrly_df(tmgsp_df_qtrly_long)

    # Makes and saves chart
    make_chart(tmgsp_df_qtrly, "TMGSP_QTR", writer)


# Xlwings makes tmgsp qtrly graph - nothing is returned
def make_xw_site_qtrly_graph(site_qtrly, site, file_name):
    # Make and clean qtrly df
    site_qtrly_df_long = df_func.make_site_qtrly_df(site_qtrly)
    site_qtrly_df = df_func.clean_tmgsp_qtrly_df(site_qtrly_df_long)

    xw_write(site_qtrly_df, file_name, site, df_cell='A66')


# Pandas makes site qtrly graph - nothing is returned
def make_pd_site_qtrly_graph(site_qtrly, site, writer):
    # Make and clean qtrly df
    site_qtrly_df = df_func.make_site_qtrly_df(site_qtrly)
    site_qtrly_df = df_func.clean_tmgsp_qtrly_df(site_qtrly_df)

    # Makes and saves chart
    make_chart(site_qtrly_df, site, writer, 65, 'G40')


########################################################################################################################
# Xlwings makes tmgsp monthly graph - nothing is returned
def make_xw_tmgsp_mntly_table(tmgsp_list, file_name):
    # Makes and formats tmgsp monthly df
    tmgsp_df = df_func.make_tmgsp_mntly_df(tmgsp_list)

    sheet_name = "TMGSP"
    xw_write(tmgsp_df, file_name, sheet_name)

# Pandas makes tmgsp monthly graph - nothing is returned
def make_pd_tmgsp_mntly_graphs(tmgsp_list, writer):
    # Makes and formats tmgsp monthly df
    tmgsp_df = df_func.make_tmgsp_mntly_df(tmgsp_list)
    # Makes and saves chart
    make_chart(tmgsp_df, "TMGSP", writer)


########################################################################################################################
# Pandas creates tmgsp monthly and quarterly graphs
def tmgsp_graphs(tmgsp_list, tmgsp_list_qtrly, writer):
    make_pd_tmgsp_mntly_graphs(tmgsp_list, writer)
    make_pd_tmgsp_qtrly_graph(tmgsp_list_qtrly, writer)


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

        monthly_df, qtrly_df = df_func.create_full_df(space_alloc_table, bldg_space_table, site, tmgsp_list, tmgsp_list_qtrly)
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
        full_wide, full_long = df_func.create_full_df(space_alloc_table, bldg_space_table, site, tmgsp_list, tmgsp_list_qtrly)

        # Makes and saves chart
        make_chart(full_wide, site, writer)
        make_pd_site_qtrly_graph(full_long, site, writer)

    tmgsp_graphs(tmgsp_list, tmgsp_list_qtrly, writer)


# Called if space graph workbook does not exist
def create_space_graph_wb(file_path, space_alloc_table, bldg_space_table):

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
    create_space_graphs(writer, space_alloc_table, bldg_space_table)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
########################################################################################################################