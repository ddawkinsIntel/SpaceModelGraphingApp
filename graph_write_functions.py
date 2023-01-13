import universal_functions as uf
import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
import os
import xlwings as xw

# Process Color Dictionary
proc_color_dict = {1222: '#ED7D31', 1270: '#A5A5A5', 1272: '#FFC000', 1274: '#5B9BD5', 1276: 'purple', 1277: '#403152',
                   1278: '#663300', 1280: '#71FFBF', 1282: '#FAC090', 1284:'#B3A2C7', 1286:'#D99694'}


# color_file location for the graph
color_file = 'SpaceModelColorMap.xlsx'

# Create the 'graphs' folder, if it does not exist
if os.path.exists(color_file):

    color_map_df = pd.read_excel(color_file, index_col=[0])

    # Convert color_map_df into a dictionary
    color_map_dict = color_map_df.to_dict()
    color_map_dict = color_map_dict['Hex Color']

    print('Using user Color Map....')
    print('color map: ', color_map_dict, '\n')

else:
    color_map_dict = proc_color_dict
    print('Using standard color map')


# Add chart series
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


def make_chart(df, site_val, pd_writer, start_row=0, chart_cell="G12"):

    # Use site value as sheet name
    sheet_name = site_val

    # Sort df rows for excel graph aesthetic purposes
    sorted_df = uf.sort_df_rows(df)

    # write df to excel sheet
    sorted_df.to_excel(pd_writer, sheet_name=sheet_name, startrow=start_row, index=False)

    # Access the XlsxWriter workbook and worksheet objects from the dataframe.
    workbook = pd_writer.book
    worksheet = pd_writer.sheets[sheet_name]

    # Create a chart object.
    area_chart = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})

    # Get shape without including wafer volume rows
    non_wafer_df = df[~(df["CapacityType"].str.contains("WSVolume") | df["CapacityType"].str.contains("TotalSpaceUsed"))]

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

        elif colname == "DedicatedSpace":
            color = "#BBBCA0"

            # Area chart series
            add_area_chart_series(area_chart, color, sheet_name, row, ncols, start_row)

        else:
            process = int(colname[:4])
            color = color_map_dict.get(process, "#fc03df")
            

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

# Not in use - Attempted to make a stacked bar chart
def make_stacked_bar(df, site_val, pd_writer, start_row=0, chart_cell="G12"):

    # Use site value as sheet name
    sheet_name = site_val

    # # Sort df rows for excel graph aesthetic purposes
    # sorted_df = uf.sort_df_rows(df)
    sorted_df = df

    # write df to excel sheet
    sorted_df.to_excel(pd_writer, sheet_name="SpaceAndCapacity", startrow=start_row, index=False)
    # sorted_df.to_excel(pd_writer, sheet_name=sheet_name, startrow=start_row, index=False)

    # Access the XlsxWriter workbook and worksheet objects from the dataframe.
    workbook = pd_writer.book
    worksheet = pd_writer.sheets[sheet_name]

    # Create a chart object.
    area_chart = workbook.add_chart({'type': 'bar', 'subtype': 'stacked'})

    # Get shape without including wafer volume rows
    non_wafer_df = df[df["NodeScenario"].isin("SpaceModel")]

    # Number of columns
    ncols = non_wafer_df.shape[1]-1

    # Configure the series of the chart from the dataframe data.
    max_row = len(non_wafer_df)
    for i in range(max_row):

        row = i
        colname = sorted_df.iloc[row, 2]
        site_colname = sorted_df.iloc[row, 1]

        # if NodeScenario value does not equal space model make a line, else make stacked bar
        if colname != "SpaceModel":
            # Add line to chart
            line_chart1 = workbook.add_chart({'type': 'line'})

            color = 'black'

            # Line chart series
            add_line_chart_series(line_chart1, color, sheet_name, row, ncols, start_row)

            # Combine Charts
            area_chart.combine(line_chart1)
        elif site_colname == "AZ":
            color = '#C00000'

            # Area chart series
            add_area_chart_series(area_chart, color, sheet_name, row, ncols, start_row)

        elif site_colname == "IR":
            color = "#00B050"

            # Area chart series
            add_area_chart_series(area_chart, color, sheet_name, row, ncols, start_row)

        elif site_colname == "IS":
            color = "#00B050"

            # Area chart series
            add_area_chart_series(area_chart, color, sheet_name, row, ncols, start_row)

        elif site_colname == "HVM1":
            color = "#00B050"

            # Area chart series
            add_area_chart_series(area_chart, color, sheet_name, row, ncols, start_row)

        elif site_colname == "HVM2":
            color = "#00B050"

            # Area chart series
            add_area_chart_series(area_chart, color, sheet_name, row, ncols, start_row)

        elif site_colname == "HVM3":
            color = "#00B050"

            # Area chart series
            add_area_chart_series(area_chart, color, sheet_name, row, ncols, start_row)

        elif site_colname == "HVM4":
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
