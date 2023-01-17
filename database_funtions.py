# DB functions below

import pyodbc
import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
import warnings
warnings.filterwarnings("ignore")
import os

# Initiates DB connection - returns a "conn"
def initiate_conn():

    # config_file location for the graph
    config_file = 'ConfigFile.xlsx'

    # Create the 'graphs' folder, if it does not exist
    if os.path.exists(config_file):

        config_df = pd.read_excel(config_file, index_col=[0])
        # Convert config_df into a dictionary
        config_dict = config_df.to_dict()
        config_dict = config_dict['Value']

        print('Using config file....')
        print('Config File: ', config_dict)
        print('\n')

    else:
        print('Please include config file in the same location as the .exe... \n')

    # Initiate Connection to the db
    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=' + config_dict['Server'] + ';'
                          'Database=' + config_dict['Database'] + ';'
                          'UID=' + config_dict['UID'] + ';'
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

# Returns wif info table
def get_wif_info(conn, id_version):
    sql = """\
    EXEC sml.pr_Wif_Get @VersionID = ?
    """

    table_wif = pd.read_sql(sql, conn, params=[id_version])

    return table_wif

# From version table, returns the version name
def get_version_name(conn, id_version, id_wif):

    # Get version and wif info
    version_info = get_version_info(conn)
    wif_info = get_wif_info(conn, id_version)

    # Get version name from version info table by filtiring by the versionId
    version = version_info[version_info["VersionID"] == int(id_version)]

    # Get and store the VersionName value from version table
    version_name = version["VersionName"].values
    version_name = version_name[0]

    # If a wif value is entered (i.e. wif value is not zero), get the wif name
    if id_wif != "0":
        # Get wif name from wif info, using the versionId
        wif = wif_info[wif_info["WifId"] == int(id_wif)]

        wif_name = wif["Name"].values
        wif_name = wif_name[0]
    
    else:
        wif_name = ""


    return version_name, wif_name


# HF - get_building_space: Returns building space table
def read_bldg_space(conn, id_version):
    sql = """\
    EXEC sml.pr_BuildingSpace_Get @VersionID = ?
    """

    table_building_space = pd.read_sql(sql, conn, params=[id_version])

    return table_building_space


def get_site(conn):
    sql = """\
    EXEC sml.pr_Site_Get
    """

    site_df = pd.read_sql(sql, conn)

    return site_df 

# HF - get_building_space: Modifies and returns building table_building_space
def space_modify_df(site_id_map, table_building_space):
    
    # Merge to add site to table_building_space df
    table_building_space = table_building_space.merge(site_id_map, how='left', left_on='SiteID', right_on='ID')
    # Replace SiteID values with Site values - change later, Site should be name used not SiteID,
    # Didn't change becasuse I didn't want to break anything when intially implementing get_site
    table_building_space['SiteID'] = table_building_space['Site']

    # Columns to remove from df
    remove_cols_list = ['Site', 'ID']
    table_building_space.drop(remove_cols_list, axis=1, inplace=True)

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

    # Get site to id mapping
    site_id_map = get_site(conn)

    table_building_space = read_bldg_space(conn, id_version)

    table_building_space = space_modify_df(site_id_map, table_building_space)

    return table_building_space


# Returns space allocation table
def get_rpt_solver(conn, id_version, id_wif):
    sql = """\
    EXEC sml.pr_Rpt_SolverSpaceAlloc @VersionID = ?, @WIFID = ?
    """
    table_space_alloc = pd.read_sql(sql, conn, params=[id_version, id_wif])

    return table_space_alloc


# Gets and Returns all needed DB data
def get_db_data(conn, id_version, wif_counter):

    # Get wif info
    wif_info_table = get_wif_info(conn, id_version)

    wif_counter = int(wif_counter)
    
    if wif_counter == 0:
        id_wif = "0"

    else:
        wif_info_table = wif_info_table[wif_info_table["WIFCounter"] == wif_counter]

        id_wif = wif_info_table.iloc[0]['WifId']
        id_wif = str(id_wif)

    # Get space alloc table
    space_alloc_table = get_rpt_solver(conn, id_version, id_wif)

    # Get building space table
    bldg_space_table = get_building_space(conn, id_version)

    # Get the versions name(i.e. 2021Q3LRP Deramp)
    version_name, wif_name = get_version_name(conn, id_version, id_wif)

    return space_alloc_table, bldg_space_table, version_name, wif_name
########################################################################################################################