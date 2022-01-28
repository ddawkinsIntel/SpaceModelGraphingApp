# DB functions below

import pyodbc
import pandas as pd

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

    # Map SiteID to Site - stored procedure 'sml.pr_Site_Get' no params
    table_building_space['SiteID'] = table_building_space['SiteID'].map({2506: 'AZ', 2507: 'IR', 2508: 'IS', 2509: 'NM' , 2552: 'LTD', 2557: 'HVM1' })

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
def get_rpt_solver(conn, id_version, id_wif):
    sql = """\
    EXEC sml.pr_Rpt_SolverSpaceAlloc @VersionID = ?, @WIFID = ?
    """
    # test = cursor.execute(sql, params)
    print("here", id_version)
    table_space_alloc = pd.read_sql(sql, conn, params=[id_version, id_wif])

    return table_space_alloc


# Gets and Returns all needed DB data
def get_db_data(conn, id_version, id_wif):

    # Get space alloc table
    space_alloc_table = get_rpt_solver(conn, id_version, id_wif)

    # Get building space table
    bldg_space_table = get_building_space(conn, id_version)

    # Get the versions name(i.e. 2021Q3LRP Deramp)
    version_name = get_version_name(conn, id_version)

    return space_alloc_table, bldg_space_table, version_name
########################################################################################################################