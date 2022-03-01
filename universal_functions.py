import pandas as pd

# Select what site to iterate through
site_list = ['AZ', 'IR', 'IS', 'LTD', 'HVM1', 'HVM2', 'HVM3', 'HVM4', 'HVM5', 'NM']

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
########################################################################################################################






########################################################################################################################
# Monthly functions below
# Makes, formats and returns tmgsp monthly df
def make_tmgsp_mntly_df(tmgsp_list):
    # Concatenate list of dfs
    tmgsp_df = pd.concat(tmgsp_list)
    tmgsp_df = tmgsp_df.groupby(['CapacityType']).sum().reset_index()

    return tmgsp_df
########################################################################################################################





########################################################################################################################
def make_wafer_site_qtrly_df(site_qtrly_df, site):
    wafer_site_qtrly_df = site_qtrly_df[site_qtrly_df["CapacityType"].str.contains("WSVolume")]
    wafer_site_qtrly_df["CapacityType"] = wafer_site_qtrly_df["CapacityType"].str.replace(" WSVolume", "")
    wafer_site_qtrly_df.rename({'CapacityType': 'Node'}, axis=1, inplace=True)

    col_list = wafer_site_qtrly_df.columns
    new_col_list = []
    for col in col_list:
        if len(col) != 6:
            new_col_list.append(col)
        else:
            year = col[2:4]
            qtr = col[-2:]
            new_col = "'"+ year + "-" + qtr
            new_col_list.append(new_col)

    wafer_site_qtrly_df.columns = new_col_list

    wafer_site_qtrly_df["Scenario"] = site
    wafer_site_qtrly_df["NodeScenario"] = "SpaceModel"

    # pop wafer_site_qtrly_df["NodeScenario"] and place at index 0
    wafer_site_qtrly_df.insert(1, "NodeScenario", wafer_site_qtrly_df.pop("NodeScenario"))
    # pop wafer_site_qtrly_df["Scenario"] and place at index 1
    wafer_site_qtrly_df.insert(1, "Scenario", wafer_site_qtrly_df.pop("Scenario"))

    return wafer_site_qtrly_df

def combine_wafer_qtrly_df(wafer_qtrly_site_list, node_rollup_df):
    # Concatenate list of dfs, fillna with 0 and groupby
    wafer_df_full = pd.concat(wafer_qtrly_site_list)
    wafer_df_full = wafer_df_full.fillna(0)

    node_rollup_df = format_node_rollup(node_rollup_df)
    wafer_node_df = pd.concat([wafer_df_full, node_rollup_df], ignore_index=True)

    return wafer_node_df


def format_node_rollup(node_rollup):
    scenario = "SLRP Q3 '21"
    # node_rollup = node_rollup[node_rollup["Scenario"] == scenario]
    # remove_cols_list = ['12/31/1899', 'Scenario']
    # node_rollup.drop(remove_cols_list, axis=1, inplace=True)

    node_rollup.rename({'12/31/1899': 'NodeScenario'}, axis=1, inplace=True)
    node_rollup.insert(2, "NodeScenario", node_rollup.pop("NodeScenario"))
    node_rollup_df = node_mapping(node_rollup)
    node_rollup_df = node_rollup_df[node_rollup_df['Node'].notna()]
    node_rollup_df = node_rollup_df.drop_duplicates(subset=['NodeScenario'], keep='last')
    node_rollup_df = node_rollup_df.fillna(0)

    return node_rollup_df

def node_mapping(df_filtered):

    # List processes that belong to a process/node label
    p1225_26 = ('1225', '1226')
    p1214_16_17_42_43 = ('1214', '1216', '1217', '1242', '1243')
    p1240 = ('1240', '1241', '1250')
    _32nm = ('1268', '1269')
    # _22nm = ('1270', '1271', '1270 Trade', '1271 Trade', '22nm')
    _22nm = ('22nm', '22nm')
    p1222 = ('1222', 'P1222', 'P1222 ULP')
    # _14nm = ('1272', '1273', '1272 Trade', '1273 Trade', '14nm')
    _14nm = ('14nm', '14nm')
    # _10nm = ('1274', '1275', '10nm')
    _10nm = ('10nm', '10nm')
    p1276 = ('1276', 'P1276')
    p1277 = ('1277', 'P1277')
    p1278_79 = ('1278', 'P1278', '1279')
    # p1278_79 = ('P1278-79', 'P1278-79')
    # p1280_81 = ('1280', 'P1280', '1281', 'P1280')
    p1280_81 = ('P1280-81', 'P1280-81')
    p1282_83 = ('P1282-83', 'P1282-83')
    _45nm = ('1266', '1266')
    p1227 = ('P1227', 'p1227')

    NodeTransDict = {}

    # Adds process/node labels' process list as a value and the "process/node" label as a value to the dictionary
    NodeTransDict.update(dict.fromkeys(p1225_26, "p1225_26"))
    NodeTransDict.update(dict.fromkeys(p1240, "1240"))
    NodeTransDict.update(dict.fromkeys(p1214_16_17_42_43, "p1214_16_17_42_43"))
    NodeTransDict.update(dict.fromkeys(_32nm, "32nm"))
    NodeTransDict.update(dict.fromkeys(_22nm, "1270"))
    NodeTransDict.update(dict.fromkeys(p1222, "1222"))
    NodeTransDict.update(dict.fromkeys(_14nm, "1272"))
    NodeTransDict.update(dict.fromkeys(_10nm, "1274"))
    NodeTransDict.update(dict.fromkeys(p1276, "1276"))
    NodeTransDict.update(dict.fromkeys(p1277, "1277"))
    NodeTransDict.update(dict.fromkeys(p1278_79, "1278"))
    NodeTransDict.update(dict.fromkeys(p1280_81, "1280"))
    NodeTransDict.update(dict.fromkeys(p1282_83, "1282"))
    NodeTransDict.update(dict.fromkeys(_45nm, "45nm"))
    NodeTransDict.update(dict.fromkeys(p1227, "1227"))


    # Maps values in the "Process" column to the proper "Process/Node" label for the Space Model if the value
    # ...in "Process" isn't in the list of processes that belong to a "Process/Node" label" the value becomes null
    # df_filtered["Process"] = df_filtered["Process"].astype(str).apply(lambda x: NodeTransDict.get(x)) # OG line
    df_filtered["Node"] = df_filtered["Node"].astype(str).apply(lambda x: NodeTransDict.get(x)) # Mod'd line for RoadmapInputs file

    return df_filtered