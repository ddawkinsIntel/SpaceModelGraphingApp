import pandas as pd

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
# Monthly functions below
# Makes, formats and returns tmgsp monthly df
def make_tmgsp_mntly_df(tmgsp_list):
    # Concatenate list of dfs
    tmgsp_df = pd.concat(tmgsp_list)
    tmgsp_df = tmgsp_df.groupby(['CapacityType']).sum().reset_index()
    # tmgsp_df = tmgsp_df.append(tmgsp_df.sum(numeric_only=True).rename('Total')).reset_index()
    return tmgsp_df

