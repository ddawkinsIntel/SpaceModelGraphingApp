# Quarterly Functions below
import pandas as pd

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
