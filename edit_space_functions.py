from fileinput import filename
import xlwings as xw
import universal_functions as uf

# When the space graph workbook already exists....

# Called if space graph workbook does exist
def edit_space_graphs(space_alloc_table, bldg_space_table, file_name, node_rollup_df):

    tmgsp_list_qtrly = []
    tmgsp_list = []
    wafer_site_qtrly_list = []
    for site in uf.site_list:

        monthly_df, qtrly_df = uf.create_full_df(space_alloc_table, bldg_space_table, site, tmgsp_list, tmgsp_list_qtrly)
        xw_write(monthly_df, file_name, site, sort=False) # Write SITE monthly and quarterly dataframe to excel
        site_qtrly_df = make_xw_site_qtrly_graph(qtrly_df, site, file_name)
        wafer_site_qtrly_df = uf.make_wafer_site_qtrly_df(site_qtrly_df, site)
        wafer_site_qtrly_list.append(wafer_site_qtrly_df)        


    make_xw_tmgsp_mntly_table(tmgsp_list, file_name)
    make_xw_tmgsp_qtrly_graph(tmgsp_list_qtrly, file_name)

    if node_rollup_df is not None:
        wafer_node_df = uf.combine_wafer_qtrly_df(wafer_site_qtrly_list, node_rollup_df)
        xw_write(wafer_node_df, file_name, "SpaceAndCapacity", sort=False) # Write SITE monthly and quarterly dataframe to excel


# Xlwings makes SITE qtrly dataframe
def make_xw_site_qtrly_graph(site_qtrly, site, file_name):
    # Make and clean qtrly df
    site_qtrly_df_long = uf.make_site_qtrly_df(site_qtrly)
    site_qtrly_df = uf.clean_tmgsp_qtrly_df(site_qtrly_df_long)

    xw_write(site_qtrly_df, file_name, site, df_cell='A66')

    return site_qtrly_df


# Xlwings makes TMGSP qtrly dataframe
def make_xw_tmgsp_qtrly_graph(tmgsp_list_qtrly, file_name):
    # Make and clean qtrly df
    tmgsp_df_qtrly_long = uf.make_tmgsp_qtrly_df(tmgsp_list_qtrly)
    tmgsp_df_qtrly = uf.clean_tmgsp_qtrly_df(tmgsp_df_qtrly_long)
    sheet_name = "TMGSP_QTR"

    xw_write(tmgsp_df_qtrly, file_name, sheet_name)

    
# Xlwings makes TMGSP monthly dataframe
def make_xw_tmgsp_mntly_table(tmgsp_list, file_name):
    # Makes and formats tmgsp monthly df
    tmgsp_df = uf.make_tmgsp_mntly_df(tmgsp_list)
    sheet_name = "TMGSP"

    xw_write(tmgsp_df, file_name, sheet_name)


def xw_write(df, file_name, sheet_name, df_cell='A1', sort=True):
    # write df to excel sheet
    wb = xw.Book(file_name)

    if sort:
        df = uf.sort_df_rows(df)

    print(sheet_name)
    sheet = wb.sheets[sheet_name]
    sheet.range(df_cell).options(index=False).value = df