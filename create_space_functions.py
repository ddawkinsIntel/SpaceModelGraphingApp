import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
import universal_functions as uf
import graph_write_functions as gwf
import edit_space_functions as esf

# When the space graph workbook does not exist....

# Starts ceation of and saves workbook
def create_space_graph_wb(file_path, space_alloc_table, bldg_space_table, node_rollup_df):

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
    create_space_graphs(writer, space_alloc_table, bldg_space_table, node_rollup_df)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()


# Creates space graphs
def create_space_graphs(writer, space_alloc_table, bldg_space_table, node_rollup_df):

    site_list = space_alloc_table.Site.unique()
    site_list.sort()
    site_list = sorted(site_list, key=len, reverse=False)

    tmgsp_list_qtrly = []
    tmgsp_list = []
    wafer_site_qtrly_list = []
    ####################################################################################################################
    for site in site_list:
        full_wide, full_long = uf.create_full_df(space_alloc_table, bldg_space_table, site, tmgsp_list, tmgsp_list_qtrly)

        # Writes df to sheet, then makes and saves chart
        gwf.make_chart(full_wide, site, writer)
        site_qtrly_df = make_pd_site_qtrly_graph(full_long, site, writer)

        wafer_site_qtrly_df = uf.make_wafer_site_qtrly_df(site_qtrly_df, site)
        wafer_site_qtrly_list.append(wafer_site_qtrly_df)   

    tmgsp_graphs(tmgsp_list, tmgsp_list_qtrly, writer)

    if node_rollup_df is not None:
        wafer_node_df = uf.combine_wafer_qtrly_df(wafer_site_qtrly_list, node_rollup_df)
        wafer_node_df.to_excel(writer, sheet_name="SpaceAndCapacity", index=False)
        # gwf.make_stacked_bar(wafer_node_df, "SpaceAndCapacity", writer, 65, 'G40')



# Pandas makes site qtrly graph
def make_pd_site_qtrly_graph(site_qtrly, site, writer):
    # Make and clean qtrly df - needs to be one line
    site_qtrly_df = uf.make_site_qtrly_df(site_qtrly)
    site_qtrly_df = uf.clean_tmgsp_qtrly_df(site_qtrly_df)

    # Makes and saves chart
    gwf.make_chart(site_qtrly_df, site, writer, 65, 'G40')

    return site_qtrly_df


# Pandas creates tmgsp monthly and quarterly graphs
def tmgsp_graphs(tmgsp_list, tmgsp_list_qtrly, writer):
    make_pd_tmgsp_mntly_graphs(tmgsp_list, writer)
    make_pd_tmgsp_qtrly_graph(tmgsp_list_qtrly, writer)


# Pandas makes tmgsp monthly graph - nothing is returned
def make_pd_tmgsp_mntly_graphs(tmgsp_list, writer):
    # Makes and formats tmgsp monthly df
    tmgsp_df = uf.make_tmgsp_mntly_df(tmgsp_list)
    # Makes and saves chart
    gwf.make_chart(tmgsp_df, "TMGSP", writer)


# Pandas makes tmgsp qtrly graph - nothing is returned
def make_pd_tmgsp_qtrly_graph(tmgsp_list_qtrly, writer):
    # Make and clean qtrly df
    tmgsp_df_qtrly_long = uf.make_tmgsp_qtrly_df(tmgsp_list_qtrly)
    tmgsp_df_qtrly = uf.clean_tmgsp_qtrly_df(tmgsp_df_qtrly_long)

    # Makes and saves chart
    gwf.make_chart(tmgsp_df_qtrly, "TMGSP_QTR", writer)
