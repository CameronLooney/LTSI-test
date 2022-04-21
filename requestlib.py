# DISCLAIMER: The code is in no way efficient, it is designed to merely complete the task. It is also not finished as
# some hard coded logic needs to be changed. It is simply the quick and dirty solution to the problem at hand
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import numpy as np
import io

# This will be a global variable to keep track of any problems along the process. This is needed to ensure the file
# will not be downloaded unless it has no errors.
error_count = 0


# Title of the Application

st.write("""

    # LTSI Tool 
    ## Instructions\n 
    - First Save all files as .xlsx \n
    - **Upload 1:** Raw LTSI download \n
    - **Upload 2:** Upload Updated Open Order File from Yesterday \n
    - **Upload 3:** Vlookup file with MPN's 
    - **Upload 4:** Upload Raw Backlog\n""")
# Need to uploads to generate Open Orders, one is a helper file which is used for computation and feedback.
# The master is the file downloaded from FrontEnd each day
upload_ltsi = st.file_uploader("Upload Raw LTSI Status File", type="xlsx")
upload_previous_open_orders = st.file_uploader("Upload Yesterdays Open Orders", type="xlsx")
upload_previous_helper = st.file_uploader("Upload MPN File", type="xlsx")

master = st.file_uploader("Upload Raw File", type="xlsx")
if st.button("Generate LTSI File"):
    if upload_ltsi is None:
        st.error("ERROR: Please upload File")

    if upload_ltsi is not None and upload_previous_open_orders is not None:
        upload = pd.read_excel(upload_ltsi, sheet_name=0, engine="openpyxl")
        open_orders = pd.read_excel(upload_previous_open_orders, sheet_name=0, engine="openpyxl")

        valid = upload[["salesOrderNum"]]
        valid["Valid in LTSI Tool"] = "TRUE"

        valid["salesOrderNum"] = valid["salesOrderNum"].astype(str)
        import re

        valid["salesOrderNum"] = [re.sub(r"[a-zA-Z]", "", x) for x in valid["salesOrderNum"]]
        valid = valid[valid["salesOrderNum"] != '']

        status_col_num = open_orders.columns.get_loc("Status (SS)")
        feedback_length = len(open_orders.columns)
        complete_feedback = [8, status_col_num]
        i = 34
        while i < feedback_length:
            complete_feedback.append(i)
            i += 1

        yesterday = open_orders.iloc[:, complete_feedback]
    # These ensure that two files have been uploaded
    if upload_previous_helper is None:
        st.error("ERROR: Please upload a viable upload_previous_helperiliary file to continue.")
    if master is None:
        st.error("ERROR: Please upload your raw LTSI download file to continue.")
    # If both files are uploaded we can begin the computation
    if upload_previous_helper is not None and master is not None:

        # read in the excel worksheets into separate pandas daatframes
        def read_excel_files(upload_previous_helper, master):
            # reading in files is time-consuming
            vlookup = pd.read_excel(upload_previous_helper, sheet_name=0, engine="openpyxl")
            previous = yesterday
            TF = valid
            master = pd.read_excel(master, sheet_name=0, engine="openpyxl")
            return vlookup, previous, TF, master

        vlookup, previous, TF, master = read_excel_files(upload_previous_helper, master)

        # this is required as some Dates are left blank and thus were lost
        # in later data manipulation
        def vlookup_date_fill(vlookup):
            vlookup.rename(columns={'MPN': 'material_num'}, inplace=True)
            vlookup['Date'] = vlookup['Date'].fillna("01.01.90")
            vlookup['Date'] = pd.to_datetime(vlookup.Date, dayfirst=True)
            vlookup['Date'] = [x.date() for x in vlookup.Date]
            vlookup['Date'] = pd.to_datetime(vlookup.Date)
            return vlookup

        vlookup = vlookup_date_fill(vlookup)

        # Functions job is to perform a vlookup function between open orders and vlookup worksheet
        def master_vlookup_merge(master, vlookup):
            master = master.merge(vlookup, on='material_num', how='left')
            return master

        # master = master_vlookup_merge(master,vlookup)

        # Function to filter out old LTSI and drop based on date
        def drop_old_dates(master):
            # if the date the order was placed is before it became valid LTSI drop from df
            rows = master[master['Date'] > master['ord_entry_date']].index.to_list()
            master = master.drop(rows).reset_index()
            return master

        # master = drop_old_dates(master)

        # when viewing the excel output file some ints in the blocks were being converted to dates in excel.
        # This function converts the columns to strings to circumvent this issue
        def order_block_type_converter(master):
            # change the column types and delete nan strings
            master['sch_line_blocked_for_delv'] = master['sch_line_blocked_for_delv'].astype(str)
            master['sch_line_blocked_for_delv'] = master['sch_line_blocked_for_delv'].replace("nan", "")
            master['del_blk'] = master['del_blk'].astype(str)
            master['del_blk'] = master['del_blk'].replace("nan", "")
            return master

        # master = order_block_type_converter(master)

        # Function to reduce row number, logic behind this function is if an order is 6 months old and still blocked
        # it is most certainly not valid open LTSI
        def delete_old_blocked_orders(master):
            six_months = datetime.now() - timedelta(188)
            rows_94 = master[
                (master['ord_entry_date'] < six_months) & (
                        master["sch_line_blocked_for_delv"] == 94)].index.to_list()
            master = master.drop(rows_94).reset_index(drop=True)
            return master

        # master = delete_old_blocked_orders(master)

        # logic is similar here. If the order is more than a year old is isnt going to be a valid open order so drop the row
        def delete_year_old_orders(master):
            twelve_months = datetime.now() - timedelta(365)
            rows_old = master[(master['ord_entry_date'] < twelve_months)].index.to_list()
            master = master.drop(rows_old).reset_index(drop=True)
            return master

        # master = delete_year_old_orders(master)

        # If we have no qty left we cant fulfill order so drop
        def delete_no_qty(master):
            master = master.loc[master['remaining_qty'] != 0]
            return master

        # master =delete_no_qty(master)

        # this is ugly hard coded logic but as of now these rows are escaping through the tests
        # this needs to be fixed
        def drop_miscellaneous(master):
            country2021drop = master[(master['ord_entry_date'].dt.year == 2021) & (master['country'].isin(
                ['Germany', 'Spain', "Turkey", "Belgium / Luxembourg", "Switzerland"]))].index.to_list()
            master = master.drop(country2021drop).reset_index(drop=True)
            return master

        # master = drop_miscellaneous(master)

        # most columns arent required in output so drop unneeded ones
        # consider calling this before dff manipulation might be faster
        def columns_to_keep():
            cols = ['sales_org', 'country', 'cust_num', 'customer_name', 'sales_dis', 'rtm',
                    'sales_ord', 'sd_line_item',
                    'order_method', 'del_blk', 'cust_req_date', 'ord_entry_date',
                    'cust_po_num', 'ship_num', 'ship_cust', 'ship_city', 'plant',
                    'material_num', 'brand', 'lob', 'project_code', 'material_desc',
                    'mpn_desc', 'ord_qty', 'shpd_qty', 'delivery_qty', 'remaining_qty',
                    'delivery_priority', 'opt_delivery_qt', 'rem_mod_opt_qt',
                    'sch_line_blocked_for_delv']
            return cols

        def drop_unneeded_cols(master):

            # APPLY REDUCTION
            reduced = master[columns_to_keep()]
            return reduced
            # master = drop_unneeded_cols(master)

        # alternative to above type converter this might be more convienient
        # only keep one
        def block_converter_alternative(reduced):
            reduced['del_blk'] = np.where(pd.isnull(reduced['del_blk']), reduced['del_blk'],
                                          reduced['del_blk'].astype(str))
            reduced['sch_line_blocked_for_delv'] = np.where(pd.isnull(reduced['sch_line_blocked_for_delv']),
                                                            reduced['sch_line_blocked_for_delv'],
                                                            reduced['sch_line_blocked_for_delv'].astype(str))

            return reduced

        # master = block_converter_alternative(master)

        # vlookup LTSI tool status worksheet with the open orders worksheet
        # this creates the validity column
        def valid_in_LTSI_tool(reduced):
            reduced.rename(columns={'sales_ord': 'salesOrderNum'}, inplace=True)
            reduced['salesOrderNum'] = reduced['salesOrderNum'].astype(int)
            TF['salesOrderNum'] = TF['salesOrderNum'].astype(int)

            reduced['holder'] = reduced.groupby('salesOrderNum').cumcount()
            TF['holder'] = TF.groupby('salesOrderNum').cumcount()
            merged = reduced.merge(TF, how='left').drop('holder', 1)
            return merged

        # master = valid_in_LTSI_tool(master)

        # join two columns and generate a new column to act as key
        def generate_unique_key(merged):
            # insert at index 8 to keep same layout
            idx = 8
            new_col = merged['salesOrderNum'].astype(str) + merged['sd_line_item'].astype(str)
            # insert merged column to act as unique key
            merged.insert(loc=idx, column='Sales Order and Line Item', value=new_col)
            merged['Sales Order and Line Item'] = merged['Sales Order and Line Item'].astype(int)
            return merged

        # master = generate_unique_key(master)

        # this generates the new validity column
        def generate_validity_column(merged):
            if 'Valid in LTSI Tool' not in merged:
                merged.rename(columns={'Unnamed: 1': 'Valid in LTSI Tool'}, inplace=True)
            merged["Valid in LTSI Tool"].fillna("FALSE", inplace=True)
            mask = merged.applymap(type) != bool
            dict = {True: 'TRUE', False: 'FALSE'}
            merged = merged.where(mask, merged.replace(dict))
            return merged

        # master = generate_validity_column(master)

        # generate and add a status column based on certain conditions]
        def generate_status_column(merged):
            conditions = [merged["del_blk"] != "",
                          merged["sch_line_blocked_for_delv"] != "",
                          merged['order_method'] == "Manual SAP",
                          merged['delivery_priority'] == 13,
                          merged["Valid in LTSI Tool"] == "TRUE",
                          ]
            outputs = ["Blocked", "Blocked", "Shippable", "Shippable", "Shippable"]
            result = np.select(conditions, outputs, "Under Review with C-SAM")
            result = pd.Series(result)
            merged['Status (SS)'] = result
            return merged

        # master = generate_status_column(master)
        def new_sdm_feedback(merged):
            merged["Action (SDM)"] = ""
            merged["Comments(SDM)"] = ""
            merged["Estimated DN Date"] = ""
            return merged

        def generate_sdm_feedback(merged):
            feedback = previous.drop('Status (SS)', 1)
            merged = merged.merge(feedback, on='Sales Order and Line Item', how='left')

            return merged

        # master = generate_sdm_feedback(master)

        def scheduled_out(merged):

            ten_days = datetime.now() + timedelta(10)
            merged.loc[(merged['cust_req_date'] > ten_days) & (
                    merged['Status (SS)'] == 'Shippable') & (
                               merged["Valid in LTSI Tool"] == 'TRUE'), 'Status (SS)'] = 'Scheduled Out'
            return merged

        def status_override(merged):
            action_sdm = merged.columns[37]
            merged[action_sdm] = merged[action_sdm].str.lower()
            merged[action_sdm] = merged[action_sdm].fillna("0")

            merged['Status (SS)'] = np.where(merged[action_sdm].str.contains('cancel', regex=False),
                                             'To be cancelled / reduced', merged['Status (SS)'])
            merged['Status (SS)'] = np.where(merged[action_sdm].str.contains('block', regex=False),
                                             'Blocked', merged['Status (SS)'])
            merged[action_sdm] = merged[action_sdm].astype(str)
            merged[action_sdm].replace(['0', '0.0'], '', inplace=True)
            return merged

        def unique_status(merged):
            standard_status = ["Shippable", "Blocked", "To be cancelled / reduced", "Under Review with C-SAM",
                               "Scheduled Out"]
            prev = previous[~previous['Status (SS)'].isin(standard_status)]
            result = prev.groupby('Sales Order and Line Item')['Status (SS)'].apply(list).to_dict()
            print(result)
            for key, value in result.items():
                merged["Status (SS)"] = np.where(merged["Sales Order and Line Item"] == key, value,
                                                 merged["Status (SS)"])
            return merged

        def open_orders_generator(master):
            step1 = master_vlookup_merge(master, vlookup)
            step2 = drop_old_dates(step1)
            step3 = order_block_type_converter(step2)
            step4 = delete_old_blocked_orders(step3)
            step5 = delete_year_old_orders(step4)
            step6 = delete_no_qty(step5)
            step7 = drop_miscellaneous(step6)
            step8 = drop_unneeded_cols(step7)
            step9 = valid_in_LTSI_tool(step8)
            step10 = generate_unique_key(step9)
            step11 = generate_validity_column(step10)
            step12 = generate_status_column(step11)
            step13 = scheduled_out(step12)
            step14 = new_sdm_feedback(step13)
            step15 = generate_sdm_feedback(step14)
            step16 = status_override(step15)
            finished = unique_status(step16)
            print(finished["Status (SS)"].unique())
            cols = columns_to_keep()
            cols.remove('sales_ord')
            cols.append('salesOrderNum')
            finished.drop_duplicates(subset=cols, keep='first', inplace=True)
            return finished

        def write_to_excel(merged):
            # Writing df to Excel Sheet
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # Write each dataframe to a different worksheet.
                # data["Date"] = pd.to_datetime(data["Date"])

                # pd.to_datetime('date')
                merged.to_excel(writer, sheet_name='Sheet1', index=False)
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                formatdict = {'num_format': 'dd/mm/yyyy'}
                fmt = workbook.add_format(formatdict)
                worksheet.set_column('K:K', None, fmt)
                worksheet.set_column('L:L', None, fmt)
                # Light yellow fill with dark yellow text.
                number_rows = len(merged.index) + 1
                yellow_format = workbook.add_format({'bg_color': '#FFEB9C'})
                worksheet.conditional_format('A2:AH%d' % (number_rows),
                                             {'type': 'formula',
                                              'criteria': '=$AH2="Under Review with C-SAM"',
                                              'format': yellow_format})
                worksheet.conditional_format('A2:AH%d' % (number_rows),
                                             {'type': 'formula',
                                              'criteria': '=$AH2="Under Review with CSAM"',
                                              'format': yellow_format})
                red_format = workbook.add_format({'bg_color': '#ffc7ce'})
                worksheet.conditional_format('A2:AH%d' % (number_rows),
                                             {'type': 'formula',
                                              'criteria': '=$AH2="Blocked"',
                                              'format': red_format})
                # C0C0C0
                green_format = workbook.add_format({'bg_color': '#c6efce'})
                worksheet.conditional_format('A2:AH%d' % (number_rows),
                                             {'type': 'formula',
                                              'criteria': '=$AH2="Shippable"',
                                              'format': green_format})
                worksheet.conditional_format('A2:AH%d' % (number_rows),
                                             {'type': 'formula',
                                              'criteria': '=$AH2="Scheduled Out"',
                                              'format': green_format})
                grey_format = workbook.add_format({'bg_color': '#C0C0C0'})
                worksheet.conditional_format('A2:AH%d' % (number_rows),
                                             {'type': 'formula',
                                              'criteria': '=$AH2="To be cancelled / reduced"',
                                              'format': grey_format})

                for column in merged:
                    column_width = max(merged[column].astype(str).map(len).max(), len(column))
                    col_idx = merged.columns.get_loc(column)
                    writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)
                    worksheet.autofilter(0, 0, merged.shape[0], merged.shape[1])
                worksheet.set_column(11, 12, 20)
                worksheet.set_column(12, 13, 20)
                worksheet.set_column(13, 14, 20)
                header_format = workbook.add_format({'bold': True,
                                                     'bottom': 2,
                                                     'bg_color': '#0AB2F7'})

                # Write the column headers with the defined format.
                for col_num, value in enumerate(merged.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                my_format = workbook.add_format()
                my_format.set_align('left')

                worksheet.set_column('N:N', None, my_format)
                writer.save()
                today = datetime.today()
                d1 = today.strftime("%d/%m/%Y")
                st.write("Download Completed File:")
                st.download_button(
                    label="Download Excel worksheets",
                    data=buffer,
                    file_name="LTSI_file_" + d1 + ".xlsx",
                    mime="application/vnd.ms-excel"
                )

        finished = open_orders_generator(master)
        write_to_excel(finished)
