import streamlit as st
import pandas as pd
import numpy as np
import datetime
from datetime import date
import os
from os.path import exists
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image
from st_clickable_images import clickable_images
from openpyxl import load_workbook
import io
from io import BytesIO
import xlsxwriter


# '''
# MAIN TODOs
# - select tabs per wkbk for a given offer - output on corresponding sheet
# - lookup cat in cattype
#
# '''


row_ct = {}

# balloons = False
clicked = 'https://images-wixmp-ed30a86b8c4ca887773594c2.wixmp.com/f/d3789c8c-0874-407c-a457-03b147f59b18/der4qbq-d8305001-dc6d-441e-ab8c-132b6f35fe63.png/v1/fill/w_1176,h_627,strp/lighting_mcqueen_by_darkmoonanimation_der4qbq-fullview.png?token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJ1cm46YXBwOjdlMGQxODg5ODIyNjQzNzNhNWYwZDQxNWVhMGQyNmUwIiwiaXNzIjoidXJuOmFwcDo3ZTBkMTg4OTgyMjY0MzczYTVmMGQ0MTVlYTBkMjZlMCIsIm9iaiI6W1t7ImhlaWdodCI6Ijw9NjI3IiwicGF0aCI6IlwvZlwvZDM3ODljOGMtMDg3NC00MDdjLWE0NTctMDNiMTQ3ZjU5YjE4XC9kZXI0cWJxLWQ4MzA1MDAxLWRjNmQtNDQxZS1hYjhjLTEzMmI2ZjM1ZmU2My5wbmciLCJ3aWR0aCI6Ijw9MTE3NiJ9XV0sImF1ZCI6WyJ1cm46c2VydmljZTppbWFnZS5vcGVyYXRpb25zIl19.1kNCLY8PHLdG7veIX1SA9cSn8HgDT0DeYvtJUSVNXTo'

# Convert cell range to pandas df
def range_to_df(ws, remove_nan=True):
    # Read the cell values into a list of lists
    data_rows = []
    for row in ws:
        data_cols = []
        for cell in row:
            data_cols.append(cell.value)
        data_rows.append(data_cols)
    df = pd.DataFrame(data_rows[1:])
    df.columns = data_rows[0]
    if remove_nan:
        df.dropna(axis=1, how='all', inplace=True)
    return df

# def balloons():
#     balloons = True

# Layout
st.set_page_config(layout="wide")
st.title('Fenced Offers - Upload and KACHOW!!!')
passwd = st.text_input('Enter Password')
if passwd == 'lightning':
    st.markdown('We process **faster** than lightning!')
    st.warning("Please upload your Fenced Offer input worksheet below.")


    # Excel INPUT
    uploaded_files = st.file_uploader("Select all ship workbooks (use ctrl to select multiple files)", accept_multiple_files=True, key='upload01')
    if len(uploaded_files) > 0:
        st.warning("Click on Lightning McQueen to run process.")
    # for uploaded_file in uploaded_files:
    #     bytes_data = uploaded_file.read()
    #     st.write("filename:", uploaded_file.name)
    #     st.write(bytes_data)

    # The car, the myth, the legend, KACHOW
    clicked = clickable_images(
        [
            'https://images-wixmp-ed30a86b8c4ca887773594c2.wixmp.com/f/d3789c8c-0874-407c-a457-03b147f59b18/der4qbq-d8305001-dc6d-441e-ab8c-132b6f35fe63.png/v1/fill/w_1176,h_627,strp/lighting_mcqueen_by_darkmoonanimation_der4qbq-fullview.png?token=eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJ1cm46YXBwOjdlMGQxODg5ODIyNjQzNzNhNWYwZDQxNWVhMGQyNmUwIiwiaXNzIjoidXJuOmFwcDo3ZTBkMTg4OTgyMjY0MzczYTVmMGQ0MTVlYTBkMjZlMCIsIm9iaiI6W1t7ImhlaWdodCI6Ijw9NjI3IiwicGF0aCI6IlwvZlwvZDM3ODljOGMtMDg3NC00MDdjLWE0NTctMDNiMTQ3ZjU5YjE4XC9kZXI0cWJxLWQ4MzA1MDAxLWRjNmQtNDQxZS1hYjhjLTEzMmI2ZjM1ZmU2My5wbmciLCJ3aWR0aCI6Ijw9MTE3NiJ9XV0sImF1ZCI6WyJ1cm46c2VydmljZTppbWFnZS5vcGVyYXRpb25zIl19.1kNCLY8PHLdG7veIX1SA9cSn8HgDT0DeYvtJUSVNXTo'
        ],
        titles=[f"Image #{str(i)}" for i in range(5)],
        div_style={"display": "flex", "justify-content": "center", "flex-wrap": "wrap"},
        img_style={"margin": "5px", "height": "200px"},
    )

    #===============================================================================
    # This is where the magic happens (when lightning is clicked) - KACHOW
    if clicked == 0:


        #TODO: form dictionary for lookup tables
        ship_dict = {'DD':'Disney Dream', 'DF':'Disney Fantasy', 'DW':'Disney Wonder', 'DM':'Disney Magic', 'WW':'Disney Wish'}

        # load lookup data
        lkup_wb = load_workbook(filename = 'LookupWKBK.xlsx', data_only=True)
        lkup_ws = lkup_wb['Lookups']
        taxes_ws = lkup_wb['Taxes']
        lkup_itens = range_to_df(lkup_ws['A2':'F350'])
        lkup_taxes = range_to_df(taxes_ws['A1':'F500'])
        iten_dict, port_dict = {}, {}
        tax_dict = {'DD':{}, 'DF':{}, 'DM':{}, 'DW':{}, 'WW':{}}
        portfrom_dict = {'DD':{}, 'DF':{}, 'DM':{}, 'DW':{}, 'WW':{}}
        for index, row in lkup_itens.iterrows():
            iten_dict[row['GEOG_AREA_CODE']] = row['Lookup']
            port_dict[row['Code']] = row['Name']
        for index, row in lkup_taxes.iterrows():
            try:
                tax_dict[row['SHIP']][row['SAIL_FROM']] = row['GVT_TAX']
                portfrom_dict[row['SHIP']][row['SAIL_FROM']] = row['PORT_FROM']
            except:
                continue


        # iten_dict = {'WESTERN':'Western Caribbean', 'BAHAMAS':'Bahamas', 'MED':'Mediterranean'}

        # set starting row
        #may need a dict per offer for starting row.... add the val if it doesnt exist, update if it does.
        # enter_row = 9
        for uploaded_file in uploaded_files:
            # Load wkbk and sheet
            wb = load_workbook(filename = uploaded_file, data_only=True)
            # Get sheet names
            worksheets = wb.sheetnames
            #TODO: multiple sheets (one per offer, 'sail' in promoted fills in the lower table row 28, excluding G-J pricing)
            for sheet_name in worksheets:
                ws = wb[sheet_name]
                if sheet_name in list(row_ct.keys()):
                    continue
                else:
                    row_ct[sheet_name] = 9
                enter_row = row_ct[sheet_name]

                # Call pd convert function
                input = range_to_df(ws['A4':'AE100'])

                # Rename columns that are duplicates
                for n in ('FVGT', 'VGT', 'IGT', 'OGT'):
                    cols = []
                    count = 1
                    for column in input.columns:
                        if column == n:
                            cols.append(f'{n}_{count}')
                            count+=1
                            continue
                        cols.append(column)
                    input.columns = cols

                print(input.columns)
                print(input['FVGT_1'])
                print("========================================")

                # grab all entered rows
                offers = input[pd.to_numeric(input['FVGT_1'], errors='coerce').notnull()]
                #TODO: repeat for VGT, OGT, IGT and UNION

                # Build output from template (assuming MTO for now) - TODO: list of filepaths
                if sheet_name in ("CAST","DCL_CAST"):
                    file_path = 'Template_CAST.xlsx'
                else:
                    file_path = 'Template_MTO.xlsx'
                # TODO: for entries that are both TA and CAST, include in TAAP

                # Pull in template
                tb = load_workbook(filename = file_path)
                ts = tb['OUTPUT']

                eb = load_workbook(filename = 'Template_ExecSummary.xlsx')
                es = tb['OUTPUT']

                if len(offers) >= 27:
                    ts.move_range("A27:K39", rows=len(offers)-26) #TODO: does not work
                for i in range(len(offers)):
                    #TODO: sailstart will not align with sailfrom .. type error?
                    if sheet_name in ("CAST","DCL_CAST"):
                        ts['A7'] = str(date.today())
                        ship = offers.iloc[i]['Ship']
                        saildatefrom = offers.iloc[i]['Sail Start']
                        ts['A'+str(enter_row)] = ship_dict[ship]
                        ts['B'+str(enter_row)] = saildatefrom
                        days = offers.iloc[i]['Days']
                        ts['C'+str(enter_row)] = days
                        ts['E'+str(enter_row)] = iten_dict[offers.iloc[i]['Destination']]
                        #TODO: Fill in F for CAT ['Promoted?']
                        ts['D'+str(enter_row)] = port_dict[portfrom_dict[ship][saildatefrom]].capitalize()
                        # TODO: check for cat promoted
                        vf_0 = offers.iloc[i]['FVGT_1']
                        ts['G'+str(enter_row)] = vf_0
                        ts['H'+str(enter_row)] = tax_dict[ship][saildatefrom]
                        grat = 14.5
                        ts['I'+str(enter_row)] = grat * days
                        ts['J'+str(enter_row)] = vf_0*days+tax_dict[ship][saildatefrom]+grat
                        ts['K'+str(enter_row)] = offers.iloc[i]['Comments']
                        # Pull in TFE info - filter lkup_taxes by ship-sail combo

                        # Add to Executive Summary
                        es['B'+str(enter_row)] = ship_dict[offers.iloc[i]['Ship']]
                        es['C'+str(enter_row)] = offers.iloc[i]['Sail Start']
                        es['D'+str(enter_row)] = offers.iloc[i]['Days']
                        es['E'+str(enter_row)] = iten_dict[offers.iloc[i]['Destination']]
                        es['F'+str(enter_row)] = offers.iloc[i]['Promoted?']
                        vf_0 = offers.iloc[i]['FVGT_1']
                        es['G'+str(enter_row)] = vf_0
                        grat = 14.5
                        es['J'+str(enter_row)] = vf_0*days+tax_dict[ship][saildatefrom]+grat
                        # ts['K'+str(enter_row)] = offers.iloc[i]['Comments']



                        # set new enter_row
                        enter_row += 1
                        lower_enter_row = 0 #TODO: add logic for lower part of sheets "Sail" in comments

                    # date = datetime.datetime.strptime(offers.iloc[i]['Sail Start'], '%Y/%m/%d')
                    # st.write()
                    # tax_row = lkup_taxes[(lkup_taxes['SHIP']==ship_dict[offers.iloc[i]['Ship']]) & (lkup_taxes['SAIL_FROM']==offers.iloc[i]['Sail Start'])]
                    else:
                        tdy = datetime.date.today()
                        first_mon = tdy - datetime.timedelta(days = tdy.weekday())
                        second_mon = tdy + datetime.timedelta(days = -tdy.weekday(), weeks=1)
                        ts['F2'] = str(ws.cell(1, 1).value)+" Sheet"
                        ts['F4'] = "Rates Valid Monday, "+str(first_mon)+" through Monday, "+str(second_mon)
                        print(offers.iloc[i])
                        ship = offers.iloc[i]['Ship']
                        saildatefrom = offers.iloc[i]['Sail Start']
                        ts['C'+str(enter_row)] = ship_dict[offers.iloc[i]['Ship']]
                        ts['D'+str(enter_row)] = offers.iloc[i]['Sail Start']
                        days = offers.iloc[i]['Days']
                        ts['E'+str(enter_row)] = days
                        ts['F'+str(enter_row)] = iten_dict[offers.iloc[i]['Destination']]
                        ts['G'+str(enter_row)] = port_dict[portfrom_dict[ship][saildatefrom]].capitalize()
                        ts['H'+str(enter_row)] = offers.iloc[i]['Promoted?']
                        # TODO: check for cat promoted
                        vf_0 = int(offers.iloc[i]['VGT_1'])
                        ts['I'+str(enter_row)] = vf_0
                        ts['J'+str(enter_row)] = vf_0*days
                        ts['K'+str(enter_row)] = 2*vf_0*days
                        # Pull in TFE info - filter lkup_taxes by ship-sail combo
                        ts['L'+str(enter_row)] = tax_dict[ship][saildatefrom]
                        # set new enter_row
                        enter_row += 1
                ts.title = sheet_name
                tb.save("temp_"+sheet_name)

                # Executive Summary
                es.title = sheet_name
                eb.save("exec_"+sheet_name)


        st.success("KACHOW we're done! - - - Download offer sheets below.")

        for file in os.listdir("./"):
            if file.startswith("temp"):
                with open(file, 'rb') as my_file: #TODO: need to check for loop logic for wkbk to sheet, ships with offers sheet to offers with ships, single sheet each... need to store this info somehow
                    st.download_button(
                        label = 'Download '+file[5::],
                        data = my_file,
                        file_name = file[5::]+'_'+str(date.today())+'.xlsx',
                        mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        # on_click = balloons()
                        )
            # if file.startswith("exec"):
            #     with open(file, 'rb') as my_file: #TODO: need to check for loop logic for wkbk to sheet, ships with offers sheet to offers with ships, single sheet each... need to store this info somehow
            #         st.download_button(
            #             label = 'Download Exec Summary',
            #             data = my_file,
            #             file_name = 'ExecutiveOfferSummary_'+str(date.today())+'.xlsx',
            #             mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            #             # on_click = balloons()
            #             )



        # if balloons:
        #     st.balloons()
