#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Mar 24 11:16:49 2025

@author: ashutoshgoenka
"""

import pandas as pd
import numpy as np
import streamlit as st

# from exif import Image as Image2


#import os
import zipfile
from zipfile import ZipFile, ZIP_DEFLATED
#import pathlib
#import shutil
#import docx
#import docxtpl
import random
from random import randint
from streamlit import session_state
import openpyxl
from openpyxl import load_workbook

st.set_page_config(layout="wide")




state = session_state
if "key" not in state:
    state["key"] = str(randint(1000, 100000000))

if "photo_saved" not in state:
    state["photo_saved"] = False

if "sample_file" not in state:
    state["sample_file"] = False
    
    
if "location_file" not in state:
    state["location_file"] = False

if "page_first_loaded" not in state:
    state["page_first_loaded"] = True
    
if "row_no" not in state:
    state["row_no"] = 1

    





cntr=1    
st.title("Flight Accounts Reconciliation")
data_file = st.file_uploader("Upload Riya Ledger/Data File", type=['csv','xlsx'],accept_multiple_files=False,key="fileUploader_data")
booked_history_file = st.file_uploader("Upload Riya Passenger Booked History File", type=['csv','xlsx'],accept_multiple_files=False,key="fileUploader_booked_history")
pas_master_file = st.file_uploader("Upload Riya Passenger Data Master File", type=['csv','xlsx'],accept_multiple_files=False,key="fileUploader_pass")
master_record_file = st.file_uploader("Upload Riya Master Record File", type=['csv','xlsx'],accept_multiple_files=False,key="fileUploader_master_history")
if pas_master_file is not None and data_file is not None and booked_history_file is not None and master_record_file is not None:
    df1 = pd.read_excel(data_file)
    df_passenger_master = pd.read_excel(pas_master_file)
    df_b1 = pd.read_excel(booked_history_file)
    df_existing = pd.read_excel(master_record_file)
    df_passenger_master.rename(columns = {'Name':'lead passenger'}, inplace = True)
    st.write("Ledger Sheet Display")
    st.write(df1)
    st.write("Passenger Master List")
    st.write(df_passenger_master)
    df1["Diff"] = df1["DateTime"].diff() 
    df1["SamePNR"] = df1["AirlinePNR"] == df1["AirlinePNR"].shift()
    df1[["SamePNR"]]= df1[["SamePNR"]].shift(-1) 
    df1[["Diff"]]= df1[["Diff"]].shift(-1) 
    df1["DateTimeNew"] = df1["DateTime"] + df1["Diff"]
    df1["DateTimeNew"] = np.where(df1["SamePNR"] == True, df1["DateTimeNew"], df1["DateTime"])
    df1["Diff"] = df1["Diff"].astype('int64').astype(int)/1000000000
    df1["DateTimeNew"] = np.where(df1["Diff"] < 2, df1["DateTimeNew"], df1["DateTime"] )
    df1["DateTime"] = df1["DateTimeNew"] 
    df1.drop("Diff",axis=1 ,inplace=True)
    df1.drop("SamePNR",axis=1 ,inplace=True)
    df1.drop(0,axis=0,inplace=True)
    opening_balance = df1["Remaining"].values[0] - df1["CreditAmount"].values[0] + df1["DebitAmount"].values[0]
    df1.drop("AgentId", axis=1, inplace=True)
    df1.drop("Ref", axis=1, inplace=True)
    df1.drop("Agency Name", axis=1, inplace=True)
    df2 = df1.assign(Value = lambda x: x.CreditAmount - x.DebitAmount) 
    df2["TransactionType"] = df2["TransactionType"].fillna("Others")
    list_transaction = list(df2.TransactionType.unique())
    df2['Airline Sales'] = np.where(df2["TransactionType"] == 'Airline Sales', df2.Value, 0)
    for i in list_transaction:
        df2[i] = np.where(df2["TransactionType"] == i, df2.Value, 0)
    df2["RiyaPNR"] = df2["RiyaPNR"].fillna("No Input")
    df2["AirlinePNR"] = df2["AirlinePNR"].fillna("No Input")
    df2.groupby(["DateTime"]).sum(numeric_only=True)
    df3 = df2.groupby(["DateTime", "Description", "RiyaPNR", "AirlinePNR"]).sum(numeric_only=True)
    df3.drop("Remaining", axis = 1 , inplace = True)
    df3.reset_index(inplace=True) 
    #df3.to_excel("output.xlsx")
    df_b1 = df_b1[df_b1["Ticket Status"] != "TO TICKET"]
    df_b1["Passenger Name Split"] = df_b1["Passenger Name"].str.split(",")
    df_b1["No of PAX"] = df_b1["Passenger Name Split"].str.len()
    try:
        df_b1[["Airline Code", "Flight Number"]] = df_b1["Flight No"].str.split(" ", expand = True)
    except:
        split_string = df_b1["Flight No"].str.split(" ", expand = True)
        df_b1[["Airline Code", "Flight Number"]]  = split_string.iloc[:,:2]
    df_b1.drop("Passenger Name Split", axis = 1, inplace=True)
    df_b2 = df_b1[["Riya PNR", "Passenger Name", "Sector", "Departure Date", "Airline Code", "Airport Id", "No of PAX"]].copy()
    passenger_list = list(df_b2["Passenger Name"].str.split(","))
    passenger_list = list(set([i for name in passenger_list for i in name]))
    condition = True
    while condition and pas_master_file is not None:
        pass_master_list = list(df_passenger_master["lead passenger"])
        passenger_list_not_in_master = []
        for i in passenger_list:
            if i not in pass_master_list:
                passenger_list_not_in_master.append(i)
        df_pass = pd.DataFrame(passenger_list_not_in_master)
        try:
            no_of_missing_pass =  df_pass[df_pass.columns[0]].count()
        except:
            no_of_missing_pass = 0
            
        if no_of_missing_pass > 0:
            st.write("Missing Passneger List")
            st.write(df_pass)
            df_pass.to_excel("missing_passenger.xlsx")
            try:
                with open("missing_passenger.xlsx", "rb") as template_file:
                    template_byte = template_file.read()
                    btn_1 = st.download_button(
                            label="Download Missing Passenger List",
                            data=template_byte,
                            file_name="missing_passenger.xlsx",
                            mime='application/octet-stream'
                            )
            except:
                pass
            cntr_str = cntr
            cntr = cntr + 1
            
            pas_master_file = st.file_uploader("Upload the Updated Riya Passenger Data Master File", type=['csv','xlsx'],accept_multiple_files=False,key="fileUploader_pass_"+str(cntr_str))
            if pas_master_file is not None:
                df_passenger_master = pd.read_excel(pas_master_file)
                df_passenger_master.rename(columns = {'Name':'lead passenger'}, inplace = True)
        else:
            condition = False
        
    if no_of_missing_pass == 0:
        # st.write("No of Missing passenger")
        # st.write(no_of_missing_pass)
        
        
        df_b2.rename(columns = {'Riya PNR':'RiyaPNR'}, inplace = True)
        df_final = pd.merge(df3, df_b2, on='RiyaPNR', how ="left")
        try:
            a= df_final["Airline Reschedule(FARE DIFFERENCE)"] 
        except:
            df_final["Airline Reschedule(FARE DIFFERENCE)"]   = 0
            
        try:
            a= df_final["Airline Other Services"] 
        except:
            df_final["Airline Other Services"]   = 0
        
        try:
            a= df_final["Airline Reschedule(SUPPILER PENALTY)"] 
        except:
            df_final["Airline Reschedule(SUPPILER PENALTY)"]   = 0
            
        try:
            df_final["Airline Reschedule(SUPPILER PENALTY)"] = df_final["Airline Reschedule(SUPPILER PENALTY)"] + df_final["Airline Reschedule(FARE DIFFERENCE)"] + df_final["Airline Other Services"]
        except:
            df_final["Airline Reschedule(SUPPILER PENALTY)"] = df_final["Airline Reschedule(SUPPILER PENALTY)"] + df_final["Airline Reschedule(FARE DIFFERENCE)"]
        
        df_final["Booking Date"] = pd.to_datetime(df_final['DateTime']).dt.date
        df_final["Booking Time"] = pd.to_datetime(df_final['DateTime']).dt.time
        df_final["Travel Date"] = pd.to_datetime(df_final['Departure Date']).dt.date
        
        try:
            a = df_final["Airline Cancellation(SOLD AMOUNT REVERSAL)"]
        except:
            df_final["Airline Cancellation(SOLD AMOUNT REVERSAL)"] = 0
            
        try: 
            a = df_final["PG Online Transfer"]
        except:
            df_final["PG Online Transfer"] = 0
        try: 
            a = df_final["PG Online Transfer Incentive"]
        except:
            df_final["PG Online Transfer Incentive"] = 0
        
        # df_final.loc[df_final["Airline Sales"]<0, "Product Type"] = "Ticket Issued"
        df_final.loc[df_final["Airline Cancellation(SOLD AMOUNT REVERSAL)"]>0, "Product Type"] = "Ticket Cancellation"
        df_final.loc[df_final["PG Online Transfer"]>0, "Product Type"] = "Deposit"
        df_final.loc[df_final["Airline Reschedule(SUPPILER PENALTY)"]<0, "Product Type"] = "Ticket Rescheduled"
        # df_final.loc[df_final["Insurance Sales"]<0, "Product Type"] = "Insurance"
        # df_final.loc[df_final["Seat Selection"]<0, "Product Type"] = "Seat Selection"
        # df_final.loc[df_final["Airline Cancellation(Seat Selection)"]>0, "Product Type"] = "Seat Selection Refund"
        df_final.loc[df_final["PG Online Transfer Incentive"]>0, "Product Type"] = "Deposit Incentive"
        # df_final.loc[df_final["Airline Baggage"]<0, "Product Type"] = "Airline Baggage"
        df_final.loc[df_final["Airline Sales"]<0, "Product Type"] = "Ticket Issued"
        # df_final.loc[df_final["Others"]!=0, "Product Type"] = "Others"
        df_final.drop("DateTime", axis=1, inplace=True)
        df_final.drop("Departure Date", axis=1, inplace=True)
        col_list = list(df_final.columns)
        df_final.drop(['CreditAmount','DebitAmount','Value'], axis=1, inplace=True)
        col_list = ['Booking Date', "Booking Time",'Product Type', 'Airport Id', 'Description','RiyaPNR', 'AirlinePNR','Passenger Name',"No of PAX",'Travel Date', 'Airline Code','Sector', 'Airline Sales','Others','Airline Commission','Airline TDS On Earnings','Service Fee','GST on Service Fee','Airline Cancellation(SOLD AMOUNT REVERSAL)','Airline Cancellation(PENALTY)','Airline Earnings Reversal','PG Online Transfer','Airline TDS Amount Reversal','Offline Adjustment','Airline Reschedule(SUPPILER PENALTY)','Airline Earnings','Insurance Sales','Insurance Commission','Insurance TDS On Earnings','Seat Selection','Airline Cancellation(Seat Selection)','PG Online Transfer Incentive','PG Online Transfer Incentive TDS','Airline Baggage',]
        df_final = df_final.reindex(columns=col_list)
        df_final["total amount charged"] = df_final.iloc[:,12:].sum(axis = 1)
        df_final["Closing balance"]  = df_final["total amount charged"].cumsum() + opening_balance 
        df_master_data = df_final.copy()
        df_master_data  = df_master_data[df_master_data["Product Type"] == "Ticket Issued"]
        col_index= ["Booking Date", "Airport Id", "Description", "RiyaPNR", "AirlinePNR", "Passenger Name", "No of PAX", "Travel Date", "Airline Code", "Sector"]
        df_master_data = df_master_data.reindex(columns=col_index)
        index_to_delete = []
        for i, row in  df_master_data.iterrows():
            if df_existing['RiyaPNR'].eq(row[3]).any():
                index_to_delete.append(i)
        df_master_data.drop(index_to_delete, inplace=True)
        st.write("Riya Master Data")
        st.write(df_existing)
        
        #filename = "Riya_Master_Record.xlsx"
        workbook = load_workbook(master_record_file)
        worksheet = workbook.active
        for i, row in df_master_data.iterrows():
            worksheet.append(list(row))
        workbook.save(master_record_file)
        
        df_master_data_2 = pd.read_excel(master_record_file)
        df_master_data_2.to_excel("Riya_Master_Record.xlsx")
        st.write(df_master_data_2)
        
        
        
        df_final["Passenger Name"] = df_final["Passenger Name"].fillna("Passenger Name Missing")
        for i, row in df_final.iterrows():
            if row[5] != "No Input" and row[7] == "Passenger Name Missing":
                val = row[5]
                print(row[5])
                rows = df_existing.index[df_existing["RiyaPNR"]==val]
                if rows.size !=0:
                    row_data = list(df_existing.iloc[rows[0],:])
                    row_data[7] = pd.to_datetime(row_data[7]).date()
                    print(row_data)
                
                    df_final.iloc[i,3:12] = row_data[1:]
        #             index_to_delete.append(i)
        
        df_final.to_excel("output_1.xlsx")
        df_temp_2 = df_final.copy()
        df_temp_2["Base Amount"] = df_temp_2.fillna(0)['Airline Sales'] + df_temp_2.fillna(0)['Service Fee']+df_temp_2.fillna(0)['GST on Service Fee'] + \
        df_temp_2.fillna(0)['Airline Cancellation(SOLD AMOUNT REVERSAL)']  + df_temp_2.fillna(0)['Insurance Sales']
        
        df_temp_2['Airline Commission'] = df_temp_2.fillna(0)['Airline Commission'] + df_temp_2.fillna(0)['Airline Earnings Reversal']
        df_temp_2['Airline TDS On Earnings'] = df_temp_2.fillna(0)['Airline TDS On Earnings'] + df_temp_2.fillna(0)['Airline TDS Amount Reversal']
    
        df_temp_2["Debit Amount"] = df_temp_2.fillna(0)["PG Online Transfer"] 
    
        df_temp_2["Total Amount"] = df_temp_2.fillna(0)['Base Amount'] + df_temp_2.fillna(0)['Airline Cancellation(PENALTY)']  + \
        df_temp_2.fillna(0)['Airline Commission'] + df_temp_2.fillna(0)['Airline TDS On Earnings'] + df_temp_2.fillna(0)["Debit Amount"]
        df_temp_2["Credit Amount"] = np.where(df_temp_2["Total Amount"]<0,df_temp_2["Total Amount"],0)
        df_temp_2["Debit Amount"] = np.where(df_temp_2["Total Amount"]>0,df_temp_2["Total Amount"],0)
    
    
        col_list = ['Supplier Code','Booking Date', 'Airline Code', 'Sector','Travel Date', 'AirlinePNR','Passenger Name', 'Base Amount', 'Airline Cancellation(PENALTY)','Airline Commission','Airline TDS On Earnings', 'Debit Amount','Credit Amount', "Closing balance", 'RiyaPNR',   'Others','Airline Cancellation(SOLD AMOUNT REVERSAL)','Airline Cancellatcion(PENALTY)','Airline Earnings Reversal','PG Online Transfer','Airline TDS Amount Reversal','Offline Adjustment','Airline Reschedule(SUPPILER PENALTY)','Airline Earnings','Insurance Sales','Insurance Commission','Insurance TDS On Earnings','Seat Selection','Airline Cancellation(Seat Selection)','PG Online Transfer Incentive','PG Online Transfer Incentive TDS','Airline Baggage',"Booking Time",]
        df_temp_2 = df_temp_2.reindex(columns=col_list)
        df_temp_2.to_excel("output_temp_2.xlsx")
        df_temp = df_final.copy()
        df_temp["Supplier Code"] = "RC"
        df_temp["DT Service Fees"] = 200*df_temp["No of PAX"]
        df_temp["Total Service Fees"] = df_temp["DT Service Fees"] - df_temp["Service Fee"]
        
        df_temp["GST Amt"] = df_temp["Total Service Fees"] * 0.18
        df_temp["Net Amt"] = (df_temp["Airline Sales"] + df_temp["Service Fee"] + df_temp["GST on Service Fee"] +  df_temp["Airline Commission"] + df_temp["Airline TDS On Earnings"])*(-1) 
        df_temp["Round Off"] = 0
        df_temp["Invoice Amt"] = - df_temp["Airline Sales"] + df_temp["Total Service Fees"] + df_temp["GST Amt"]
        
        col_list = ['Supplier Code','Booking Date', 'Airline Code', 'Sector','Travel Date', 'AirlinePNR',"No of PAX",'Product Type','Passenger Name', 'Airline Sales','Service Fee','GST on Service Fee','Airline Commission','Airline TDS On Earnings','Total Service Fees','GST Amt', 'Net Amt', 'Round Off', 'Invoice Amt', 'Airport Id', 'Description','RiyaPNR',   'Others','Airline Cancellation(SOLD AMOUNT REVERSAL)','Airline Cancellatcion(PENALTY)','Airline Earnings Reversal','PG Online Transfer','Airline TDS Amount Reversal','Offline Adjustment','Airline Reschedule(SUPPILER PENALTY)','Airline Earnings','Insurance Sales','Insurance Commission','Insurance TDS On Earnings','Seat Selection','Airline Cancellation(Seat Selection)','PG Online Transfer Incentive','PG Online Transfer Incentive TDS','Airline Baggage',"Booking Time",]
        df_temp = df_temp.reindex(columns=col_list)
        df_temp['Airline Sales'] = df_temp['Airline Sales']*(-1)
        df_temp['Service Fee'] = df_temp['Service Fee'] *(-1)
        df_temp['GST on Service Fee'] = df_temp['GST on Service Fee'] * (-1)
        df_temp['Airline TDS On Earnings'] = df_temp['Airline TDS On Earnings']  * (-1)
        df_customer =df_final.copy()
        df_customer = df_customer[df_customer["RiyaPNR"] != "No Input"]
        df_customer["Airline/Insuranance Charges"] = df_customer.fillna(0)['Airline Sales'] + df_customer.fillna(0)['Airline Cancellation(PENALTY)'] + df_customer.fillna(0)['Airline Reschedule(SUPPILER PENALTY)'] + df_customer.fillna(0)['Insurance Sales'] + df_customer.fillna(0)['Seat Selection'] + df_customer.fillna(0)['Airline Baggage']
        # df_customer["Airline/Insuranance Charges"] = -1*df_customer["Airline/Insuranance Charges"]
        df_customer["Refund/Credit"] = df_customer.fillna(0)['Airline Cancellation(SOLD AMOUNT REVERSAL)'] + df_customer.fillna(0)["Airline Cancellation(Seat Selection)"] 
        # df_customer["Refund/Credit"] = -1*df_customer["Refund/Credit"]
            
        df_customer_1 = df_customer[["Booking Date", "Booking Time", "Product Type", "Airport Id", "Description", "RiyaPNR", "AirlinePNR", "Passenger Name", "No of PAX", "Travel Date",'Airline Code', 'Sector', "Airline/Insuranance Charges", "Refund/Credit" ,'Service Fee', 'GST on Service Fee']].copy()
        
        df_customer_1['Airline/Insuranance Charges'] = df_customer_1['Airline/Insuranance Charges'] *(-1)
        df_customer_1['Refund/Credit'] = df_customer_1['Refund/Credit'] *(-1)
        df_customer_1['Service Fee'] = df_customer_1['Service Fee'] *(-1)
        df_customer_1['GST on Service Fee'] = df_customer_1['GST on Service Fee'] *(-1)
        
        
        df_customer_1.rename(columns = {'Service Fee':'Supplier Service Fees'}, inplace = True)
        df_customer_1.rename(columns = {'GST on Service Fee':'GST on Supplier Service Fees'}, inplace = True)
        df_customer_1["DT Service Fees"] = 200*df_customer_1["No of PAX"]
        df_customer_1["Total Service Fees"] = df_customer_1["DT Service Fees"] + df_customer_1["Supplier Service Fees"]
        df_customer_1["CGST/IGST"] = 0
        df_customer_1["CGST"] = 0
        df_customer_1["SGST"] = 0
        df_customer_1["IGST"] = 0
        df_customer_1["Invoice Value"] = df_customer_1["CGST"] + df_customer_1["SGST"] + df_customer_1["IGST"] + df_customer_1["Total Service Fees"]
        df_customer_1["Payable Amount"] =  df_customer_1['Airline/Insuranance Charges'] + df_customer_1["Invoice Value"] + df_customer_1["Refund/Credit"]
        # df_customer_1["B2B/B2C"] = "B2C"
        df_customer_1.to_excel("output_2.xlsx")
        df_customer_dom = df_customer_1[(df_customer_1["Airport Id"] == "Domestic") & (df_customer_1["Product Type"] =="Ticket Issued")]
        df_remaining  = pd.concat([df_customer_1,df_customer_dom]).drop_duplicates(keep=False)
        df_customer_dom["lead passenger"] = df_customer_dom["Passenger Name"].str.split(",")
        df_customer_dom["lead passenger"] = df_customer_dom["lead passenger"].apply(lambda x: x[0])
        df_passenger_master["B2B/B2C"] = "B2C"
        df_passenger_master["B2B/B2C"] = df_passenger_master["B2B/B2C"].where(df_passenger_master["GST Number"].isna(),"B2B")
        df_passenger_master.rename(columns = {'Name':'lead passenger'}, inplace = True)
        df_dom_final = pd.merge(df_customer_dom, df_passenger_master, on='lead passenger', how ="left")
        df_dom_final["CGST/IGST"] = df_dom_final["CGST/IGST"].where(df_dom_final["State"] != "Maharashtra", 1)
        df_dom_final["CGST/IGST"] = df_dom_final["CGST/IGST"].where(df_dom_final["State"] == "Maharashtra", 2)
        df_dom_final["CGST"] = df_dom_final["CGST"].where(df_dom_final["State"] != "Maharashtra", df_dom_final["Total Service Fees"]* 0.09)
        df_dom_final["SGST"] = df_dom_final["SGST"].where(df_dom_final["State"] != "Maharashtra", df_dom_final["Total Service Fees"]* 0.09)
        df_dom_final["IGST"] = df_dom_final["IGST"].where(df_dom_final["State"] == "Maharashtra", df_dom_final["Total Service Fees"]* 0.18)
        df_dom_final["Invoice Value"] = df_dom_final["CGST"] + df_dom_final["SGST"] + df_dom_final["IGST"] + df_dom_final["Total Service Fees"]
        df_dom_final["Payable Amount"] = df_dom_final["Invoice Value"] + df_dom_final["Airline/Insuranance Charges"] + df_dom_final["Refund/Credit"]
        
        col_list_2 = ['Booking Date', 'Invoice to', 'City', 'State', 'GST Number', 'Airline Code', 'Sector', 'Travel Date', 'AirlinePNR', 'Passenger Name', 'No of PAX', "Airline/Insuranance Charges", "Refund/Credit" , 'Supplier Service Fees', 'GST on Service Fee', 'DT Service Fees','Total Service Fees',"CGST/IGST",'CGST', 'SGST', 'IGST', 'Invoice Value','Payable Amount', "B2B/B2C"]
        df_dom_final = df_dom_final.reindex(columns=col_list_2)
        
        df_dom_final["Temp"] = df_dom_final.index+2
        df_dom_final["Temp"] = df_dom_final["Temp"].astype(str)
        df_dom_final["Total Service Fees"] = "=O"+df_dom_final["Temp"]+"+Q"+df_dom_final["Temp"]
        df_dom_final["CGST"] = "=R"+df_dom_final["Temp"]+"*(2-S"+df_dom_final["Temp"]+")*0.09"
        df_dom_final["SGST"] = "=R"+df_dom_final["Temp"]+"*(2-S"+df_dom_final["Temp"]+")*0.09"
        df_dom_final["IGST"] = "=R"+df_dom_final["Temp"]+"*(S"+df_dom_final["Temp"]+"-1)*0.18"
        df_dom_final["Invoice Value"] = "=R"+df_dom_final["Temp"] + "+T"+df_dom_final["Temp"] + "+U"+df_dom_final["Temp"] + "+V"+df_dom_final["Temp"]
        df_dom_final["Payable Amount"] = "=M"+df_dom_final["Temp"] + "+W"+df_dom_final["Temp"] 
        df_dom_final.drop("Temp",axis=1 ,inplace=True)
        df_dom_final.to_excel("domestic_final.xlsx")
        
        df_customer_intl = df_customer_1[(df_customer_1["Airport Id"] == "International") & (df_customer_1["Product Type"] =="Ticket Issued")]
        df_remaining = pd.concat([df_remaining,df_customer_intl]).drop_duplicates(keep=False)
        df_customer_intl["lead passenger"] = df_customer_intl["Passenger Name"].str.split(",")
        df_customer_intl["lead passenger"] = df_customer_intl["lead passenger"].apply(lambda x: x[0])
        df_intl_final = pd.merge(df_customer_intl, df_passenger_master, on='lead passenger', how ="left")
        df_intl_final = df_intl_final.reindex(columns=col_list_2)
        
        df_intl_final["CGST/IGST"] = df_intl_final["CGST/IGST"].where(df_intl_final["State"] != "Maharashtra", 1)
        df_intl_final["CGST/IGST"] = df_intl_final["CGST/IGST"].where(df_intl_final["State"] == "Maharashtra", 2)
        df_intl_final["Temp"] = df_intl_final.index+2
        df_intl_final["Temp"] = df_intl_final["Temp"].astype(str)
        df_intl_final["Total Service Fees"] = "=O"+df_intl_final["Temp"]+"+Q"+df_intl_final["Temp"]
        df_intl_final["CGST"] = "=R"+df_intl_final["Temp"]+"*(2-S"+df_intl_final["Temp"]+")*0.09"
        df_intl_final["SGST"] = "=R"+df_intl_final["Temp"]+"*(2-S"+df_intl_final["Temp"]+")*0.09"
        df_intl_final["IGST"] = "=R"+df_intl_final["Temp"]+"*(S"+df_intl_final["Temp"]+"-1)*0.18"
        df_intl_final["Invoice Value"] = "=R"+df_intl_final["Temp"] + "+T"+df_intl_final["Temp"] + "+U"+df_intl_final["Temp"] + "+V"+df_intl_final["Temp"]
        df_intl_final["Payable Amount"] = "=M"+df_intl_final["Temp"] + "+W"+df_intl_final["Temp"] 
        df_intl_final.drop("Temp",axis=1 ,inplace=True)
        
        df_intl_final.to_excel("international_final.xlsx")
        df_all_flight = df_intl_final.copy(deep=True)
        df_all_flight["Domestic/International"] = "International"
        df_temp_dom = df_dom_final.copy(deep=True)
        df_temp_dom["Domestic/International"] = "Domestic"
        # df_all_flight = df_all_flight.append(df_temp_dom)
        df_all_flight = pd.concat([df_all_flight, df_temp_dom], ignore_index=True)
        df_all_flight = df_all_flight.sort_values(by=['Booking Date'], ascending=True)
        df_all_flight = df_all_flight.reset_index(drop=True)
        
        df_all_flight["Temp"] = df_all_flight.index+2
        df_all_flight["Temp"] = df_all_flight["Temp"].astype(str)
        df_all_flight["Total Service Fees"] = "=O"+df_all_flight["Temp"]+"+Q"+df_all_flight["Temp"]
        df_all_flight["CGST"] = "=R"+df_all_flight["Temp"]+"*(2-S"+df_all_flight["Temp"]+")*0.09"
        df_all_flight["SGST"] = "=R"+df_all_flight["Temp"]+"*(2-S"+df_all_flight["Temp"]+")*0.09"
        df_all_flight["IGST"] = "=R"+df_all_flight["Temp"]+"*(S"+df_all_flight["Temp"]+"-1)*0.18"
        df_all_flight["Invoice Value"] = "=R"+df_all_flight["Temp"] + "+T"+df_all_flight["Temp"] + "+U"+df_all_flight["Temp"] + "+V"+df_all_flight["Temp"]
        df_all_flight["Payable Amount"] = "=M"+df_all_flight["Temp"] + "+W"+df_all_flight["Temp"] 
        df_all_flight.drop("Temp",axis=1 ,inplace=True)
        
        
        
        df_all_flight.to_excel("All_Tickets_final.xlsx")
            
        passenger_list = list(df_customer_intl["Passenger Name"].str.split(","))
        passenger_list = list(set([i for name in passenger_list for i in name]))
        df_customer_ticket_cancellation = df_customer_1[(df_customer_1["Product Type"] =="Ticket Cancellation")]
        df_remaining = pd.concat([df_remaining,df_customer_ticket_cancellation]).drop_duplicates(keep=False)
        df_customer_ticket_cancellation["lead passenger"] = df_customer_ticket_cancellation["Passenger Name"].str.split(",")
        df_customer_ticket_cancellation["lead passenger"] = df_customer_ticket_cancellation["lead passenger"].apply(lambda x: x[0])
        df_cancel_final = pd.merge(df_customer_ticket_cancellation, df_passenger_master, on='lead passenger', how ="left")
        df_cancel_final = df_cancel_final.reindex(columns=col_list_2)
        df_cancel_final["CGST"] = df_cancel_final["CGST"].where(df_cancel_final["State"] != "Maharashtra", df_cancel_final["Total Service Fees"]* 0.09)
        df_cancel_final["SGST"] = df_cancel_final["SGST"].where(df_cancel_final["State"] != "Maharashtra", df_cancel_final["Total Service Fees"]* 0.09)
        df_cancel_final["IGST"] = df_cancel_final["IGST"].where(df_cancel_final["State"] == "Maharashtra", df_cancel_final["Total Service Fees"]* 0.18)
        df_cancel_final["Invoice Value"] = df_cancel_final["CGST"] + df_cancel_final["SGST"] + df_cancel_final["IGST"] + df_cancel_final["Total Service Fees"]
        df_cancel_final["Payable Amount"] = df_cancel_final["Invoice Value"] + df_cancel_final["Airline/Insuranance Charges"] + df_cancel_final["Refund/Credit"]
        
        df_cancel_final.to_excel("cancellation_final.xlsx")
        
        df_customer_ticket_reschedule = df_customer_1[(df_customer_1["Product Type"] =="Ticket Rescheduled")]
        df_remaining = pd.concat([df_remaining,df_customer_ticket_reschedule]).drop_duplicates(keep=False)
        df_customer_ticket_reschedule["lead passenger"] = df_customer_ticket_reschedule["Passenger Name"].str.split(",")
        df_customer_ticket_reschedule["lead passenger"] = df_customer_ticket_reschedule["lead passenger"].apply(lambda x: x[0])
        df_reschedule_final = pd.merge(df_customer_ticket_reschedule, df_passenger_master, on='lead passenger', how ="left")
        df_reschedule_final = df_reschedule_final.reindex(columns=col_list_2)
        df_reschedule_final["CGST"] = df_reschedule_final["CGST"].where(df_reschedule_final["State"] != "Maharashtra", df_reschedule_final["Total Service Fees"]* 0.09)
        df_reschedule_final["SGST"] = df_reschedule_final["SGST"].where(df_reschedule_final["State"] != "Maharashtra", df_reschedule_final["Total Service Fees"]* 0.09)
        df_reschedule_final["IGST"] = df_reschedule_final["IGST"].where(df_reschedule_final["State"] == "Maharashtra", df_reschedule_final["Total Service Fees"]* 0.18)
        df_reschedule_final["Invoice Value"] = df_reschedule_final["CGST"] + df_reschedule_final["SGST"] + df_reschedule_final["IGST"] + df_reschedule_final["Total Service Fees"]
        df_reschedule_final["Payable Amount"] = df_reschedule_final["Invoice Value"] + df_reschedule_final["Airline/Insuranance Charges"] + df_reschedule_final["Refund/Credit"]
        
        df_reschedule_final.to_excel("rescheduling_final.xlsx")
            
        
        df_customer_insurance= df_customer_1[(df_customer_1["Product Type"] =="Insurance")]
        df_remaining = pd.concat([df_remaining,df_customer_insurance]).drop_duplicates(keep=False)
        df_customer_insurance["lead passenger"] = df_customer_insurance["Passenger Name"].str.split(",")
        df_customer_insurance["lead passenger"] = df_customer_insurance["lead passenger"].apply(lambda x: x[0])
        df_insurance_final = pd.merge(df_customer_insurance, df_passenger_master, on='lead passenger', how ="left")
        df_insurance_final = df_insurance_final.reindex(columns=col_list_2)
        df_insurance_final["CGST"] = df_insurance_final["CGST"].where(df_insurance_final["State"] != "Maharashtra", df_insurance_final["Total Service Fees"]* 0.09)
        df_insurance_final["SGST"] = df_insurance_final["SGST"].where(df_insurance_final["State"] != "Maharashtra", df_insurance_final["Total Service Fees"]* 0.09)
        df_insurance_final["IGST"] = df_insurance_final["IGST"].where(df_insurance_final["State"] == "Maharashtra", df_insurance_final["Total Service Fees"]* 0.18)
        df_insurance_final["Invoice Value"] = df_insurance_final["CGST"] + df_insurance_final["SGST"] + df_insurance_final["IGST"] + df_insurance_final["Total Service Fees"]
        df_insurance_final["Payable Amount"] = df_insurance_final["Invoice Value"] + df_insurance_final["Airline/Insuranance Charges"] + df_insurance_final["Refund/Credit"]
        
        df_insurance_final.to_excel("insurance_final.xlsx")
        
        df_customer_seat_selection= df_customer_1[(df_customer_1["Product Type"] =="Seat Selection") | (df_customer_1["Product Type"] =="Seat Selection Refund")]
        df_remaining = pd.concat([df_remaining,df_customer_seat_selection]).drop_duplicates(keep=False)
        df_customer_seat_selection["lead passenger"] = df_customer_seat_selection["Passenger Name"].str.split(",")
        df_customer_seat_selection["lead passenger"] = df_customer_seat_selection["lead passenger"].apply(lambda x: x[0])
        df_seat_selection_final = pd.merge(df_customer_seat_selection, df_passenger_master, on='lead passenger', how ="left")
        df_seat_selection_final = df_seat_selection_final.reindex(columns=col_list_2)
        df_seat_selection_final["CGST"] = df_seat_selection_final["CGST"].where(df_seat_selection_final["State"] != "Maharashtra", df_seat_selection_final["Total Service Fees"]* 0.09)
        df_seat_selection_final["SGST"] = df_seat_selection_final["SGST"].where(df_seat_selection_final["State"] != "Maharashtra", df_seat_selection_final["Total Service Fees"]* 0.09)
        df_seat_selection_final["IGST"] = df_seat_selection_final["IGST"].where(df_seat_selection_final["State"] == "Maharashtra", df_seat_selection_final["Total Service Fees"]* 0.18)
        df_seat_selection_final["Invoice Value"] = df_seat_selection_final["CGST"] + df_seat_selection_final["SGST"] + df_seat_selection_final["IGST"] + df_seat_selection_final["Total Service Fees"]
        df_seat_selection_final["Payable Amount"] = df_seat_selection_final["Invoice Value"] + df_seat_selection_final["Airline/Insuranance Charges"] + df_seat_selection_final["Refund/Credit"]
        
        df_seat_selection_final.to_excel("seat_final.xlsx")
        
        final_output_file_list = ["output_2.xlsx", "All_Tickets_final.xlsx", "domestic_final.xlsx", "international_final.xlsx", "All_Tickets_final.xlsx", "cancellation_final.xlsx", 
                                  "rescheduling_final.xlsx", "insurance_final.xlsx",  "seat_final.xlsx", "Riya_Master_Record.xlsx"]
        
        
        zip_path = "final_output.zip"
        
        with ZipFile(zip_path, 'w', ZIP_DEFLATED) as zip:
            for file in final_output_file_list:
                zip.write(file, arcname=file)
                
                
        with open("final_output.zip", "rb") as fp:
            btn = st.download_button(
                label="Download ZIP",
                data=fp,
                file_name="final_output.zip",
                mime="application/zip"
            )
        
        # try:
        #     with open("output_2.xlsx", "rb") as template_file:
        #         template_byte = template_file.read()
        #         btn_1 = st.download_button(
        #                 label="Download Output File",
        #                 data=template_byte,
        #                 file_name="output_2.xlsx",
        #                 mime='application/octet-stream'
        #                 )
        # except:
        #     pass
    