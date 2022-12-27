

# from fileinput import filename
import datetime
import calendar as cal
import re
import sys

import pandas as pd
from pandas import DataFrame

# import os
# !/usr/bin/env python3

    # read all the lines in file
    # separate the ones that are incomes
    # separate the ones that are FPI, SO, TFR, BGC PAyment
        # FPI(Faster Payments Inwards) : usually meetings offerings
        # SO Standing order
        # TFR: Transfer between account
        # BGC: Cheque or cash deposit
    # Extract the FPI ones as form answers, name, transaction date, purpose, etc
    # Allow user to select the correct one to save to CQ
    # All user to save this info into spreadsheet containing original input plus saved CQ details, marking the saved ones as saved

    #####
        # After breaking down description for income
            # Separate and mark offering and tithe
            # Seprate other designated offering
                # Kunle
                # IDP
                # Honorarioum
            # Use meeting dates to guess the remaining
    #####

    ## March 3
        # March bank accout name to Name in CQ
            # Extract Bank Account First Name Last Name/Initials
            # Extract CQ Last Name Fist Name/Initials and Title
            # For each name in Bank Statements find the matching name and number in CQ sheet and save it in spreadsheet with fields
            # CQF fielsds
            # "NUMBER","TITLE","FULL_NAME","DATE_RECEIVED","PURPOSE_ITEM (Must Look up Sheet2)","AMOUNT (£)","BRIEF_DESCRIPTION","PAYMENT_TYPE (eg: cheque, electronic transfer, standing order etc)","GIFT_AID? (Yes/No)","DESIGNATED? (Yes/No)""
            #                   "Name"      "Date"          "Purpose"   "Amount"    "Description"   "Payment Type ie Transaction type", "Gift Aid" "Designated"
            



ColumnNames = ["Transaction Date", "Transaction Type", "Transaction Description", "Debit Amount", "Credit Amount", "Balance"]
DecemberRetreatDays = ["11DEC21"]
WatchNightServiceDates = ["01JAN21", "31DEC21"]
CQ_INCOME_TEMPLATE = []
cq_name_col = ""
bank_name_col = 'BANK_NAME'
MINIMUM_NAME_LENGTH = 2

def load_cq_template() -> type[DataFrame]:
    pd.set_option('display.max_columns', None)
    income_template_file = "C:\\Users\\ukyade\\Downloads\\OFNC\\inc_upload_manchester_template.xlsx"
    print("Loading CQ file from ", income_template_file)
    cq_inc_df = pd.read_excel(income_template_file, index_col=0, skiprows=4)
    # print(cq_inc_df.head(5))
    CQ_INCOME_TEMPLATE = cq_inc_df.columns
    return cq_inc_df


# Get the cq_name of a name from bank statement
# if name found then mark it as used in the CQ list, so we don't link it again
def get_cq_name(bank_name):
    new_df = pd.DataFrame(columns=['Bank Name', 'CQ_Name'])
    cq_df = load_cq_template()
    global cq_name_col
    cq_name_col = cq_df.columns[1]
    # bank_name_col = 'BANK_NAME'
    cq_df[bank_name_col] = "none"
    cq_names = cq_df[cq_name_col]
    for name in bank_name:
        if isinstance(name, str):
            # print(cq_df.columns)
            if name_in_cq(cq_df, name, name, True):
                pass
            else:
                # Now lets try and turn them round if its initials and a last name
                names = name.split()
                if len(names) == 2:
                    if len(names[0]) == 1 or len(names[1]) == 1:
                        print("No matching values for: ", name, " Trying Last Name")
                        # Now lets try and use last name only if you can find it
                        # split name by space
                        test_last_name = max(names,
                                             key=len)  # find which one is longest, assuming one of them is initial
                        if len(test_last_name) > MINIMUM_NAME_LENGTH and name_in_cq(cq_df, test_last_name, name, False, "**"):
                            pass
                        else:
                            print("No matching values for: ", name, " using last name:", test_last_name)
                    else:
                        # No initial, try guessing last name and initial
                        test_last_name_with_initial = names[0] + " " + names[1][0]
                        if name_in_cq(cq_df, test_last_name_with_initial, name, True, "+++"):
                            print("Name found by guessing initial as:", test_last_name_with_initial)
                        else:
                            test_last_name_with_initial = names[1] + " " + names[0][0]
                            if name_in_cq(cq_df, test_last_name_with_initial, name, True, "==="):
                                print("Name found second try at guessing initial as:", test_last_name_with_initial)
                            else:
                                # This is really pushing it now just find any of the names
                                pass



                else:
                    print("Bank name length is more than 2, length is:", len(names))




    print(cq_df.loc[(cq_df[bank_name_col] != "none"), (cq_name_col, bank_name_col)])

    # print(type(cq_names), type(cq_df))
    return new_df



# Check if name can be found in dataframe list from CQ cq_df
# cq_df
# name: Name to search from
# bank_name: Original version of name as retrieved from bank database
# try_reorder: Try to search for name by changing the order of first and last if not found the first time
# suffix: Append the name with this value in the final list
def name_in_cq(cq_df, name, bank_name, try_reorder = False, suffix =""):
    try:
        out_df = cq_df.loc[cq_df[cq_name_col].str.contains(name) & (cq_df[bank_name_col] == "none")]
    except re.error:
        return False

    full_name = name
    if out_df.shape[0] == 0 and try_reorder:
        print("No matching values for: ", name, " Trying reordering")
        re_ordered, full_name = try_reorder_name(full_name)
        if re_ordered:
            try:
                out_df = cq_df.loc[cq_df[cq_name_col].str.contains(full_name) & (cq_df[bank_name_col] == "none")]
            except re.error:
                return False

            if(out_df.shape[0] > 0):
                print("+++ Found Name by re-ordering", full_name)


    if out_df.shape[0] > 0:
        cq_df.at[out_df.index[0], bank_name_col] = bank_name + " " + suffix
        # print(cq_df.at[out_df.index[0], bank_name_col])
        if out_df.shape[0] > 1:
            print(name, "has ", len(out_df.shape), " matches")

        return True

    return False


def try_reorder_name(full_name):
    names = full_name.split()
    if len(names) == 2:
        full_name = names[1] + " " + names[0]
        return True, full_name

    return False, full_name

def get_name_from_cq(bank_name):
    new_df = pd.DataFrame(columns=['Bank Name', 'CQ_Name'])
    cq_df = load_cq_template()
    cq_names = cq_df[cq_df.columns[1]]
    for name in bank_name:
        for index, cq_name in cq_names.items():
            if isinstance(name, str) and name in cq_name:
                print(name, " CQ is ", cq_name)
        # print(name)

    print(type(cq_names), type(cq_df))
    return new_df


def main():
    cq_income_df = load_cq_template()
    full_name_list = cq_income_df[cq_income_df.columns[1]]
    # print(full_name_list.to_list())
    branchMeetingDates = getMeetingDate(2021,cal.SUNDAY,1) + DecemberRetreatDays + WatchNightServiceDates

    #print(branchMeetingDates)
    #return
    accountFile = "C:\\Users\\ukyade\\Downloads\\OFNC\\OFNC.csv"

    accountDf = pd.read_csv(accountFile, sep=",", skiprows = 0)

    #print(accountDf.head())
    FPIDataFrame = accountDf[accountDf['Transaction Type'] == 'FPI']
    #timeDF = pd.DataFrame(columns = ['Bank Description', 'Time'] )

    # OGUNMODIMU O TITHE DONATION RP4659985468854500 206412     10 04JAN21 08:36
    # split to 0: OGUNMODIMU O TITHE DONATION RP4659985468854500 206412 10 04JAN21 | 08:36
    timeDF = FPIDataFrame['Transaction Description'].str.rsplit(n=1, expand=True)

    #timeDF['Transaction Description', "Transaction Time"] = descDF.str.split(n=1)
    #accountDf['Time'] = timeDF[1]

    # split to OGUNMODIMU O TITHE DONATION RP4659985468854500 206412 10 | 04JAN21  
    accountDf['Real Date'] = timeDF[0].str.rsplit(n=1, expand=True)[1]
    # print(accountDf.head(6))
    # exit()
    #  OGUNMODIMU | O | TITHE DONATION RP4659985468854500 206412 10 
    nameDF = timeDF[0].str.split(n=2, expand=True) 
    accountDf['Name'] = nameDF[0] + " " + nameDF[1]

    # split to OGUNMODIMU O TITHE DONATION | RP4659985468854500 | 206412| 10 | 04JAN21 
    # Second split OGUNMODIMU | O | TITHE DONATION
    accountDf['Bank Description'] = timeDF[0].str.rsplit(n=4, expand=True)[0].str.split(n=2, expand=True)[2]

    # Insert empty column at the last column
    accountDf.insert(len(accountDf.columns), 'Purpose',"")

    # Do for standing orders
    soDF = accountDf[accountDf['Transaction Type'] == 'SO']
    splitDesc = soDF['Transaction Description'].str.split(n=2, expand=True)
    accountDf.loc[soDF.index, 'Name'] = splitDesc[0] + " " + splitDesc[1]
    accountDf.loc[soDF.index, 'Bank Description'] = splitDesc[2]
    #accountDf.loc[soDF.index, 'Real Date'] = soDF['Date']

    #accountDf.loc[(soDF[ soDF['Bank Description'].str.contains("OFFERING|TITHE") == True]).index, 'Purpose'] = "Tithes and Offering"

    # Get all tithe and offering
        #Get all income
    incomeDF = accountDf[accountDf['Credit Amount'].notnull()]
    accountDf.loc[(incomeDF[incomeDF['Bank Description'].str.contains("OFFERING|TITHE") == True]).index, 'Purpose'] = "Tithes and Offering"
    FPIDataFrame = incomeDF[(incomeDF['Transaction Type'] == 'FPI') & (incomeDF['Bank Description'].str.contains("OFFERING|TITHE") == True)]
    descDF = FPIDataFrame['Real Date'].isin(branchMeetingDates)
    accountDf.loc[descDF[descDF == True].index, 'Description'] = "MANCHESTER BRANCH MEETINGS"

    accountDf.loc[(incomeDF[incomeDF['Transaction Description'].str.contains("STWDSHP") == True]).index, ['Purpose', 'Description']] =  ["Tithes and Offering", "Stewardship Tithes and Offering"]
    accountDf.loc[(incomeDF[incomeDF['Transaction Type'] == 'DEP']).index, ['Purpose', 'Description']] = ["Tithes and Offering", "Tithes and Offering Desc"]



    FPIDataFrame = accountDf[(accountDf['Transaction Type'] == 'FPI') & (accountDf['Purpose'] == '')]
    #print((purposeDF[ purposeDF['Bank Description'].str.contains("KUNLE|IDP") == True]).index.values)
    accountDf.loc[(FPIDataFrame[ FPIDataFrame['Bank Description'].str.contains("IDC|IDP|HONORARIUM|HONOURARIUM|BENEVOLENCE") == True]).index, 'Purpose'] = "Benevolence"
    #IDP Xmas Gift
    accountDf.loc[(FPIDataFrame[ FPIDataFrame['Bank Description'].str.contains("IDP") == True]).index, 'Description'] = "IDP Xmas Gift"
    accountDf.loc[(FPIDataFrame[ FPIDataFrame['Bank Description'].str.contains("POSTAGE | SHIPPING") == True]).index, ['Purpose', 'Description']] = ["Benevolence", "IDP Donation Shipping"]
    accountDf.loc[(FPIDataFrame[ FPIDataFrame['Bank Description'].str.contains("HONORARIUM|HONOURARIUM") == True]).index, 'Description'] = "Donation for Honourarium"
    accountDf.loc[(FPIDataFrame[ FPIDataFrame['Bank Description'].str.contains("BENEVOLENCE") == True]).index, 'Description'] = "Benevolence"

    FPIDataFrame = accountDf[(accountDf['Transaction Type'] == 'FPI') & (accountDf['Purpose'] == '')]
    accountDf.loc[(FPIDataFrame[ FPIDataFrame['Bank Description'].str.contains("KUNLE|BRO\ K") == True]).index, ['Purpose', 'Description']] = ["Welfare", "Benevolence & Welfare"]

    #Refund
    accountDf.loc[(incomeDF[ (incomeDF['Purpose'] == '') & (incomeDF['Transaction Type'] == 'TFR') &
                     (incomeDF['Transaction Description'].str.contains("REFUND") == True)]).index, ['Purpose', 'Description']] = ["Top House Refund", "Top House Refund"]

    # Get all data without Purpose deposited on Meeting days
    noPurposeDF = accountDf[(accountDf['Transaction Type'] == 'FPI') & (accountDf['Purpose'] == '')]
    noPurposeDF = noPurposeDF['Real Date'].isin(branchMeetingDates)
    accountDf.loc[noPurposeDF[noPurposeDF == True].index, ['Purpose', 'Description']] = ["Tithes and Offering", "MANCHESTER BRANCH MEETINGS"]
    # accountDf.loc[noPurposeDF[noPurposeDF == True].index, 'Description'] = "MANCHESTER BRANCH MEETINGS"

    # Do Watch night service
    noPurposeDF = accountDf[(accountDf['Transaction Type'] == 'FPI')]
    accountDf.loc[noPurposeDF.loc[noPurposeDF['Real Date'] == WatchNightServiceDates[-2]].index, 'Description'] = "Watchnight Service 2021"
    accountDf.loc[noPurposeDF.loc[noPurposeDF['Real Date'] == WatchNightServiceDates[-1]].index, 'Description'] = "Watchnight Service 2022"

    # Just fill the rest as Tithes and Offering
    noPurposeDF =  accountDf[(accountDf['Transaction Type'] == 'FPI') & (accountDf['Purpose'] == '')]
    accountDf.loc[noPurposeDF.index, ['Purpose', 'Description', 'Please Check']] = ["Tithes and Offering", "Tithes and Offering", "Yes"]
    # accountDf.loc[noPurposeDF.index, 'Purpose'] = "Tithes and Offering"

    # Description fields
    noPurposeDF = accountDf[(accountDf['Transaction Type'] == 'FPI')]
    accountDf.loc[accountDf[
                                ((accountDf['Bank Description'].str.contains('RETREAT') == True) |
                                (noPurposeDF['Real Date'].str == DecemberRetreatDays[0]))].index, 'Description'] = "December Retreat Offering"
    # print(accountDf.loc[accountDf['Description'].isnull() & (accountDf['Purpose'] == "Tithes and Offering")])
    # exit()
    accountDf.loc[accountDf[(accountDf['Description'].isnull()) & (accountDf['Purpose'] == "Tithes and Offering")].index, 'Description'] = "Tithes and Offering"

    #accountDf.to_csv(os.path.splitext(accountFile)[0] + "OFNCExpand5.csv")
    group_df = accountDf.groupby(['Name'])
    name_list = accountDf['Name'].unique()
    # print(name_list)
    get_cq_name(name_list)
    exit()

    purposeDF['Purpose'] = purposeDF[ purposeDF['Bank Description'].str.contains("KUNLE|IDP") == True]['Bank Description']

    purposeDF['Purpose'] = purposeDF.where(pd.isnull(purposeDF['Purpose']), "Benevolence")['Purpose']
    purposeDF['Purpose'] = purposeDF.where(purposeDF['Bank Description'].str.contains("OFFERING|TITHE") == False ,
                                        "Tithes and Offering")['Purpose']

    # Do Standing order
    soDF = accountDf[accountDf['Transaction Type'] == 'SO']
    #soDF['Purpose'] = soDF.where(soDF['Transaction Type'].str.contains("OFFERING|TITHE") == False,"Tithes and Offering")['Purpose']

    print(soDF.index.values)#[purposeDF['Purpose'].isnull()])

    accountDf['Purpose'] = purposeDF['Purpose']

    #accountDf.to_csv(os.path.splitext(accountFile)[0] + "OFNCExpand.csv")
    #print(descDF['Transaction Description'].split()[-1])

    # Split Transaction Description

    return

def getMeetingDate(year, dayNum, frequency) :
    #print(datetime.date(year,1, 1))
    #print(dayNum)

    dateList = []
    for i in range(1,13) :
        n = 0
        index = 0
        while n < frequency:
            if(cal.monthcalendar(year, i)[1][dayNum]) :
                n += 1

            index += 1


        sundayDate = datetime.date(year,i,cal.monthcalendar(year, i)[index][dayNum]).strftime("%d%b%y").upper()
        dateList.append(sundayDate)

    #print(dateList)
    return dateList


if __name__ == '__main__':
    main()


    # Bank name format if one letter it is initial, otherwise it is last name
    # OFNC NAme if it has initial set initial, else it is last name first name