import pandas as pd
import numpy as np
import os, sys



def load_data(filename, foldername='raw_purchases'):
    # Loads excel file to analyze.

    filepath = os.path.join(os.getcwd(), foldername, filename)
    return pd.read_excel(filepath, engine='openpyxl')



def remove_CDN(x):
    # Function to detect if the tag "CDN$" exists in the total cost of the dataframe.

    if (type(x) == str) and (x[0] == 'C'):
        x = "".join(x.split())  # remove all spaces
        return x[4:]
    else:
        return x



def preprocess(df):
    # General cleaning of the data.

    # clean data
    df = df[~pd.isna(df.total) & (
                df.total != 'pending')]  # some items might have NaN because they were grouped with other items
    # df = df[:-1]                 # drop last row (useless info)

    # change data types appropriately
    df.total = df.total.apply(remove_CDN).astype('float64')  # convert total price into numerical
    df.refund = df.refund.apply(remove_CDN).astype('float64')  # find refund values
    df.gift = df.gift.apply(remove_CDN).astype('float64')  # find giftcard values
    df.date = pd.to_datetime(df.date)  # convert to date

    # delete duplicates
    df = df.drop_duplicates(subset=['order id'])

    return df



def filter_by_date(df, startdate, enddate):
    # Crops the data by the specified dates.

    df_cropped = df[(df.date >= startdate) & (df.date <= enddate)].groupby(['to']).sum()
    df_cropped['net cost'] = df_cropped.total - df_cropped.refund  # - df_results.gift

    return df_cropped



def get_final_amount(df):
    return df[["total", "refund", "net cost"]]



def save_totals(df, startdate, enddate, savefolder='results'):
    filename = str(startdate) + '-' + str(enddate) + '_TOTALS.xlsx'
    df.to_excel(os.path.join(os.getcwd(), savefolder, filename))



def save_user_receipts(df, startdate, enddate, savefolder='results\\receipts'):
    # Saves the full list of purchases for each unique name.

    namelist = df['to'].unique()

    for name in namelist:
        filename = str(startdate) + '-' + str(enddate) + '_' + str(name) + '_TOTALS.xlsx'
        df[df.to == name].to_excel(os.path.join(os.getcwd(), savefolder, filename))




def main():
    # Arguments: datafile, startdate, enddate, [True/False for saving user receipts]

    # get arguments
    filename, startdate, enddate = sys.argv[1:-1]

    # process data
    df = load_data(filename)
    df = preprocess(df)
    df_cropped = filter_by_date(df, startdate, enddate)
    df_totals = get_final_amount(df_cropped)
    save_totals(df_totals, startdate, enddate)

    if bool(sys.argv[-1]) == True:
        save_user_receipts(df, startdate, enddate)



if __name__ == '__main__':
    main()










