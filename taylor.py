import pandas as pd
import datetime as dt

df = pd.read_csv("C:/Users/19518/Downloads/taylor_data.csv")

df["Date"] = pd.to_datetime(df["Date"], format="%m/%d/%y").dt.date
df["Todays Date"] = dt.datetime.now().date()
df["To Ship On"] = pd.to_datetime(df["To Ship On"], format="%m/%d/%y").dt.date

df["Days Since Order Received"] = (df["Todays Date"] - df["Date"]).dt.days
df["Lead Time"] = (df["Todays Date"] - df["To Ship On"]).dt.days

df = df[['Date', 'Todays Date', 'To Ship On', 'Days Since Order Received',
         'Lead Time', 'Document Number', 'PO Number', 'Customer', 'Status', 'Open Amount']]

