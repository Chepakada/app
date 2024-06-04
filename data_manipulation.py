import polars as pl
from datetime import datetime as dt
df = pl.read_csv("C:/Users/19518/Downloads/taylor_data.csv")

fixed = df\
    .drop(["Days since order received"])\
    .with_columns([
        pl.col("Date").str.strptime(pl.Date, "%m/%d/%y").alias("Date"),
        pl.lit(dt.now().date()).alias("Todays Date"),
        pl.col("To Ship On").str.strptime(pl.Date, "%m/%d/%y").alias("To Ship On")])\
    .with_columns([(pl.col("Todays Date") - pl.col("Date")).dt.days().alias("Days Since Order Received"),
                (pl.col("Todays Date") - pl.col("To Ship On")).dt.days().alias("Lead Time")])\
    .select(['Date', 'Todays Date', 'To Ship On', 'Days Since Order Received',
             'Lead Time', 'Document Number', 'PO Number', 'Customer', 'Status',
            'Open Amount'])

print(fixed)