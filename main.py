"""
Description:
    Read csv file and create excel sheet after processing.

Usage:
    python main.py file_path

Paramerts:
    file_path = path of csv file
"""

import sys
import os
from datetime import datetime
import pandas as pd


def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)


def get_sales_csv():
    if len(sys.argv) == 1:
        print("Please provide file path.")
        exit()

    file_path = sys.argv[1]
    if not os.path.exists(file_path):
        print("File path is not correct.")
        exit()

    return file_path


def create_orders_dir(sales_csv):
    directory = os.path.dirname(sales_csv)
    now = datetime.now()
    sub_dir_name = f"Orders_{now.year}-{now.month}-{now.day}"
    sub_dir_path = os.path.join(directory, sub_dir_name)
    if not os.path.exists(sub_dir_path):
        os.mkdir(sub_dir_path)
    return sub_dir_path


def process_sales_data(sales_csv, orders_dir):
    df = pd.read_csv(sales_csv)
    df["TOTAL PRICE"] = df["ITEM PRICE"] * df["ITEM QUANTITY"]

    df = df[
        [
            "ORDER ID",
            "ORDER DATE",
            "ITEM NUMBER",
            "PRODUCT LINE",
            "PRODUCT CODE",
            "ITEM QUANTITY",
            "ITEM PRICE",
            "TOTAL PRICE",
            "STATUS",
            "CUSTOMER NAME"
        ]
    ]
    
    df = df.sort_values(by=["ITEM NUMBER"])
    grouped_df = df.groupby("ORDER ID")

    writer = pd.ExcelWriter(orders_dir + '/orders.xlsx', engine='xlsxwriter')

    workbook = writer.book
    money_fmt = workbook.add_format({"num_format": "$#,###.##", "align": "center"})
    align_fmt = workbook.add_format({"align": "center"})

    for index, _grp_df in grouped_df:
        grand_price = round(_grp_df["TOTAL PRICE"].sum(), 2)
        _grp_df = _grp_df.drop(["ORDER ID"], axis=1)
        _grp_df.loc[len(_grp_df.index)] = ["", "", "", "", "", "GRAND TOTAL", f"${grand_price:.2f}", "", ""]

        _grp_df.to_excel(writer, sheet_name=str(index), index=False)
        worksheet = writer.sheets[str(index)]
        worksheet.set_column("A:I", 12, align_fmt)
        worksheet.set_column("F:G", 12, money_fmt)
        worksheet.autofit()

    writer.close()
    return

if __name__ == '__main__':
    main()
    