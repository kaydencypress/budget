import csv
import string
from datetime import datetime
import openpyxl
from openpyxl.chart import BarChart,PieChart,Reference

directory = "/home/dev_iant/workspace/github.com/kaydencypress/budget/"
credit_csv = directory + "chase_export.csv"
category_map_csv = directory + "category_map.csv"
outfile_csv = directory + "export.csv"
outfile_xlsx = directory + "export.xlsx"

class Transaction:
    def __init__(self,date,description,amount,type,category):
        self.date = date
        self.description = description
        self.amount = amount
        self.type = type
        self.category = category

def import_credit_csv(filepath):
    transactions = []
    try:
        with open(filepath) as f:
            csv_reader = csv.reader(f)
            row_count = 0
            for row in csv_reader:
                if row_count != 0:
                    date = row[0]
                    description = row[2]
                    type = row[4] 
                    amount = float(row[5])
                    category = get_category(description)
                    transactions.append(Transaction(date,description,amount,type,category))
                row_count+=1
    except Exception as e:
        print(e)
    return transactions

def read_category_map(filepath):
    category_map = {}
    try:
        with open(filepath) as f:
            csv_reader = csv.reader(f)
            for row in csv_reader:
                description = row[0]
                category = row[1]
                category_map[description] = category
    except Exception as e:
        print(e)
    return category_map

def get_category(description):
    category_map = read_category_map(category_map_csv)
    detail = None
    if "*" in description:
        description = description.split("*")
        description = description[0]
        detail = description[1]
    if "#" in description:
        description = description.split("#")[0]
    description = description.rstrip(string.digits)
    description = description.rstrip()

    if detail and detail in category_map:
        return category_map[detail]
    if description in category_map:
        return category_map[description]
    return "Other"

def calc_monthly_totals(transactions):
    totals = {}
    for transaction in transactions:
        if transaction.category != "Payment":
            transaction_month = datetime.strptime(transaction.date,"%m/%d/%Y").strftime("%m-%Y")
            if transaction_month not in totals:
                totals[transaction_month] = {}
                totals[transaction_month]["Total"] = 0
            if transaction.category not in totals[transaction_month]:
                totals[transaction_month][transaction.category] = 0
            totals[transaction_month][transaction.category] += transaction.amount
            totals[transaction_month]["Total"] += transaction.amount
    return totals

def report_monthly_totals(totals):
    for month in totals:
        print(f"======= {month} =======")
        for category in totals[month]:
            print(f"{category}: {format(totals[month][category],'.2f')}")
    return

def report_transaction_details(transactions):
    for transaction in transactions:
        print(f"Transaction on {transaction.date} for {transaction.amount}: {transaction.description} [{transaction.category}]")
    return

def export_totals_csv(totals,filepath):
    with open(filepath,'w+') as f:
        for month in totals:
            for category in totals[month]:
                f.write(f"{month},{category},{format(totals[month][category],'.2f')}\n")
        #f.write(f"{transactions.date},{transactions.amount},{transactions.description},{transactions.category}")
    return

def export_excel(totals,transactions,filepath):
    wb = openpyxl.Workbook()
    for current_month in totals:
        # add data for total spending by category
        summary_title = current_month + " Summary"
        summary_sheet = wb.create_sheet(summary_title)
        header = ["Category","Amount"]
        summary_sheet.append(header)
        for category in totals[current_month]:
            row = [category,totals[current_month][category]]
            summary_sheet.append(row)

        # create pie chart for total spending by category
        # excludes headers (row 1) and total (row 2)
        chart = PieChart()
        data_range = Reference(summary_sheet,min_col=2,min_row=3,max_col=2,max_row=len(totals[current_month])+2)
        label_range = Reference(summary_sheet,min_col=1,min_row=3,max_col=1,max_row=len(totals[current_month])+2)
        chart.add_data(data_range,titles_from_data=True)
        chart.set_categories(label_range)
        chart.title = "Expenses by Category"
        summary_sheet.add_chart(chart)

        # add data for individual transaction details
        detail_title = current_month + " Transactions"
        detail_sheet = wb.create_sheet(detail_title)
        header = ["Date","Amount","Description","Category","Type"]
        detail_sheet.append(header)
        for transaction in transactions:
            transaction_month = datetime.strptime(transaction.date,"%m/%d/%Y").strftime("%m-%Y")
            if transaction_month == current_month:
                row = [transaction.date,transaction.amount,transaction.description,transaction.category,transaction.type]
                detail_sheet.append(row)

    wb.save(filepath)
    return

def main():
    credit_transactions = import_credit_csv(credit_csv)
    totals = calc_monthly_totals(credit_transactions)
    export_excel(totals,credit_transactions,outfile_xlsx)

main()