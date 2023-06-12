import csv
import string
import os
from datetime import datetime
import openpyxl
from openpyxl.chart import BarChart,ProjectedPieChart,Reference
from openpyxl.chart.label import DataLabelList
import re

dir = "/home/dev_iant/workspace/github.com/kaydencypress/budget/"
import_dir = dir + "import/"
category_map_file = dir + "category_map.csv"
categories_file = dir + "categories.txt"
outfile_csv = dir + "export.csv"
outfile_xlsx = dir + "export.xlsx"
bool_categorize_unmapped = True

class Transaction:
    def __init__(self,date,description,amount,type,category):
        self.date = date
        self.description = description
        self.amount = amount
        self.type = type
        self.category = category

class Menu:
    def __init__(self,prompt,menu_options):
        self.prompt = prompt
        self.menu_options = menu_options
        self.__prompt_user = False
        self.selection = None

    def print_menu(self):
        print("\n======================\n" + self.prompt)
        for option in self.menu_options:
            print(f"    {option[0]}: {option[1]}")
        return

    def get_user_input(self):
        self.__prompt_user = True
        while self.__prompt_user:
            self.print_menu()
            selection = input("Choose menu option: ")
            if selection.isalpha():
                selection = selection.upper()
            if selection not in dict(self.menu_options):
                print("Invalid input. Please enter the number/letter of one of the available menu options.")
            else:
                self.__prompt_user = False
        return selection

def import_transactions(dir):
    transactions = []
    try:
        for filename in os.listdir(dir):
            filepath = os.path.join(dir,filename)
            with open(filepath) as f:
                csv_reader = csv.reader(f)
                row_count = 0
                for row in csv_reader:
                    if filename.startswith("Chase"):
                        if row_count != 0:
                            date = row[0]
                            description = row[2]
                            type = row[4] 
                            amount = float(row[5])
                            if type == "Payment":
                                category = "Payment"
                            else:
                                category = get_category(description)
                            transactions.append(Transaction(date,description,amount,type,category))
                    elif filename.startswith("Checking"):
                        date = row[0]
                        description = row[4]
                        type = "Checking"
                        amount = float(row[1])
                        category = get_category(description)
                        transactions.append(Transaction(date,description,amount,type,category))
                    else:
                        print(f"Skipping invalid file: {filename}")
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
    category_map = read_category_map(category_map_file)
    # separate details and comments
    detail = None
    if "*" in description:
        description = description.split("*")
        detail = description[1]
        description = description[0]
    if "#" in description:
        description = description.split("#")[0]
    else:
        description = description

    # remove numbers and dates
    pattern = r'[0-9]|\/'
    description = re.sub(pattern,'',description)
    description = description.rstrip()

    if detail:
        if detail in category_map:
            return category_map[detail]
    if description in category_map:
        return category_map[description]
    elif bool_categorize_unmapped:
        return categorize_unmapped_transactions(description,category_map_file,categories_file,detail)
    else:
        return "Other"

def map_description_or_detail(description,detail):
    if detail:
        # Prompt user whether to create categorize mapping based on description or detail
        menu_items = []
        menu_items.append(("1","Description"))
        menu_items.append(("2","Detail"))
        menu_items.append(("S","[Skip Expense]"))
        menu_items.append(("Q","[Quit Categorizing]"))
        prompt = f"Categorizing expense: {description} (Detail: {detail})\nCategorize based on description or detail?"
        description_detail_menu = Menu(prompt,menu_items)
        selection = description_detail_menu.get_user_input()
        # Evaluate user response
        if selection == "Q":
            global bool_categorize_unmapped
            bool_categorize_unmapped = False
            return
        if selection == "S":
            return
        if selection == "2":
            return detail
        if selection == "1":
            return description
    else:
        return description

def create_category_mapping(description,map_file,categories_file):
    categories = [line.rstrip() for line in open(categories_file)]
    # Prompt user to map a category
    menu_items = []
    menu_num = 1
    for category in categories:
        menu_items.append((str(menu_num),category))
        menu_num+=1
    menu_items.append(("N","[Add New Category]"))
    menu_items.append(("S","[Skip Expense]"))
    menu_items.append(("Q","[Quit Categorizing]"))
    categorize_expense_menu = Menu("Create a category mapping for unmapped expense: " + description,menu_items)
    selection = categorize_expense_menu.get_user_input()
    # Evaluate user response
    # skip categorizing current transation and stop categorizing transactions
    if selection == "Q":
        global bool_categorize_unmapped
        bool_categorize_unmapped = False
        return "Other"
    # skip categorizing current transation
    elif selection == "S":
        return "Other"
    # prompt user and add new category name
    elif selection == "N":
        prompt_user = "True"
        while prompt_user:
            category = input("Enter name of new category: ")
            if category.isalnum():
                prompt_user = False
        with open(categories_file,'a') as f:
            f.write(category + "\n")
    # user selected an existing category
    else:
        category = categories[int(selection) - 1]
    # Add new mapping to category map file
    with open(map_file,'a') as f:
        row = description + "," + category + "\n"
        f.write(row)
    return category

def categorize_unmapped_transactions(description,map_file,categories_file,detail=None):
    str_to_evaluate = map_description_or_detail(description,detail)
    if str_to_evaluate:
        category = create_category_mapping(str_to_evaluate,map_file,categories_file)
    else:
        category = "Other"
    return category

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

        # create chart for total spending by category
        # excludes headers (row 1) and total (row 2)
        projected_pie = ProjectedPieChart()
        projected_pie.type = 'bar'
        data_range = Reference(summary_sheet,min_col=2,min_row=3,max_col=2,max_row=len(totals[current_month])+1)
        label_range = Reference(summary_sheet,min_col=1,min_row=3,max_col=1,max_row=len(totals[current_month])+1)
        projected_pie.add_data(data_range,titles_from_data=False)
        projected_pie.set_categories(label_range)
        projected_pie.dataLabels = DataLabelList()
        projected_pie.dataLabels.showPercent = True
        projected_pie.dataLabels.showLeaderLines = True
        projected_pie.dataLabels.showCatName = True
        projected_pie.title = "Expenses by Category"
        projected_pie.legend.position = 'b'
        summary_sheet.add_chart(projected_pie)

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
    try:
        wb.save(filepath)
        print(f"Wrote output report: {filepath}")
    except Exception as e:
        print(f"Error: unable to write output report: {e}")
    return

def main():
    transactions = import_transactions(import_dir)
    totals = calc_monthly_totals(transactions)
    # report_transaction_details(transactions)
    # report_monthly_totals(totals)
    export_excel(totals,transactions,outfile_xlsx)

main()