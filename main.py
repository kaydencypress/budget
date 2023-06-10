import csv
import os.path
import string

directory = "/home/dev_iant/workspace/github.com/kaydencypress/budget/"
credit_csv = directory + "chase_export.csv"
category_map_csv = directory + "category_map.csv"

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
                    amount = row[5]
                    category = get_category(type,description)
                    current_transaction = Transaction(date,description,amount,type,category)
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

def get_category(type,description):
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

def main():
    credit_transactions = import_credit_csv(credit_csv)
    for transaction in credit_transactions:
        print(f"Transaction on {transaction.date} for {transaction.amount}: {transaction.description} [{transaction.category}]")

main()