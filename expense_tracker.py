import csv
import os
import argparse
from datetime import datetime

import matplotlib.pyplot as plt
from openpyxl import Workbook

FILE_NAME = "transactions.csv"

def create_file():

    if not os.path.exists(FILE_NAME):

        with open(FILE_NAME, "w", newline="", encoding="utf-8") as file:

            writer = csv.writer(file)

            writer.writerow([
                "Date",
                "Type",
                "Category",
                "Amount",
                "Description"
            ])



def add_transaction(transaction_type, category, amount, description):

    create_file()

    date = datetime.now().strftime("%Y-%m-%d")

    with open(FILE_NAME, "a", newline="", encoding="utf-8") as file:

        writer = csv.writer(file)

        writer.writerow([
            date,
            transaction_type,
            category,
            amount,
            description
        ])

    print("\nTransaction added successfully!")



def read_transactions():

    create_file()

    transactions = []

    with open(FILE_NAME, "r", encoding="utf-8") as file:

        reader = csv.DictReader(file)

        for row in reader:

            transactions.append(row)

    return transactions



def show_transactions(transactions):

    if not transactions:

        print("\nNo transactions found.")

        return

    print("\n" + "=" * 80)
    print("TRANSACTIONS")
    print("=" * 80)

    for index, transaction in enumerate(transactions, start=1):

        print(
            f"{index}. "
            f"Date: {transaction['Date']} | "
            f"Type: {transaction['Type']} | "
            f"Category: {transaction['Category']} | "
            f"Amount: ₹{transaction['Amount']} | "
            f"Description: {transaction['Description']}"
        )



def calculate_summary(transactions):

    income = 0
    expense = 0

    for transaction in transactions:

        amount = float(transaction["Amount"])

        if transaction["Type"].lower() == "income":

            income += amount

        elif transaction["Type"].lower() == "expense":

            expense += amount

    balance = income - expense

    print("\n" + "=" * 50)
    print("FINANCIAL SUMMARY")
    print("=" * 50)

    print(f"Total Income  : ₹{income:.2f}")
    print(f"Total Expense : ₹{expense:.2f}")
    print(f"Balance       : ₹{balance:.2f}")



def monthly_summary(month):

    transactions = read_transactions()

    monthly_transactions = []

    for transaction in transactions:

        if transaction["Date"].startswith(month):

            monthly_transactions.append(transaction)

    print(f"\nMonthly Summary: {month}")

    calculate_summary(monthly_transactions)



def filter_transactions(category=None, transaction_type=None):

    transactions = read_transactions()

    filtered = []

    for transaction in transactions:

        category_match = True
        type_match = True

        if category:

            category_match = (
                transaction["Category"].lower()
                == category.lower()
            )

        if transaction_type:

            type_match = (
                transaction["Type"].lower()
                == transaction_type.lower()
            )

        if category_match and type_match:

            filtered.append(transaction)

    show_transactions(filtered)



def export_excel():

    transactions = read_transactions()

    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "Transactions"

    headers = [
        "Date",
        "Type",
        "Category",
        "Amount",
        "Description"
    ]

    sheet.append(headers)

    for transaction in transactions:

        sheet.append([
            transaction["Date"],
            transaction["Type"],
            transaction["Category"],
            transaction["Amount"],
            transaction["Description"]
        ])

    sheet.column_dimensions["A"].width = 15
    sheet.column_dimensions["B"].width = 15
    sheet.column_dimensions["C"].width = 20
    sheet.column_dimensions["D"].width = 15
    sheet.column_dimensions["E"].width = 40

    workbook.save("expense_report.xlsx")

    print("\nExcel report created: expense_report.xlsx")


def create_chart():

    transactions = read_transactions()

    categories = {}

    for transaction in transactions:

        if transaction["Type"].lower() == "expense":

            category = transaction["Category"]

            amount = float(transaction["Amount"])

            if category in categories:

                categories[category] += amount

            else:

                categories[category] = amount

    if not categories:

        print("\nNo expense data available for chart.")

        return

    names = list(categories.keys())

    amounts = list(categories.values())

    plt.figure(figsize=(8, 5))

    plt.bar(names, amounts)

    plt.title("Expenses by Category")

    plt.xlabel("Category")

    plt.ylabel("Amount")

    plt.xticks(rotation=45)

    plt.tight_layout()

    plt.savefig("expense_chart.png")

    plt.show()

    print("\nChart saved as expense_chart.png")



def main():

    parser = argparse.ArgumentParser(
        description="Expense Tracker CLI"
    )

    parser.add_argument(
        "--add-income",
        nargs=3,
        metavar=("CATEGORY", "AMOUNT", "DESCRIPTION"),
        help="Add income"
    )

    parser.add_argument(
        "--add-expense",
        nargs=3,
        metavar=("CATEGORY", "AMOUNT", "DESCRIPTION"),
        help="Add expense"
    )

    parser.add_argument(
        "--show",
        action="store_true",
        help="Show all transactions"
    )

    parser.add_argument(
        "--summary",
        action="store_true",
        help="Show total income, expense and balance"
    )

    parser.add_argument(
        "--month",
        help="Monthly summary. Example: 2026-10"
    )

    parser.add_argument(
        "--category",
        help="Filter by category"
    )

    parser.add_argument(
        "--type",
        choices=["income", "expense"],
        help="Filter by transaction type"
    )

    parser.add_argument(
        "--excel",
        action="store_true",
        help="Export transactions to Excel"
    )

    parser.add_argument(
        "--chart",
        action="store_true",
        help="Create expense chart"
    )

    args = parser.parse_args()

    create_file()



    if args.add_income:

        category = args.add_income[0]

        amount = args.add_income[1]

        description = args.add_income[2]

        add_transaction(
            "Income",
            category,
            amount,
            description
        )

   

    elif args.add_expense:

        category = args.add_expense[0]

        amount = args.add_expense[1]

        description = args.add_expense[2]

        add_transaction(
            "Expense",
            category,
            amount,
            description
        )

    

    elif args.show:

        transactions = read_transactions()

        show_transactions(transactions)

    

    elif args.summary:

        transactions = read_transactions()

        calculate_summary(transactions)


    elif args.month:

        monthly_summary(args.month)

    elif args.category or args.type:

        filter_transactions(
            args.category,
            args.type
        )


    elif args.excel:

        export_excel()


    elif args.chart:

        create_chart()

    else:

        parser.print_help()


if __name__ == "__main__":

    main()