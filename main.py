# This is a sample Python script.
import random
import math
from typing import List

import pandas as pd
import constraint


# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.


def process_marks(filename, no_of_cos):
    number_cos = 4
    total_head = ""
    df = pd.read_excel(filename, skiprows=range(0, 3))

    # Get the list of column names
    column_names = list(df.columns)

    # Get the column names to be printed
    columns_to_print = ['RollNo', 'Name']
    for column_name in column_names:
        # print(column_name)
        if 'Total' in column_name:
            total_head = column_name
            columns_to_print.append(column_name)
    # Print the columns
    total = no_of_cos*5
    data_rows =[]
    for ind in df.index:
        mark_per = df[total_head][ind]/100

        if not pd.isnull(df['RollNo'][ind]):
            print(df['Name'][ind], df[total_head][ind])
            cos = constraint_solve(no_of_cos, math.floor(total*mark_per),math.ceil(total*mark_per))
            row = {'RollNo':df['RollNo'][ind], 'Name': df['Name'][ind], 'Total':df[total_head][ind]}
            row.update(cos)
            data_rows.append(row)

    return  data_rows





def constraint_solve(no_of_cos, lower, upper):
    problem = constraint.Problem()

    # Define variable names and their domains dynamically
    variables = []
    for i in range(no_of_cos):
        var_name = f"co{i}"
        problem.addVariable(var_name, range(1, 6))  # Variables can be 1 to 10
        variables.append(var_name)

    # Define the sum constraint
    def sum_constraint(*args):
        return lower <= sum(args) <= upper

    problem.addConstraint(sum_constraint, variables)

    # Find and print solutions
    solutions = problem.getSolutions()

    index = random.randint(0, len(solutions) - 1)
    return solutions[index]

def print_to_file(data_rows,no_of_cos):
    cos = {}
    for i in range(no_of_cos):
        var_name = f"co{i}"
        cos[var_name] =[]
    data = {
    'RollNo': [],
    'Name': [],
    'Total': [] }
    data.update(cos)


    df = pd.DataFrame(data)




    for row_data in data_rows:
        df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
        #df = df.append(row_data, ignore_index=True)

    # Save the DataFrame to an Excel file
    excel_file = 'example.xlsx'
    df.to_excel(excel_file, sheet_name='Sheet1', index=False, engine='openpyxl')

    print(f"Excel file '{excel_file}' created successfully.")

# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    #print(constraint_solve(4,18,20))
    print_to_file(process_marks('Sheet.xlsx',5),5)


