import os

from flask import Flask, request, render_template, send_file
import pandas as pd
from werkzeug.utils import secure_filename
import constraint
import random
import math
from io import BytesIO
app = Flask(__name__)

@app.route("/upload_file", methods=["POST"])
def upload_file():
    file = request.files["file"]
    no_of_co = int(request.form["noofco"])
    df = pd.read_excel(file, skiprows=range(0, 3))
    return print_to_file(process_marks(df,no_of_co),no_of_co)

@app.route("/")
def index():
    return render_template("index.html")


def process_marks(df, no_of_cos):
    total_head = ""
    column_names = list(df.columns)
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
            cos = constraint_solve(no_of_cos, math.floor(total*mark_per),math.ceil(total*mark_per))
            row = {'RollNo':df['RollNo'][ind], 'Name': df['Name'][ind], 'Total':df[total_head][ind]}
            row.update(cos)
            data_rows.append(row)
    return data_rows


def constraint_solve(no_of_cos, lower, upper):
    problem = constraint.Problem()
    variables = []
    for i in range(no_of_cos):
        var_name = f"co{i}"
        problem.addVariable(var_name, range(1, 6))  # Variables can be 1 to 10
        variables.append(var_name)

    def sum_constraint(*args):
        return lower <= sum(args) <= upper

    problem.addConstraint(sum_constraint, variables)
    # Find and print solutions
    solutions = problem.getSolutions()
    index1 = random.randint(0, len(solutions) - 1)
    return solutions[index1]


def print_to_file(data_rows, no_of_cos):
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
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1')
    writer.close()
    output.seek(0)
    return send_file(output, download_name='output.xlsx', as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
