import os


from flask import Flask, request, render_template, send_file
import pandas as pd

from constraint import Problem
import random
import math
from io import BytesIO

app = Flask(__name__)




def process_qumark(qumark):
    items = qumark.split(',')
    parsed_data = []
    # Iterate over each item
    for item in items:
        # Split each item by dash to extract individual parts
        parts = item.split('-')
        parsed_data.append((parts[0], parts[1], parts[2]))

    return parsed_data

def distribute_with_max(total, max_values):
    num_columns = len(max_values)

    # Generate random values for each column within the range of 0 to its maximum value
    column_values = [random.randint(0, max_value) for max_value in max_values]

    # Calculate the sum of the generated values
    current_total = sum(column_values)

    # Adjust the values to ensure that the sum of all columns equals the total value
    if current_total != total:
        # Determine the adjustment needed
        adjustment = total - current_total

        # Distribute the adjustment randomly across the columns
        while adjustment != 0:
            # Randomly select a column
            column_index = random.randint(0, num_columns - 1)

            # Adjust the value of the selected column
            max_value = max_values[column_index]
            if adjustment > 0 and column_values[column_index] < max_value:
                column_values[column_index] += 1
                adjustment -= 1
            elif adjustment < 0 and column_values[column_index] > 0:
                column_values[column_index] -= 1
                adjustment += 1

    return column_values


def process_compo_mark(df, qumark, componet):
    total_head = ""
    column_names = list(df.columns)
    #print(column_names)
    data_rows = []
    columns_to_print = ['RollNo', 'Name']
    for column_name in column_names:
        # print(column_name)
        if componet in column_name:
            total_head = column_name
            columns_to_print.append(column_name)
    # Print the columns
    processed_qmark = process_qumark(qumark)
    qmarkonly = [int(item[1]) for item in processed_qmark]
    cos = [item[2] for item in processed_qmark]
    qhead = [item[0] for item in processed_qmark]
    row = {'RollNo': "", 'Name': "", total_head: ""}
    row.update({head: total for head, total in zip(qhead,  qmarkonly)})
    data_rows.append(row)
    row = {'RollNo': "", 'Name': "", total_head: ""}
    row = {head: total for head, total in zip(qhead,  cos)}
    data_rows.append(row)
    columns_to_print.append(qhead)
    #print(qmarkonly)
    for ind in df.index:
        mark = df[total_head][ind]
        if math.isnan(mark):
            break
        #print(mark)
        distrubted= distribute_with_max(mark,qmarkonly)

        row = {'RollNo': df['RollNo'][ind], 'Name': df['Name'][ind], total_head: df[total_head][ind]}
        row.update({head: total for head, total in zip(qhead,  distrubted)})
        data_rows.append(row)
    return columns_to_print, data_rows,qhead





@app.route("/mark_dist", methods=["POST"])
def mark_dist():
    file = request.files["file"]
    qumark = request.form["qumark"]
    componet = request.form["componet"]
    df = pd.read_excel(file)
    columns_to_print, data_rows,qhead= process_compo_mark(df, qumark, componet)
    return print_to_file_distributed(columns_to_print, data_rows,qhead)



@app.route("/")
def index():
    return render_template("ce.html")

@app.route("/index.html")
def index_html():
    return render_template("ce.html")




def print_to_file_distributed(coloums_head, data_rows,qheads):

    avgval = []
    df = pd.DataFrame()
    for row_data in data_rows:
        df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1')
    writer.close()
    output.seek(0)
    return send_file(output, download_name='output.xlsx', as_attachment=True)


@app.errorhandler(500)
def internal_error(error):

     return render_template("ceerror.html")

if __name__ == "__main__":
    from waitress import serve

    serve(app, host="0.0.0.0", port=8080)
    total = 12
    max_values = [4, 4, 5, 10]

    distributed_values = distribute_with_max(total, max_values)
    if distributed_values:
        print("Sum of distributed values:", sum(distributed_values))
    else:
        print("No solution found!")

    #app.run(debug=True, port=5001, host='0.0.0.0')
    # print(maximize_multiple_variables(4,16,18))
