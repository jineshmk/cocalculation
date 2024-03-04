import os


from flask import Flask, request, render_template, send_file
import pandas as pd
import pulp
from constraint import Problem
import random
import math
from io import BytesIO

app = Flask(__name__)





def distribute_mark(total_mark,  question_marks_list):
    prob = pulp.LpProblem("Maximize_Threshold", pulp.LpMinimize)
    marks_vars = {}
    threshold_vars = {}
    for q in question_marks_list:
        question_number, max_marks, threshold, priority = q
        variable_name = f"{question_number}"

        marks_vars[question_number] = pulp.LpVariable(variable_name, lowBound=0, upBound=max_marks, cat='Integer')

        #threshold_vars[question_number] = pulp.LpVariable(f"Threshold_{question_number}", lowBound=0, cat='Continuous')

    # Add objective function (maximize threshold value)
    prob += pulp.lpSum(marks_vars.values()), "Maximize_Threshold_Value"

    # Add constraints to ensure distributed marks are equal total marks available
    prob += pulp.lpSum(marks_vars.values()) == total_mark, "Total_Marks_Constraint"
    for question_number, marks_var in marks_vars.items():
        prob += marks_var >= 0, f"Non_Negative_Marks_{question_number}"

    # Add constraints to ensure marks meet threshold percentage
    for q in question_marks_list:

        question_number, max_marks, threshold, priority = q

        if priority == 1:
            prob += marks_vars[question_number] >= max_marks * threshold / 100, f"Threshold{question_number}"
        else:
            prob += marks_vars[question_number] >= random.randint(0, max_marks), f"Random_Distribution_{question_number}"

    # Solve the problem

    prob.solve()

    marks_dict = {question_number: marks_vars[question_number].varValue for question_number in marks_vars}
    # Set small values close to zero to zero
    epsilon = 1e-10  # Adjust this value as needed
    for question_number, mark in marks_dict.items():
        if abs(mark) < epsilon:
            marks_dict[question_number] = 0
    if all(mark >= 0 for mark in marks_dict.values()):
        return marks_dict
    else:
       return distribute_mark(total_mark,question_marks_list)










@app.route("/mark_dist", methods=["POST"])
def mark_dist():
    print("here")
    mark_excel = pd.read_excel(request.files["total"])
    mark_dist_excel = pd.read_excel(request.files["qmark"])
    qumark_list = mark_dist_excel[['No', 'Mark', 'Threshold','Priority']].values.tolist()
    componet = request.form["componet"]
    mark_list = mark_excel[['RollNo','Name',componet]].values.tolist()

    data_rows = []
    for roll_no,name,mark in mark_list:
        row = {'RollNo':roll_no, 'Name': name, 'Total': mark}
        dist_mark = distribute_mark(mark,qumark_list)
        row.update(dist_mark)
        data_rows.append(row)
    df = pd.DataFrame()
    for row_data in data_rows:
        df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1')
    writer.close()
    output.seek(0)
    return send_file(output, download_name='output.xlsx', as_attachment=True)

@app.route("/")
def index():
    return render_template("ce.html")

@app.route("/index.html")
def index_html():
    return render_template("ce.html")


@app.errorhandler(500)
def internal_error(error):

     return render_template("ceerror.html")

if __name__ == "__main__":
    from waitress import serve

    serve(app, host="0.0.0.0", port=8080)

