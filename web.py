import os

#import pulp
from flask import Flask, request, render_template, send_file
import pandas as pd
#from werkzeug.utils import secure_filename
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
    return print_to_file(process_grade(df, no_of_co), no_of_co)


@app.route("/")
def index():
    return render_template("index.html")

@app.route("/index.html")
def index_html():
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
    total = no_of_cos * 5
    data_rows = []
    aver = [0] * no_of_cos;
    average = df.loc[:, total_head].mean()

    for ind in df.index:
        mark_per = df[total_head][ind] / 100
        upper = df[total_head][ind]+average

        if total*mark_per > average/100:
            max_value = total*mark_per
        else:
            max_value = average/100
        if max_value < no_of_cos:
             max_value = no_of_cos
        if not pd.isnull(df['RollNo'][ind]):

            cos =constraint_solve(no_of_cos, math.ceil(max_value))
            #cos = maximize_multiple_variables(no_of_cos,total/100 * mark_per, (total+average)/100 * mark_per)
            # cos = maximize_multiple_variables(no_of_cos, math.floor(total*mark_per+(average-df[total_head][ind])/100),math.ceil(total*mark_per+(average-df[total_head][ind])/100))
            row = {'RollNo': df['RollNo'][ind], 'Name': df['Name'][ind], 'Total': df[total_head][ind]}
            #sorted_values = sorted(cos.values(), reverse=True)
            #cos = {k: sorted_values[i] for i, k in enumerate(cos.keys())}
            row.update(cos)
            data_rows.append(row)
    return data_rows

def process_grade(df, no_of_cos):
    total_head = ""
    column_names = list(df.columns)
    columns_to_print = ['RollNo', 'Name']
    for column_name in column_names:
        # print(column_name)
        if 'Total' in column_name:
            total_head = column_name
            columns_to_print.append(column_name)
    # Print the columns
    columns_to_print.append('Grade')
    total = no_of_cos * 5
    data_rows = []
    aver = [0] * no_of_cos;
    average = df.loc[:, total_head].mean()

    for ind in df.index:
        grade = df['Grade'][ind]
        lower = 0
        upper = 0
        if not pd.isnull(df['RollNo'][ind]):
            grade = df['Grade'][ind]
            if grade == "O":
                lower = 5
                upper = 5
            elif grade == "A+":
                lower = 5
                upper = 5
            elif grade == "A":
                lower = 4
                upper = 5
            elif grade == "B+":
                lower = 3
                upper = 4
            elif grade == "B":
                lower = 2
                upper = 4
            elif grade == "C":
                lower = 2
                upper = 3
            elif grade == "P":
                lower = 1
                upper = 2
            else:
                lower = 1
                upper = 1
            cos = constraint_solve_for_grade(no_of_cos, lower,upper)
            #row = {'RollNo': df['RollNo'][ind], 'Name': df['Name'][ind], 'Total': df[total_head][ind],'Grade': df['Grade'][ind]}
            row = {'RollNo': df['RollNo'][ind], 'Name': df['Name'][ind]}
            if random.randrange(1, 5) > 1:
                sorted_values = sorted(cos.values(), reverse=True)
                cos = {k: sorted_values[i] for i, k in enumerate(cos.keys())}
            row.update(cos)
            data_rows.append(row)
    return data_rows

def maximize_multiple_variables(no_of_cos, lower, upper):
    problem = pulp.LpProblem("Maximization Problem", pulp.LpMaximize)

    variables = {}

    for i in range(no_of_cos):
        var_name = f"co{i}"
        variables[var_name] = pulp.LpVariable(var_name, lowBound=1, upBound=5)


    # Add the sum constraint
    problem += pulp.lpSum(variables.values()) >= lower
    problem += pulp.lpSum(variables.values()) <= upper

    # Define the objective to maximize
    objective = pulp.lpSum(variables[var_name] for var_name in variables)
    problem += objective

    problem.solve()

    if pulp.LpStatus[problem.status] == "Optimal":
        optimal_solution = {var.name: var.varValue for var in problem.variables()}
        return optimal_solution
    else:
        return None


def constraint_solve(no_of_cos,  upper):
    problem = constraint.Problem()
    variables = []
    for i in range(no_of_cos):
        var_name = f"co{i + 1}"
        problem.addVariable(var_name, range(1, 6))  # Variables can be 1 to 10
        variables.append(var_name)

    def sum_constraint(*args):
        return sum(args) <= upper
    # problem.addConstraint()
    problem.addConstraint(sum_constraint, variables)
    #problem.setObjective(lambda *args: sum(args[var_name] for var_name in variables_to_maximize), maximize=True)
    # Find and print solutions
    solutions = problem.getSolutions()
    #index1 = random.randint(0, len(solutions) - 1)
    if solutions:
        max_solution = max(solutions, key=lambda x: sum(x[var] for var in variables))
        return max_solution
    else:
        return None

    #return solutions[index1]

def constraint_solve_for_grade(no_of_cos,  lower, upper):

    variables = {}
    for i in range(no_of_cos):
        var_name = f"co{i + 1}"
        variables.update()
        variables[var_name] = random.randrange(lower, upper+1)
    return variables

def print_to_file(data_rows, no_of_cos):
    cos = {}
    average = {}
    percentage = {}
    for i in range(no_of_cos):
        var_name = f"co{i + 1}"
        cos[var_name] = []
    data = {
        'RollNo': [],
        'Name': []}
    data.update(cos)
    df = pd.DataFrame(data)
    for row_data in data_rows:
        df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
    for i in range(no_of_cos):
        var_name: str = f"co{i + 1}"
        average[var_name] = df.loc[:, var_name].mean()
        percentage[var_name] = average[var_name]/5*100
    df = pd.concat([df, pd.DataFrame([average])], ignore_index=True)
    df = pd.concat([df, pd.DataFrame([percentage])], ignore_index=True)
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1')
    writer.close()
    output.seek(0)
    return send_file(output, download_name='output.xlsx', as_attachment=True)

@app.errorhandler(500)
def internal_error(error):

     return render_template("error.html")

if __name__ == "__main__":
    from waitress import serve

    serve(app, host="0.0.0.0", port=80)

    #app.run(debug=True, port=5001, host='0.0.0.0')
    # print(maximize_multiple_variables(4,16,18))
