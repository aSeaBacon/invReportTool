from flask import Flask, render_template, request, url_for, redirect, send_file
import xlsxwriter
import re
from operator import itemgetter
from datetime import datetime, date
import webbrowser

app = Flask(__name__)
     
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/converter', methods=['GET', 'POST'])
def invConvert():
    #Create Excel Sheet and add headers row
    column_names = ["ORG.", "Column 1", "CL TYP SUB", "LOC", "NUMBER", "DATE", "AUTH NO.", "ACT", "FROM/TO MMS CODE", "UNITS", "AMOUNT", "RECEIVED"]
    workbook = xlsxwriter.Workbook('static/output.xlsx')
    worksheet = workbook.add_worksheet()
    for i, name in enumerate(column_names):
        worksheet.write(0, i, name)


    #Iterate through text files uploaded, find all lines which are order entires (not summaries) and add them to OrderList
    orderList = []

    for file in request.files.getlist('filename'):
        file.save('static/temp.txt')
        with open('static/temp.txt', 'r') as f:
            for line in f:
                if re.search("\d{2}/\d{2}/\d{2}", line) and not re.search("DIVISION", line):
                    tempLine = line.replace("?", "")
                    orderItems = tempLine.split()
                    orderList.append(orderItems)

    #If lines start with Location# and not Org#, add previous lines org#/column1/CL TYP SUB (these orders are part of multiline orders)
    for item in orderList:
        if len(item[0]) == 2:
            for i in range(3):
                item.insert(i, prevItem[i])
        prevItem = item


    #If amount of items==13/11, merge items 8 + 9 (13) or add an empty string(11) (From/To MMS CODE is either blank, 3 digit number, or 2 strings)
    for item in orderList:
        if len(item) == 13:
            temp = item[9]
            item[8] = str(item[8]) + ' ' + str(item[9])
            item.pop(9)

        elif len(item) == 11:
            item.insert(8, "")

    #Convert dates to datetime objects for sorting
    for item in orderList:
        item[5] = datetime.strptime(item[5], "%m/%d/%y")

    #Sort entries by org# then by date, then by CL TYP SUB
    sorted_orders = sorted(orderList, key=itemgetter(0, 5, 3))

    #Populate Excel sheet and close
    for i, item in enumerate(sorted_orders):
        for j, entry in enumerate(item):
            #Convert datetime object to string
            if j==5:
                worksheet.write(i+1, j, entry.strftime("%m/%d/%y"))
            else:
                worksheet.write(i+1, j, entry)

    worksheet.autofit()
    workbook.close()

    return send_file('static/output.xlsx')


if __name__ == '__main__':
    webbrowser.open('http://127.0.0.1:5000/')
    app.run(debug = True, use_reloader=False)
