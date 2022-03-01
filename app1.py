import json
from flask import Flask,render_template, request, jsonify, url_for,Response, redirect
import pandas as pd
from flask_cors import cross_origin
#import matplotlib.pyplot as plt
#import xlwt
# need to install xlrd and openpyxl to use read_excel function in pandas
import os
import csv
import shutil
app = Flask(__name__)
@app.route('/')
@app.route('/home')
@cross_origin()
def home_page():

    return render_template('index1.html')

@app.route('/about')
def about():
    title= "About time point screening"
    return render_template('about.html', title=title)

@app.route('/clear_data',methods=['GET', 'POST'])
def clear_data():
    filepath = ('Static\Input.xlsx')
    filepath1 = ('Static\Temp.xlsx')
    if os.path.isfile(filepath):
        os.remove(filepath)
        comment =" file has been cleared"
    if os.path.isfile(filepath1):
        os.remove(filepath1)
        comment = "  Temp File has been cleared"
    else:
        comment = "     !!!! No file Exists"
    return render_template('download.html', clear_data=comment)

@app.route('/data', methods=['GET', 'POST'])
def data():
    if request.method == 'POST':
        try:

            #file = request.form['upload-file']
            file = request.files['upload-file']
            if not os.path.isdir('Static'):
                os.mkdir('Static')

            filepath1 = os.path.join('Static', file.filename)
            print(filepath1)
            if os.path.isfile(filepath1):
                os.remove(filepath1)
                print("source file is copied to dataframe and deleted")
            file.save(filepath1)
            #dest = os.path.join('Static' + 'worksheet.xlsx' )
            #os.rename(filepath1,dest)
            filepath = ('Static\working_data.xlsx')
            #data1 = pd.read_excel(file)

            #data1 = pd.read_excel(filepath1,sheet_name='Temp vs Time', skiprows = 3)
            #temp_data= pd.read_excel(filepath1,sheet_name='Temperature difference', skiprows = 1)
            pressure_data = pd.read_excel(filepath1, sheet_name='Input',index_col='Time', skiprows=1)
            pressure_data = pressure_data[pressure_data.filter(regex='^(?!Unnamed)').columns]
            # Remove the 1st row with units from the table
            pressure_data = pressure_data.drop(pressure_data.index[[0]])
            data1 = pressure_data
            data1.to_excel('Static\Temp.xlsx')
            coll = pressure_data.columns.to_list()
            for col in pressure_data:
                pressure_data[col] = pd.to_numeric(pressure_data[col], errors='coerce',downcast='float')
            data22 = pd.read_excel(filepath1, sheet_name='Interpolate_data', index_col='Time')
            # keep the data to be interpolated in sheet named Interpolate _data
            # make sure for data22 the excel has file name Interpolate_data and has 1st column with "Time"
            data22 = data22[data22.filter(regex='^(?!Unnamed)').columns]
            aa = data22.index.values
            ser= pd.Series(aa)
            #aa = pd.to_numeric(aa)
            # creating a blank dataframe with columns from the inputs and row values to be interpolated
            new = pd.DataFrame(columns=coll, index=ser)
            # combining two dataframe into one dataframe to find the missing value by interpolation with index method.
            concatenated = pd.concat([pressure_data , new])
            con = concatenated.interpolate(method='index')
            # save the interpolated dataframe to excel file.
            #con.to_excel('Static\Interpolated_file.xlsx')
            os.remove(filepath1)
            print(con.shape)
            return Response(
                con.to_csv(),
                mimetype="text/csv",
                headers={"Content-disposition":
                             "attachment; filename=Interpolated file.csv"})
        except Exception as e:
            print('The Exception message is: ', e)
            return 'something is wrong in Interpolate module, check for input file'


if __name__   ==   '__main__':
    app.run(debug=True)
