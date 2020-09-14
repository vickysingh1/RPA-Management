from flask import Flask, render_template, request, url_for, redirect,jsonify,session,g, request, redirect, url_for,send_from_directory
import re
import os
import sys
import json
import pymysql
from datetime import timedelta
from openpyxl import Workbook
import flask_login
import csv
from functools import wraps
app = Flask(__name__)
app.config['SECRET_KEY'] = 'd369342136ecd032f8b4a930b6bb2e0e'

###########################################Database Connection Variable##########################################################
Host_Url="127.0.0.1"
DB_User="root"
DB_Password="vicky"
DB_name="rpa"
Port="3306"


####################################################################################################
def login_required(f):
    @wraps(f)
    def wrap(*args, **kwargs):
        if 'logged_in' in session:   
            return f(*args, **kwargs)
        else:
            return redirect(url_for('Login'))
    return wrap

@app.before_request
def before_request():
    session.parmanent=True
    app.permanent_session_lifetime=timedelta(seconds=30)
    session.modified=True
    g.user=flask_login.current_user



@app.route('/')
def Login():
    return render_template('Login.html')







@app.route('/check', methods=['POST'])
def check():
    try:
        db=pymysql.connect(Host_Url,DB_User,DB_Password,DB_name)
        print("connected")
        cursor=db.cursor()
        username=request.form['Log_in']
        password=request.form['pass']
        User_Name=username
        cursor.execute('SELECT * FROM user WHERE userid = %s AND password = %s', (username, password))
        account = cursor.fetchone()   # Fetch one record and return result
        print(account)
        if account:
                # Create session data, we can access this data in other routes
                session['logged_in']=True
                session['username']=request.form['Log_in']
                msg="Login Successfull"
                return json.dumps(msg)
                
        else:
                # Account doesnt exist or username/password incorrect
                msg = 'Incorrect username/password!'
                return json.dumps(msg)
    except Exception as ex:
        print(ex)
    
    
@app.route('/index')
@login_required
def indexmain():
    return render_template('index.html')
@app.route('/ind', methods=['POST'])
@login_required
def index1():
    try:
        db=pymysql.connect(Host_Url,DB_User,DB_Password,DB_name)
        print("connected successfully")
        cursor = db.cursor()
        id=request.form['id']
        id=int(id)
        Tc=request.form['TestCases']
        Tc=str(Tc)
        org=request.form['Organization']
        org=int(org)
        project=request.form['project']
        project=str(project)
        serviceid=request.form['Serviceid']
        serviceid=int(serviceid)
        service=request.form['Service']
        service=str(service)
        active=request.form['Active']
        active=int(active)
        location=request.form['Loc']
        location=str(location)
        Created=request.form['Created_By']
        Created=str(Created)
        Casetype=request.form['abc']
        Casetype=str(Casetype)
        TestType=request.form['TestType']
        TestType=str(TestType)
        # print(id, Tc, org,project,serviceid, service, active, location, Created, Casetype, TestType)
        cursor.execute("""INSERT INTO Test_Cases(Test_Case_ID,RPA_Test_Case,Organisation_ID,Project,Service_ID,Service_Name,isActive,Script_Location,CreatedBY,TestCase_Type,Testing_Type) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",(id,Tc,org,project,serviceid,service,active,location,Created,Casetype,TestType))
        # cursor.execute(execute, {'Test_Case_ID':id, 'RPA_Test_Case':Tc, 'Organisation_ID':org,'Project':project,'Service_ID':serviceid, 'Service_Name':service, 'isActive':active,'Script_Location':location,'CreatedBY':Created,'TestCase_Type':Casetype,'Testing_Type':TestType})
        db.commit()
        data1="Data save successfully "
    except Exception:
        db.close()
        data1="Data is Not Save Successfully"
    return json.dumps(data1)



@app.route('/showdata')
@login_required
def showdata():
    try:
        db=pymysql.connect(Host_Url,DB_User,DB_Password,DB_name)
        # print("connected successfully")
        cursor = db.cursor()
        cursor.execute("select * from Test_Cases")
        data1 = cursor.fetchall()
        data2=data1[::-1]
        db.commit()
        db.close()
    except Exception as ex:
        print(ex)
        db.close()
    return render_template('showdata.html',data2=data2)



@app.route("/getPlotCSV")
@login_required
def getCSV():
    try:
        db=pymysql.connect(Host_Url,DB_User,DB_Password,DB_name)
        # print("connected successfully")
        cursor = db.cursor()
        cursor.execute("select * from Test_Cases")
        data2 = cursor.fetchall()
        wb = Workbook(write_only=True)
        data_ws = wb.create_sheet("testdata")
        # write header
        data_ws.append(["Test_Case_ID","RPA_Test_Case","Organisation_ID","Project","Service_ID","Service_Name","isActive","Script_Location","CreatedBY","TestCase_Type","Testing_Type"])
        # write data
        for data in data2:
            Test_Case_ID = data[0]
            RPA_Test_Case = data[1]
            Organisation_ID = data[2]
            Project = data[3]
            Service_ID = data[4]
            Service_Name = data[5]
            isActive = data[6]
            Script_Location = data[7]
            CreatedBY = data[8]
            TestCase_Type = data[9]
            Testing_Type = data[10]
            data_ws.append([Test_Case_ID,RPA_Test_Case,Organisation_ID,Project,Service_ID,Service_Name,isActive,Script_Location,CreatedBY,TestCase_Type,Testing_Type])

        wb.save("testdata.xlsx")
    except Exception as ex:
        print(ex)
        db.close()
    excel_file = 'testdata.xlsx'
    return send_from_directory(filename = excel_file,directory = "./",as_attachment = True,cache_timeout = 0)
@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for('Login'))






if __name__=="__main__":
    app.run(debug=True)
    
    #app.run()