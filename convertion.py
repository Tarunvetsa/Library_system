import json
from os import access
import smtplib
from datetime import datetime
from email.message import EmailMessage
from openpyxl import load_workbook
from datetime import date
from datetime import timedelta, datetime
from flask import Flask, render_template, request, jsonify   
 

app = Flask(__name__)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/home/")
def home1():
    return render_template("home.html")       
    
@app.route("/about/")
def home6():
    return render_template("about.html")       
    
@app.route("/iss/")
def home5():
    return render_template("iss.html")

@app.route("/return/")
def home4():
    return render_template("return.html")    

@app.route("/submitJSON", methods=["POST"])
def processJSON(): 
    jsonStr = request.get_json()
    jsonObj = json.loads(jsonStr) 
    
    response = ""
    # temp1=jsonObj['temp1']
    # temp2=jsonObj['temp2']
    # response+="<b> The input temparatures in Celcius are: <b>"+temp1+" and "+temp2+"</b><br>"
        
    
    
    	    
    return response
@app.route("/submitJSON4", methods=["POST",'GET'])
def processJSON4(): 
    jsonStr = request.get_json()
    jsonObj = json.loads(jsonStr) 
    response = ""
    name=jsonObj['name']
    roll=jsonObj['roll']
    gmail=jsonObj['gmail']
    Book=jsonObj['Book']
    Access=jsonObj['access']
    response = ""
    duedate=4

    msg=EmailMessage()
    msg['Subject']='Book Return'
    msg['From']='libraryiitbhilai@gmail.com'
    msg['To']=gmail
    f=open("return.txt","w")
    f.write(f"Dear {name} \n")
    f.close()
    f=open("return.txt","a")
    f.write(f"\nThe following book(s) have been returned to the library by you:\nTitle: {Book}\nAccession No.: {Access}\n\nThank you for visiting IITBhilai Library.\n\nNOTE:\nThe Fine Structure will be Following (Which is applicable to all users):\nF1. First 2 week:   2 Rupees will be charged per day.\nF2. After 2 weeks:  100 rupees will be charged per week.\n\n\nLibrary ,IIT Bhilai\nAcademic Building ,Raipur\nMob No-7725932105\n ")
    f.close()
    #determining the total no of entries in the db file
    # with open('count.txt','r') as f:
    #     no=int(f.read())+1
    
    with open("return.txt") as myfile:
        body=myfile.read()
        msg.set_content(body)
    myfile.close()

    server=smtplib.SMTP_SSL("smtp.gmail.com",465)
    server.login("libraryiitbhilai@gmail.com","library@123")
    server.send_message(msg)
    server.quit()




    return response

@app.route("/submitJSON2", methods=["POST",'GET'])
def processJSON1(): 
    jsonStr = request.get_json()
    jsonObj = json.loads(jsonStr) 
    response = ""
    name=jsonObj['name']
    roll=jsonObj['roll']
    gmail=jsonObj['gmail']
    Book=jsonObj['Book']
    Access=jsonObj['access']
 
    today = date.today()
    today=str(today)
    year=int(today[0:4])
    month=int(today[5:7])
    datetoday=int(today[8:10])


    issuedate=today
    issue = datetime(year,month,datetoday)

    duedate=issue + timedelta(days=14)
    duedate=str(duedate)
    t="23:59"
    duedate=duedate[:10]+t
    # print(ans[:10])
    # print(today)


    workbook = load_workbook(filename="sample.xlsx")
    sheet = workbook.active
    no=7
    
    f=open("index.txt","r")
    index=f.read()
    no=int(index)
    no=no+1
    print(type(index))

    f.close()

    f=open("index.txt","w")

    f.write(str(no))
    f.close()




    sheet[f"A{no}"] = name
    sheet[f"C{no}"] = roll
    sheet[f"E{no}"] = gmail
    sheet[f"H{no}"] = Book
    sheet[f"M{no}"] = Access
    sheet[f"O{no}"] = issuedate


    workbook.save(filename="sample.xlsx")



    response+="<b> The input temparatures in Celcius are: <b>"+name+" and "+roll+"</b><br>"
    msg=EmailMessage()
    msg['Subject']='Book Issue'
    msg['From']='libraryiitbhilai@gmail.com'
    msg['To']=gmail
    f=open("b.txt","w")
    f.write(f"Dear {name} \n")
    f.close()
    f=open("b.txt","a")
    f.write(f"\nThe following book(s) have been issued :\nTitle: {Book}\nAccession No.: {Access}\nDue Date : {duedate}\n\nThank you for visiting IITBhilai Library.\n\nNOTE:\nThe Fine Structure will be Following (Which is applicable to all users):\nF1. First 2 week:   2 Rupees will be charged per day.\nF2. After 2 weeks:  100 rupees will be charged per week.\n\n\nLibrary ,IIT Bhilai\nAcademic Building ,Raipur\nMob No-7725932105\n ")
    f.close()
    #determining the total no of entries in the db file
    # with open('count.txt','r') as f:
    #     no=int(f.read())+1
    
    with open("b.txt") as myfile:
        body=myfile.read()
        msg.set_content(body)
    myfile.close()

    server=smtplib.SMTP_SSL("smtp.gmail.com",465)
    server.login("libraryiitbhilai@gmail.com","library@123")
    server.send_message(msg)
    server.quit()


    # response+="<b>  <b>"+name+" and "+roll+"</b><br>"
    

    
   
    # temp1=jsonObj['temp1']
    # temp2=jsonObj['temp2']
    # response+="<b> The input temparatures in Celcius are: <b>"+temp1+" and "+temp2+"</b><br>"
        
    
    
    	    
    return response    


    
    
if __name__ == "__main__":
    app.run(debug=True)
    
    
