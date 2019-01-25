#region Imports
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib, getpass, time, xlrd, datetime, calendar, time
#endregion

dt = datetime.datetime.now()

def doNothing():
    return

def checkWeekend():
    while (datetime.datetime.now().strftime("%A") == "Saturday" or datetime.datetime.now().strftime("%A") == "Sunday"):
        doNothing()

def getAssignments():
    loc = "/Users/yourname/Desktop/WorkFlow.xlsx"
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    worksheet = []
    for i in range(sheet.nrows):
        worksheet.append(sheet.cell_value(i,0) + " : DONE\n")
    return worksheet

def getAndSendWorkflow():
    # create message object instance
    msg = MIMEMultipart()
    
    # setup the parameters of the message
    msg['From'] = "YourEmailHere@test.com"
    password = "YourPasswordHere"
    msg['To'] = "WorkflowEmail@test.com"
    msg['Subject'] = "Workflow " + str(dt.month) + "." + str(dt.day)

    # Reads Excel file and adds each task to a list assignments
    assignments = getAssignments()

    # Loops the list into the message body
    for i in assignments:
        msg.attach(MIMEText(i, 'plain'))

    #create server
    server = smtplib.SMTP('your smtp server here: 999') 
    server.starttls()
    
    # Login Credentials for sending the mail
    server.login(msg['From'], password)
    
    # send the message via the server.
    try:
        server.sendmail(msg['From'], msg['To'], msg.as_string())
    except Exception as e:
        print("ERROR: " + e)
    finally:
        time.sleep(2)
        server.quit()
    print ("Successfully sent email to: " + str((msg['To'])) + " at: " + str(datetime.datetime.now().time()))
    GetDates()

def todayAt(hr, min=0, sec=0, micros=0):
       now = datetime.datetime.now().time()
       now = now.replace(hour=hr, minute=min, second=sec, microsecond=micros)
       now = now.strftime("%H:%M:%S")
       return now

def dateFormat(dateTime):
    now = dateTime
    now = now.strftime("%H:%M:%S")
    return now

def RunMail(sOT):
    while True:
        checkWeekend()
        if (str(dateFormat(datetime.datetime.now().time())) == str(sOT)):
            getAndSendWorkflow()

def GetDates():
    if (dt.strftime("%A") == "Friday"):
        sendOffTime = todayAt(12)
    else:
        sendOffTime = todayAt((17, 30))
    checkWeekend()
    RunMail(sendOffTime)

GetDates()