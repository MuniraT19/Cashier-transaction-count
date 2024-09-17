'''
Creator: Munira Tabassum
Last Update: 09/16/2024
'''
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import datetime
import time
from tkinter import *
from functools import partial
import os
import glob
import pandas as pd
from pandas_ods_reader import read_ods
import sys
from msedge.selenium_tools import Edge
from zipfile import ZipFile
from openpyxl import *

OUTPUT_FILE = r'C:\Users\mtabassum\PycharmProjects\Cashier transaction count\cashier_weekly_transaction.xlsx'

def getInputs():
    MONDAY = datetime.date.today() - datetime.timedelta(days=datetime.date.today().weekday()+7)

    #from datetime import datetime, timedelta
    def cleanup():
        tkWindow.destroy()
    def func(event=None):
        cleanup()

    tkWindow = Tk()
    tkWindow.geometry('400x150')
    tkWindow.title('Cashier Transactions- ACE credentials')

    #username label and text entry box
    usernameLabel = Label(tkWindow, text="User Name").grid(row=0, column=0)
    username = StringVar()
    username.set(os.getlogin())
    usernameEntry = Entry(tkWindow, textvariable=username).grid(row=0, column=1)
    #password label and password entry box
    passwordLabel = Label(tkWindow,text="Password").grid(row=1, column=0)
    password = StringVar()
    passwordEntry = Entry(tkWindow, textvariable=password, show='*').grid(row=1, column=1)

    dateLabel = Label(tkWindow, text="Last Monday ->").grid(row=2, column=0)
    reportingDate = StringVar()
    reportingDate.set(datetime.datetime.strftime(MONDAY, '%m-%d-%Y'))
    dateEntry = Entry(tkWindow, textvariable=reportingDate).grid(row=2, column=1)

    #login button
    tkWindow.bind('<Return>', func)
    loginButton = Button(tkWindow, text="Login", command=cleanup).grid(row=4, column=0, columnspan=2, pady=8)
    tkWindow.mainloop()

    inputUsername= username.get()
    inputPassword= password.get()
    reportDate = reportingDate.get()

    if inputPassword=="" or reportDate=="":
        programNotification("Missing Password or Date Field!")
        quit()

    if "-" in reportDate:
        parts=reportDate.split("-")
    else:
        parts=reportDate.split("\\")
    #Monday date for reporting purposes (temporary as datetime obj to calculate daterange)
    excelDate = datetime.datetime(int(parts[-1]), int(parts[0]),int(parts[1]))

    endDate = excelDate + datetime.timedelta(days=5)
    startDate = endDate - datetime.timedelta(days=20)
    excelDate = datetime.datetime.strftime(excelDate, '%m-%d-%Y')

    return startDate, excelDate, endDate, inputUsername, inputPassword

#-------- FUNCTIONS BEGIN -----------------------------------------------------------
#try to click an HTML element, throw an error if it can not find this item after 40 sec
def tryXPath(itemName, driver):
    answer=1
    begin = time.time() + 200
    item=0
    while answer and (begin > time.time()):
        try:
            item = driver.find_element(By.XPATH, itemName)
            answer=0
        except: print("")

    if item==0:
        programNotification("Program Error! Could not find " + itemName)
        driver.quit()
    return item

def tryLink(itemName, driver):
    answer=1
    begin = time.time()+200
    item=0
    while answer and (begin > time.time()):
        try:
            item = driver.find_element(By.LINK_TEXT, itemName)
            answer=0
        except: print("")

    if item==0:
        driver.quit()
        programNotification("Program Error! Could not find " + itemName)
    return item
#------------------------------------------------------------------------------------
def login(driver):
    time.sleep(2)
    #LOGIN PROCESS
    username= tryXPath("//input[@type='text']", driver)
    username.send_keys(inputUsername) #Recieves login/password from GUI Above
    password=tryXPath("//input[@type='password']", driver)
    password.send_keys(inputPassword)
    login=tryXPath("//button[@type='button']", driver)
    login.click()

#Open Reports--------------------------------------------------------------------------
def chooseReport(reportName, waitTime, begin, end, driver):

    tryXPath("//span[contains(text(), 'Report')]", driver).click()
    tryXPath("//span[contains(text(), 'All Reports')]", driver).click()
    tryXPath("//span[@id='caption2_d-p']", driver).click()
    #Search for Report Name
    time.sleep(3)
    #filter = tryXPath("//input[@placeholder='Filter   ']")
    filter = tryXPath("//input[@id='d-q']", driver)
    filter.send_keys(reportName)
    filter.send_keys(Keys.RETURN)
    time.sleep(4)
    generate = tryXPath("//td[@headers='d-7-CH']", driver)
    generate.click() #Click Generate Button

    #INPUT REPORTING DATE
    startDate = tryXPath("//input[@name='k-6']", driver)
    startDate.send_keys(begin)
    endDate = tryXPath("//input[@name='k-7']", driver)
    endDate.send_keys(end)
    endDate.send_keys(Keys.RETURN)

    tryXPath("//span[contains(text(), 'Yes')]", driver).click()
    #Waiting for report to load then click CUBE
    time.sleep(waitTime)
    tryXPath("//span[contains(text(), 'Cube')]", driver).click()

#----------------------------------------------------------------------------------------
def chooseView(viewName, driver):
    time.sleep(5)
    while True:
        try:
            #Select View which is passed in on function call, see line 144 and on
            findView = tryXPath("//span[contains(text(), 'Views')]", driver)
            findView.click()
            tryLink(viewName, driver).click()
            break
        except:
            pass
    #Select View which is passed in on function call, see line 144 and on    
    #pickView.click()
    time.sleep(4)
    passed=False
    while not passed:
        try:
            tryLink('Export', driver).click()
            #Export to spreadsheet
            tryXPath("//a[@title='Menu']", driver).click()
            passed=True
        except: continue
    tryXPath("//span[contains(text(), 'Spreadsheet')]", driver).click()
#-----------FILES--------------------------------FILES------------------------------------
def changeFileName(newFileName):
    folder_path = "C:\\Users\\" + inputUsername + "\\Downloads\\*.ods"
    files = glob.glob(folder_path)

    while True:
        currList = glob.glob(folder_path)
        max_file = list(set(files) - set(currList)) + list(set(currList) - set(files))
        #print(len(max_file))
        if len(max_file) > 0:
            break
        time.sleep(1)
    downloadDF = read_ods(max_file[0])
    downloadDF.to_excel("C:\\Users\\" + inputUsername + "\\Downloads\\"+newFileName+".xlsx", index=False)
#-----------------------------------------------------------------------------------------
def checkPassword(driver):
    time.sleep(2)
    #Check for a message that says password is almost expired
    try:
        driver.find_element_by_xpath("//span[contains(text(), 'OK')]").click()
    except: print("")
#-----------------------------------------------------------------------------------------
def programNotification(message):
    def cleanup():
        tkWindow.destroy()
    def func(event=None):
        cleanup()

    tkWindow = Tk()
    tkWindow.geometry("500x500")
    text = Text(tkWindow, height=100, width=100)
    text.pack()
    text.insert(END, message)
    tkWindow.mainloop()
#-----------BEGIN Execution-------------------------------------------------------------------
#-----------BEGIN-----------------------------------------------------------------------------
if __name__ == "__main__":
    try:
        wb = load_workbook(OUTPUT_FILE)
        wb.save(OUTPUT_FILE) #Have to save to check if it is truly in USE
        
    except:
        programNotification("OPS Statistics is currently open by someone else!")
        quit()

    #Ask user for input
    reportDate, excelDate, saturday, inputUsername, inputPassword = getInputs()
    
    #Setup Selenium driver and launch ACE Webpage
    driver = Edge(r"L:\Team\TRE-Team-Automation\Development\Driver\msedgedriver.exe")
    ace = "https://ace-staging.arlingtonva.local/ACS/Tkh9XmzW/#2"
    driver.get(ace)
    time.sleep(2)
    # #First login, check password and throw popup error if WRONG
    try: login(driver)
    except: 
        programNotification("Password is incorrect! Please try again.")
        quit()
    checkPassword(driver) #Check for message - Password expires in 15 days

    chooseReport("Payment Statistics", 10, datetime.datetime.strftime(reportDate, '%m-%d-%Y'), 
                                            datetime.datetime.strftime(saturday, '%m-%d-%Y'), driver)
    chooseView('A_CashierProj', driver)
    changeFileName('OPS_Cashier_Report')

# -------------- WEB SCRAPING COMPLETE ----------------------------------------------------

    df = pd.read_excel("C:\\Users\\" + inputUsername + "\\Downloads\\OPS_Cashier_Report.xlsx", skiprows=[-1])
    database = pd.read_excel(OUTPUT_FILE, sheet_name="Payments")

    # outputDF = []
    # #LATER FILTER ON Deposit PAYMENT DATE
    # df['Deposit'] =  pd.to_datetime(df['Deposit'], format='%d-%b-%Y')
    # #Cast every date in Deposit to the Monday date
    # df = df[(df['Deposit']>= reportDate) & (df['Deposit']<=saturday)]
    # df['Deposit'] = df['Deposit'].apply(lambda x: x - datetime.timedelta(days=x.weekday()))

    # #Group all data by Deposit
    # df = df.groupby(["Deposit","Source","Account Type","Payment Channel"]).agg({'Count':'sum','Amount':'sum'}).reset_index()
    # #df.to_excel("Testing.xlsx")

    # cashier = df[((df['Source'] == 'WACHOVIA CREDIT CARD (027)') |  (df['Source'] == 'CASHIERING CHECK & CASH (062)')) & (df['Account Type'] != 'CRIF Payment') & (df['Payment Channel'] != "Batch")]
    # crif = df[((df['Source'] == 'WACHOVIA CREDIT CARD (027)') |  (df['Source'] == 'CASHIERING CHECK & CASH (062)')) & (df['Account Type'] == 'CRIF Payment')]
    # wire_edi = df[((df['Source'] == 'WACHOVIA CREDIT CARD (027)') |  (df['Source'] == 'CASHIERING CHECK & CASH (062)')) & (df['Account Type'] != 'CRIF Payment') & (df['Payment Channel'] == "Batch")]
    # mortgage = df[df['Source'] == 'MORTGAGE COMPANY (MRG)']
    # tax_service = df[df['Source'] == 'TAX SERVICE (TXS)'] 
    # fleet = df[df['Source'] == 'FLEET (FLE)']
    # opc = df[df['Source'] == 'OPC (081)']
    # merkle = df[(df['Source'] == 'WACHOVIA BUSINESS TAX (73B)') |  (df['Source'] == 'WACHOVIA DELINQUENT (73Q)') | (df['Source'] == 'WACHOVIA RETAIL (073)') | (df['Source'] == 'WACHOVIA WHOLESALE (73W)')]

    # abd = df[df['Source'] == 'AUTOMATED BANK DEBIT (ABD)']
    # echeck = df[df['Source'] == 'WACHOVIA CAPP ECHECK (29C)']
    # creditCard = df[df['Source'] == 'PAYMENT PORTAL CRED CD (089)']

    # checkfree = df[df['Source'] == 'CHECKFREE (CKF)']
    # ebox =  df[df['Source'] == 'Wells Fargo E-Box (EBX)']
    # sod = df[df['Source'].str.contains('(SD1)')]
    # ncc = df[df['Source'] == 'NATIONWIDE CREDIT CORP (NCC)']
    # #police = df[df['Source'] == 'POLICE CASHIERING CASH/CHECK/CARDS (075)'] 
    
    # typesDict = {"Cashier":cashier,"CRIF":crif,"Wire_EDI":wire_edi,
    #              "Mortgage":mortgage,"Tax Service":tax_service,"Fleet":fleet,
    #              "OPC_IVR":opc,"Merkle":merkle,"ABD":abd,"E-Check":echeck,
    #              "Credit Card":creditCard,"Checkfree":checkfree,"E-Box":ebox,
    #              "SOD":sod,"NCC":ncc}

    # #Cycle through each subdf in typesDict to append to the database DF
    # for item in typesDict:
    #     currdf = typesDict[item]
    #     currdf["Category"] = item

    #     currdf = currdf.groupby(["Deposit","Category"]).agg({'Count':'sum','Amount':'sum'}).reset_index()
    #     currdf.columns=['Week', 'Category', 'Count', '$_Amount']
    #     database = database.append(currdf)

    # database = database.sort_values("Count").groupby(["Week","Category"], as_index=False).max()
    # #print(database.tail(20))
    # with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    #     database.to_excel(writer, sheet_name='Payments', index=False)

    # driver.close()
    # programNotification("Data is downloaded! \n*Open OPS Statistics Report to review.")
