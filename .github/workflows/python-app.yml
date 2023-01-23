import requests as rq
import json
# import pyodbc
# import pandas as pd
import openpyxl
import getpass

username = getpass.getuser()
# connection = pyodbc.connect("C://Users//"+username+"//Desktop//LoanSphere.xlsx")
# script = """SELECT URL FROM Sheet1 where TC_ID='TC_1'"""
# result = pd.read_sql_query(script,connection)
# print(result)
r = rq
book = openpyxl.load_workbook("C://Users//" + username + "//Desktop//LoanSphere.xlsx")
sheet = book.get_sheet_by_name("Business")
sheet2 = book.get_sheet_by_name("TestData")
execute_Sheet = book.get_sheet_by_name("ExecutionSheet")
# bb = execute_Sheet.cell(row=2,column=6).value
# print (bb)
# for i in range(bb):
for i in range(184):
    if execute_Sheet.cell(row=i + 2, column=5).value == "Y":
        # url = sheet.cell(row=i+2, column = 2).value

        creditAccount = str(sheet2.cell(row=i + 2, column=2).value)
        LoanNumber = sheet.cell(row=i + 2, column=3).value
        url = "https://bkfsqa.phhsvcs.com/cmentws/v1/loans/payments/eligibility?loanNumber=" + str(
            creditAccount) + "&debuggingTool=true"
        print(url)
        creditAccount = str(sheet2.cell(row=i + 2, column=2).value)
        print(creditAccount)
        amountOwed = str(sheet2.cell(row=i + 2, column=3).value)
        nextPaymentDueDateMMDDYY = str(sheet2.cell(row=i + 2, column=4).value)
        mis2 = str(sheet2.cell(row=i + 2, column=5).value)
        badCheckStop = str(sheet2.cell(row=i + 2, column=6).value)
        creditFirstName = str(sheet2.cell(row=i + 2, column=7).value)
        debitFullName1 = str(sheet2.cell(row=i + 2, column=8).value)
        creditLastName = str(sheet2.cell(row=i + 2, column=9).value)
        debitFullName2 = str(sheet2.cell(row=i + 2, column=10).value)
        mis13 = str(sheet2.cell(row=i + 2, column=11).value)
        socialLastFour = str(sheet2.cell(row=i + 2, column=12).value)
        ssNum = str(sheet2.cell(row=i + 2, column=13).value)
        debitFullName2co = str(sheet2.cell(row=i + 2, column=14).value)
        mis9 = str(sheet2.cell(row=i + 2, column=15).value)
        debitAddress1 = str(sheet2.cell(row=i + 2, column=16).value)
        debitCity = str(sheet2.cell(row=i + 2, column=17).value)
        debitState = str(sheet2.cell(row=i + 2, column=18).value)
        debitZip = str(sheet2.cell(row=i + 2, column=19).value)
        mis5 = str(sheet2.cell(row=i + 2, column=20).value)
        fcStopCode = str(sheet2.cell(row=i + 2, column=21).value)
        paymentInFullStopCode = str(sheet2.cell(row=i + 2, column=22).value)
        mis8 = str(sheet2.cell(row=i + 2, column=23).value)
        mis14 = str(sheet2.cell(row=i + 2, column=24).value)
        mis11 = str(sheet2.cell(row=i + 2, column=25).value)
        mis12 = str(sheet2.cell(row=i + 2, column=26).value)
        mis10 = str(sheet2.cell(row=i + 2, column=27).value)
        mis4 = str(sheet2.cell(row=i + 2, column=28).value)
        mis15 = str(sheet2.cell(row=i + 2, column=29).value)
        mis16 = str(sheet2.cell(row=i + 2, column=30).value)
        mis17 = str(sheet2.cell(row=i + 2, column=31).value)
        mis18 = str(sheet2.cell(row=i + 2, column=32).value)
        mis19 = str(sheet2.cell(row=i + 2, column=33).value)
        mis20 = str(sheet2.cell(row=i + 2, column=34).value)

        firstPrincipalBalance = str(sheet2.cell(row=i + 2, column=35).value)
        currMinPaymentAmount = str(sheet2.cell(row=i + 2, column=36).value)
        pymtDateDiff = str(sheet2.cell(row=i + 2, column=37).value)

        naConsentSts = str(sheet2.cell(row=i + 2, column=38).value)
        requestSystem = str(sheet2.cell(row=i + 2, column=39).value)
        bkStatus = str(sheet2.cell(row=i + 2, column=40).value)
        planType = str(sheet2.cell(row=i + 2, column=41).value)
        usr3pos3cxx = str(sheet2.cell(row=i + 2, column=42).value)
        typeAcquisitionCode = str(sheet2.cell(row=i + 2, column=43).value)
        acquisitionDate = str(sheet2.cell(row=i + 2, column=44).value)
        firstDueDate = str(sheet2.cell(row=i + 2, column=45).value)
        correspondentNumber = str(sheet2.cell(row=i + 2, column=46).value)
        lnBkrCd = str(sheet2.cell(row=i + 2, column=47).value)
        loType = str(sheet2.cell(row=i + 2, column=48).value)
        todaysDate = str(sheet2.cell(row=i + 2, column=49).value)
        originalMortgageAmount = str(sheet2.cell(row=i + 2, column=50).value)
        usr1Pos7bxx = str(sheet2.cell(row=i + 2, column=51).value)
        highPricedInd = str(sheet2.cell(row=i + 2, column=52).value)
        loanPurposeCode = str(sheet2.cell(row=i + 2, column=53).value)
        dd = str(sheet2.cell(row=i + 2, column=56).value)
        rAPI = """{ 
            "creditAccount":""" + dd + creditAccount + dd + """,
            "amountOwed":""" + dd + amountOwed + dd + """,
            "nextPaymentDueDateMMDDYY":""" + dd + nextPaymentDueDateMMDDYY + dd + """,
            "mis2":""" + dd + mis2 + dd + """,
            "badCheckStop":""" + dd + badCheckStop + dd + """,
            "creditFirstName":""" + dd + creditFirstName + dd + """,
            "debitFullName1":""" + dd + debitFullName1 + dd + """,
            "creditLastName":""" + dd + creditLastName + dd + """,
            "debitFullName2":""" + dd + debitFullName2 + dd + """,
            "mis13":""" + dd + mis13 + dd + """,
            "socialLastFour":""" + dd + socialLastFour + dd + """,
            "ssNum":""" + dd + ssNum + dd + """,
            "debitFullName2co":""" + dd + debitFullName2co + dd + """,
            "mis9":""" + dd + mis9 + dd + """,
            "debitAddress1":""" + dd + debitAddress1 + dd + """,
            "debitCity":""" + dd + debitCity + dd + """,
            "debitState":""" + dd + debitState + dd + """,
            "debitZip":""" + dd + debitZip + dd + """,
            "mis5":""" + dd + mis5 + dd + """,
            "fcStopCode":""" + dd + fcStopCode + dd + """,
            "paymentInFullStopCode":"""+ dd + paymentInFullStopCode + dd + """,
            "mis8":""" + dd + mis8 + dd + """,
            "mis14":""" + dd + mis14 + dd + """,
            "mis11":""" + dd + mis11 + dd + """,
            "mis12":""" + dd + mis12 + dd + """,
            "mis10":""" + dd + mis10 + dd + """,
            "mis4":""" + dd + mis4 + dd + """,
            "mis15":""" + dd + mis15 + dd + """,
            "mis16":""" + dd + mis16 + dd + """,
            "mis17":""" + dd + mis17 + dd + """,
            "mis18":""" + dd + mis18 + dd + """,
            "mis19":""" + dd + mis19 + dd + """,
            "mis20":""" + dd + mis20 + dd + """,
            "firstPrincipalBalance":""" + dd + firstPrincipalBalance + dd + """,
            "currMinPaymentAmount":""" + dd + currMinPaymentAmount + dd + """,
            "pymtDateDiff":""" + dd + pymtDateDiff + dd + """,
            "naConsentSts":""" + dd + naConsentSts + dd + """,
            "requestSystem":""" + dd + requestSystem + dd + """,
            "bkStatus":""" + dd + bkStatus + dd + """,
            "planType":""" + dd + planType + dd + """,
            "usr3pos3cxx":""" + dd + usr3pos3cxx + dd + """,
            "typeAcquisitionCode":""" + dd + typeAcquisitionCode + dd + """,
            "acquisitionDate":""" + dd + acquisitionDate + dd + """,
            "firstDueDate":""" + dd + firstDueDate + dd + """,
            "correspondentNumber":""" + dd + correspondentNumber + dd + """,
            "lnBkrCd":""" + dd + lnBkrCd + dd + """,
            "loType":""" + dd + loType + dd + """,
            "todaysDate":""" + dd + todaysDate + dd + """,
            "originalMortgageAmount":""" + dd + originalMortgageAmount + dd + """,
            "usr1Pos7bxx":""" + dd + usr1Pos7bxx + dd + """,
            "highPricedInd":""" + dd + highPricedInd + dd + """,
            "loanPurposeCode":""" + dd + loanPurposeCode + dd + """,
            "section184Ind":""" + dd + dd + """,
            "tribalAffiliation":""" + dd + dd + """
            }"""

        print(rAPI)

        c3 = sheet.cell(row=i + 2, column=3)
        c3.value = rAPI
        a = r.post(url, data=rAPI,
                   headers={'Content-Type': 'application/json', 'user': 'svcsBkfsQA', 'password': 'yX#u3&cZh!'},
                   verify=False)
        # d = ssl.SSLContext()

        # requestAPI = sheet.cell(row=2, column=3).value
        # print (requestAPI)
        # url="https://bkfsqa.phhsvcs.com/cmentws/v1/loans/123/1234567890/testclients/payments/eligibility"
        # https://bkfsqa.phhsvcs.com/cmentws/v1/loans/payments/eligibility?loanNumber=7241926687&debuggingTool=true
        # https://bkfsqa.phhsvcs.com/cmentws/v1/loans/{clientId}/{loanNumber}/testclients/payments/eligibility"
        # c3 = sheet.cell(row=i+2, column = 3)
        # c3.value= rAPI
        # Parameters = {'PaymentEffectiveDate':'2000-01-01'}
        # a = r.get(url,headers = {'Content-Type':'application/json','user':'SvcsBkfsQA', 'password':'yX#u3&cZh!'},params = Parameters, verify = False)
        # a = r.post(url,data=rAPI,headers = {'ContentType':'application/json','user': 'svcsBkfsQA', 'password':'yX#u3&cZh!'}, verify = False)
        # d = ssl.SSLContext()
        c1 = sheet.cell(row=i + 2, column=4)
        c1.value = a.status_code
        c2 = sheet.cell(row=i + 2, column=5)
        c2.value = a.text
        book.save("C://Users//" + username + "//Desktop//LoanSphere.xlsx")
        # a=r.get(url)
        print(a.status_code)
        response = a.text
        print(response)
    else:
        print("not included")
