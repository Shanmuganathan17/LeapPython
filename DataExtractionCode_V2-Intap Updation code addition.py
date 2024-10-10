import logging
import requests
import urllib3
import json
import pandas as pd
import os
import getpass
import re
import math
import warnings
import openpyxl
from datetime import datetime, date
import win32com.client as win32
import socket
import sys
import traceback
import requests
import json
import pandas as pd
import win32gui, win32con


warnings.filterwarnings('ignore')
urllib3.disable_warnings()
global ExcelAPI
global path
global globalInbox
response_Status = 0
ExcelAPI = {}
globalInbox = {}
mailEnabled = 0
count = {}
caseId = []
lstDataSetData = []
countMatchFlag = False;
#configPath = "D:\\LeapPython\\config\\DatasetExtractionAPIList_PostUATChanges.xlsx"



path =os.path.join(os.getcwd(), 'config', 'data.json')

print(path)
with open(path) as config:
    enate_config = json.load(config)
    print(enate_config)
    config.close()

configPath=enate_config['path']
logFolderPath =enate_config['logpath']
mailEnabled=enate_config['mailenable']
canIdDigit=enate_config['candidateIdLen']
appIdDigit=enate_config['applicationIdLen']
if not os.path.exists(logFolderPath):
    os.makedirs(logFolderPath)
print("Config Path- "+ configPath)
file_name = str(datetime.now().strftime("%d %b %Y"))+' IntapUpdation Log File'
logging.basicConfig(level=logging.INFO,filename=f'{logFolderPath}\\{file_name}.log', format='%(asctime)s INFO %(message)s', datefmt='%y/%m/%d %H:%M:%S')
datevar=str(datetime.now().strftime("%d %b %Y "))
subject = f"RPA- Intap Updation BOT Status - on Date : {datevar} on {socket.gethostname()} machine"
#recipients = ["shanmuganathan.s@infosys.com", "sangtrash.ansari@infosys.com"]
recipients=enate_config['recipients'].split(',')

def send_email(subject, body, recipients, attachments=None):
    if (mailEnabled == 1):
        try:
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)  # 0 represents an email item
            # Set email properties
            mail.Subject = subject
            mail.Body = body
            mail.To = ";".join(recipients)
            mail.Sensitivity = 2
            # Attach files if specified
            if attachments:
                for attachment in attachments:
                    mail.Attachments.Add(attachment)
            # Send the email
            #mail.Display(False)
            mail.Send()
            #app = pywinauto.Application(backend="uia").connect(title="Microsoft Azure Information Protection")
            #app.window(title="Confirm and Send").Save.click()
        except Exception as e:
            logging.error("ERROR in sent mail event : " + str(e))
            print(str(e))
            print(traceback.format_exc())
    else:
        logging.info("Mail application is not enabled. Hence, we cannot send mails")


def APIExcel(path):  # storing excel config values in to dict
    try:
        df = pd.read_excel(path)
        ExcelAPI['AuthURL'] = df.loc[0][1]
        ExcelAPI['Payload'] = df.loc[1][1]
        ExcelAPI['InboxcountAPI'] = df.loc[2][1]
        ExcelAPI['DataCountAPI'] = df.loc[3][1]
        ExcelAPI['DataSetAPI'] = df.iloc[4][1]
        ExcelAPI['GlobalCaseAPI'] = df.iloc[5][1]
        ExcelAPI['intapAuthorizationUrl'] = df.loc[6][1]
        ExcelAPI['proposeOfferUrl'] = df.loc[7][1]
        ExcelAPI['getCandidateDataAPI'] = df.loc[8][1]
        ExcelAPI['getCategoryURL'] = df.loc[9][1]
        ExcelAPI['strGetRemoteWorkLocationURL'] = df.iloc[10][1]
        ExcelAPI['strGetEduationCategoryURL'] = df.iloc[11][1]
        ExcelAPI['strGetPuCodeURL'] = df.iloc[12][1]
        ExcelAPI['strGetLOPCodeURL'] = df.loc[13][1]
        ExcelAPI['strGetRoleCodeURL'] = df.loc[14][1]
        ExcelAPI['strGetCandidateDataAPI'] = df.loc[15][1]
        ExcelAPI['postSalaryUpdateAPI'] = df.iloc[16][1]
        ExcelAPI['leapStatusUpdateUrl'] = df.iloc[17][1]
        ExcelAPI['getEducationConsideredList']=df.iloc[18][1]
        ExcelAPI['inputExcelPath'] = df.iloc[19][1]
        return ExcelAPI
    except Exception as e:
        logging.error("ERROR in Config File Read Method : " + str(e))
        print(str(e))
        print(traceback.format_exc())

APIExcel(configPath)
# APIExcel('C:\\leapPython\\LateralOffer\\Config\\Attachment API (2).xlsx')

authUrl = ExcelAPI['AuthURL']
payload = ExcelAPI['Payload']
InboxCountUrl = ExcelAPI['InboxcountAPI']
DataCountAPI = ExcelAPI['DataCountAPI']
DataSetAPI = ExcelAPI['DataSetAPI']
GlobalCaseUrl = ExcelAPI['GlobalCaseAPI']
intapAuthorizationUrl = ExcelAPI['intapAuthorizationUrl']
proposeOfferUrl = ExcelAPI['proposeOfferUrl']
getCandidateDataAPI = ExcelAPI['getCandidateDataAPI']
getCategoryURL = ExcelAPI['getCategoryURL']
strGetRemoteWorkLocationURL = ExcelAPI['strGetRemoteWorkLocationURL']
strGetEduationCategoryURL = ExcelAPI['strGetEduationCategoryURL']
strGetPuCodeURL = ExcelAPI['strGetPuCodeURL']
strGetLOPCodeURL = ExcelAPI['strGetLOPCodeURL']
strGetRoleCodeURL = ExcelAPI['strGetRoleCodeURL']
strGetCandidateDataAPI = ExcelAPI['strGetCandidateDataAPI']
postSalaryUpdateAPI = ExcelAPI['postSalaryUpdateAPI']
leapStatusUpdateUrl = ExcelAPI['leapStatusUpdateUrl']
strGetEducationConsideredListAPI =ExcelAPI['getEducationConsideredList']
inputExcelPath = ExcelAPI['inputExcelPath']

# Leap URL Authorization
def leapAuthorization():
    bearerToken=None
    try:
        headers = {
            'Content-Type': 'text/plain',
            'Cookie': 'XSRF-TOKEN=21bbf30b-6a99-414d-9aef-dfc16045ab3c'
        }
        response = requests.request("POST", authUrl, headers=headers, data=payload)
        response_Status = response.status_code
        logging.info("Leap Authorization Response Status : " + str(response_Status))
        if response_Status == 200:
            bearerToken = response.json()['id_token']
            logging.info("Leap Authorization Token :" +str(bearerToken))
        else:
            bearerToken=None
            logging.error("ERROR in Leap Authorization Connection : status is not 200")
            body = """
    Hi,
    
        Found authorization issue in leap. So, kindly check and work on it.
    
    Thanks,
    RPA Script"""
            # send_email(subject, body, recipients)
    except Exception as e:
        bearerToken=None
        body = """
    Hi,
    
        Found authorization issue in leap. So, kindly check and work on it.
    
    Thanks,
    RPA Script"""
        # send_email(subject, body, recipients)
        print(str(e))
        print(traceback.format_exc())
        logging.error("ERROR in Leap Authorization Connection : " + str(e))
    return bearerToken


# Leap case status Update
def statusUpdateInLeap(URL, strLeapToken, strStatus, strReasonCode, strComments, strBID, strCaseId):
    statusUpdateResult = False
    try:
        strUrl = URL.format(strCaseId)
        payload = json.dumps({
            "contents": {
                "status_": strStatus
            },
            "variables": {
                "reasonCode": strReasonCode,
                "comments": strComments,
                "BID": strBID
            }
        })
        headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer {}'.format(strLeapToken)
        }
        response = requests.request("POST", strUrl, headers=headers, data=payload)
        print(response.text)
        logging.info("Leap case status update api response statuscode :" + str(response.status_code))
        if response.status_code == 200:
            statusUpdateResult = True
        else:
            statusUpdateResult = False

    except Exception as e:
        print(str(e))
        print(traceback.format_exc())
        statusUpdateResult = False
        logging.error("ERROR in Leap status update api fn call: " + str(e))
    return statusUpdateResult


# Intap Authorization
def intapAuthorization(authorizationUrl):
    bearer_token = None
    try:
        postData = {"username": "intapbatchuser", "password": "intapbatchuser@12345", 'grant_type': 'password',
                    'client_id': "infyclient"}
        x = requests.post(authorizationUrl, data=postData)
        # print(x.text)
        logging.info("Intap Authorization API Response status code :" + str(int(x.status_code)))
        if x.status_code == 200:
            tokenResult = json.loads(x.text)
            # print(type(tokenResult))
            bearer_token = tokenResult["access_token"]
            print(bearer_token)
            logging.info("Intap Authorization Token :" + str(bearer_token))
        else:
            bearer_token = None
    except requests.exceptions.ConnectionError:
        connectionStatus = "Connection refused"
        print(connectionStatus)

    except Exception as e:
        print(str(e))
        print(traceback.format_exc())
        bearer_token = None
        logging.error("ERROR in Intap Authorization fn : " + str(e))
    return bearer_token


# Application ID Check Fn Call
def applicationIDCheckAPIfn(applicationIDCheckAPI, strToken):
    applicationIdLst = []
    try:
        url = applicationIDCheckAPI
        payload = ""
        headers = {
            'Authorization': 'Bearer {}'.format(strToken)
        }
        response = requests.request("GET", url, headers=headers, data=payload)
        #print(response.text)
        logging.info("GET Application ID API Response status code:" + str(int(response.status_code)))
        if response.status_code == 200:
            candidateDetails = json.loads(response.text)
            print(len(candidateDetails))
            for i in candidateDetails:
                applicationIdLst.append(i["applicationId"])
    except Exception as e:
        print(str(e))
        print(traceback.format_exc())
        logging.error("ERROR in Application ID Check fn : " + str(e))
    return applicationIdLst


# Get Category Code Fn Call
def getCategoryCode(getCategoryURL, strToken, strCategory):
    strCategoryCode = None
    try:
        payload = {}
        headers = {
            'Authorization': 'Bearer {}'.format(strToken)
        }
        response = requests.request("GET", getCategoryURL, headers=headers, data=payload)
        logging.info("GET Category Code API Response status code :" + str(int(response.status_code)))
        if response.status_code == 200:
            # print(response.text)
            categoryCodeDetails = json.loads(response.text)
            # print(len(categoryCodeDetails))
            for i in categoryCodeDetails:
                if strCategory.lower() == i["value"].lower():
                    strCategoryCode = i["key"]
                    print("category Code -" + strCategoryCode)
        logging.info("category Code value :" + str(strCategoryCode))
    except Exception as e:
        print(str(e))
        print(traceback.format_exc())
        strCategoryCode = None
        logging.error("ERROR in get Category code fn : " + str(e))
    return strCategoryCode


# get remote work location code fn call
def getRemoteWorkLocationCode(strGetRemoteWorkLocationURL, strToken, strRemoteWrkLocationName):
    strRemoteWorkLocationCode = None
    try:
        strRemoteWrkLocationURL = strGetRemoteWorkLocationURL.format(strRemoteWrkLocationName)
        payload = {}
        headers = {
            'Authorization': 'Bearer {}'.format(strToken)
        }
        response = requests.request("GET", strRemoteWrkLocationURL, headers=headers, data=payload)
        logging.info("GET Remote work location API Response code :" + str(int(response.status_code)))
        if response.status_code == 200:
            # print(response.text)
            categoryCodeDetails = json.loads(response.text)
            # print(len(categoryCodeDetails))
            for i in categoryCodeDetails:
                if strRemoteWrkLocationName.lower() in i["city_state_country"].lower():
                    strRemoteWorkLocationCode = i["city_id"]
                    print("Remotework Location Code " + str(strRemoteWorkLocationCode))
                    break
        logging.info("Remote work location Code value :" + str(strRemoteWorkLocationCode))
    except Exception as e:
        print(str(e))
        print(traceback.format_exc())
        strRemoteWorkLocationCode = None
        logging.error("ERROR in GET remote work location fn : " + str(e))
    return strRemoteWorkLocationCode


# get education category code fn call
def getEducationCategryCode(strGetEduationCategoryURL, strToken, strHighestQualification):
    strEducationCategoryCode = None
    try:
        payload = {}
        headers = {
            'Authorization': 'Bearer {}'.format(strToken)
        }
        response = requests.request("GET", strGetEduationCategoryURL, headers=headers, data=payload)
        logging.info("GET Education Category Code API Response code :" + str(int(response.status_code)))
        if response.status_code == 200:
            # print(response.text)
            educationCategoryCodeDetails = json.loads(str(response.text))
            # print(educationCategoryCodeDetails['fields'][0]['options'])
            educationCategoryOption = educationCategoryCodeDetails['fields'][0]['options']
            # educationCategoryList=json.loads(educationCategoryOption)
            for i in educationCategoryOption:
                if strHighestQualification.lower() == i["value"].lower():
                    strEducationCategoryCode = i["key"]
                    print("Education Category Code -" + strEducationCategoryCode)
        logging.info("Education category Code value :" + str(strEducationCategoryCode))
    except Exception as e:
        print(str(e))
        print(traceback.format_exc())
        strEducationCategoryCode = None
        logging.error("ERROR in GET Education category code fn : " + str(e))
    return strEducationCategoryCode


# get education category code fn call
def getEducationConsideredCode(strGetEduationConsideredURL, strToken, strEducationQualification):
    strEducationConsideredCode = None
    try:
        payload = {}
        headers = {
            'Authorization': 'Bearer {}'.format(strToken)
        }
        response = requests.request("GET", strGetEduationConsideredURL, headers=headers, data=payload)
        logging.info("GET Education Category Code API Response code :" + str(int(response.status_code)))
        if response.status_code == 200:
            # print(response.text)
            educationConsideredCodeDetails = json.loads(str(response.text))
            # print(educationCategoryCodeDetails['fields'][0]['options'])
            #print(len(educationConsideredCodeDetails))
            #educationConsideredOption = educationConsideredCodeDetails['fields'][0]['options']
            # educationCategoryList=json.loads(educationCategoryOption)
            for i in range(len(educationConsideredCodeDetails)):
                #print(educationConsideredCodeDetails[i]['value'])
                if strEducationQualification.lower() == educationConsideredCodeDetails[i]['value'].lower():
                    strEducationConsideredCode = educationConsideredCodeDetails[i]['key']
                    print("Education Category Code -" + strEducationConsideredCode)
                    break
        logging.info("Education category Code value :" + str(strEducationConsideredCode))
    except Exception as e:
        print(str(e))
        print(traceback.format_exc())
        strEducationConsideredCode = None
        logging.error("ERROR in GET Education category code fn : " + str(e))
    return strEducationConsideredCode



# get pu code fn call
def getPUCode(strUrl, strToken, strpuUnit):
    strPuCode = None
    try:
        payload = {}
        headers = {
            'Authorization': 'Bearer {}'.format(strToken)
        }
        response = requests.request("GET", strUrl, headers=headers, data=payload)
        logging.info("GET PU Code API Response code :" + str(int(response.status_code)))
        if response.status_code == 200:
            # print(response.text)
            puCodeOption = json.loads(str(response.text))

            for i in puCodeOption:
                if strpuUnit.lower() in i["value"].lower():
                    strPuCode = (i["key"])
                    print("Pu Code -" + strPuCode)
                    break
        logging.info("PU Code value :" + str(strPuCode))
    except Exception as e:
        print(str(e))
        print(traceback.format_exc())
        strPuCode = None
        logging.error("ERROR in GET PU Code fn : " + str(e))
    return strPuCode

# get LOP code fn call
def getLOPCode(strUrl, strToken, strLOP):
    strLOPCode = None
    try:
        payload = {}
        headers = {
            'Authorization': 'Bearer {}'.format(strToken)
        }
        response = requests.request("GET", strUrl, headers=headers, data=payload)
        logging.info("GET LOP API Response code :" + str(int(response.status_code)))
        if response.status_code == 200:
            # print(response.text)
            puCodeOption = json.loads(str(response.text))

            for i in puCodeOption:
                if strLOP.lower() in i["value"].lower():
                    strLOPCode = (i["key"])
                    print("str Lop Code -" + strLOPCode)
                    break
        logging.info("LOP code value :" + str(strLOP))
    except Exception as e:
        print(str(e))
        print(traceback.format_exc())
        strLOPCode = None
        logging.error("ERROR in GET Lop Code FN : " + str(e))
    return strLOPCode


# get role desc code fn call
def getRoleDesCode(strGetRoleCodeURL, strToken, strRoleDesignation):
    strRoleCode = None
    try:
        payload = {}
        headers = {
            'Authorization': 'Bearer {}'.format(strToken)
        }
        response = requests.request("GET", strGetRoleCodeURL, headers=headers, data=payload)
        logging.info("GET ROLE Code API Response code :" + str(int(response.status_code)))
        if response.status_code == 200:
            # print(response.text)
            puCodeOption = json.loads(str(response.text))
            for i in puCodeOption:
                if strRoleDesignation.lower()==i["value"].lower():
                    strRoleCode = (i["key"])
                    print("Role Code " + strRoleCode)
                    break
        logging.info("role Code value  :" + str(strRoleCode))
    except Exception as e:
        print(str(e))
        print(traceback.format_exc())
        strRoleCode = None
        logging.error("ERROR in GET Role Code Fn : " + str(e))
    return strRoleCode


# get account code fn call
def getAccountCode(GetCandidateDataAPI, strToken, applicationId, strAccountName):
    strAccountCode = None
    try:
        payload = {}
        headers = {
            'Authorization': 'Bearer {}'.format(strToken)
        }
        response = requests.request("GET", GetCandidateDataAPI, headers=headers, data=payload)
        logging.info("GET Account name code API Response code :" + str(int(response.status_code)))
        if response.status_code == 200:
            # print(response.text)
            accountNameOption = json.loads(str(response.text))
            accountNameFields = accountNameOption["fields"]
            index = 0
            for i in accountNameFields:
                if "account name" in i["displayName"].lower():
                    accountNameOptionList = accountNameFields[index]["options"]
                    for j in accountNameOptionList:
                        if strAccountName.lower() == j["value"].lower():
                            strAccountCode = j["key"]
                            break
                    break
                index = index + 1
        logging.info("Account Code value :" + str(strAccountCode))
    except Exception as e:
        print(str(e))
        print(traceback.format_exc())
        strAccountCode = None
        logging.error("ERROR in GET Account code fn : " + str(e))
    return strAccountCode


# get salary Update fn call
def postSalaryUpdate(strurl, strToken):
    boolSalaryUpdateStatus = False
    try:
        payload = json.dumps({})
        headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer {}'.format(strToken)
        }
        response = requests.request("POST", strurl, headers=headers, data=payload)
        print(response.text)
        logging.info("Salary Update API Response code :" + str(int(response.status_code)))
        if response.status_code == 200:
            boolSalaryUpdateStatus = True
        else:
            boolSalaryUpdateStatus = False
        logging.info(" Salary update status :" + str(boolSalaryUpdateStatus))
    except Exception as e:
        print(str(e))
        print(traceback.format_exc())
        boolSalaryUpdateStatus = False
        logging.error("ERROR in salary update fn : " + str(e))
    return boolSalaryUpdateStatus

# post data updation fn call
def proposeOfferAPICall(proposeOfferUrl, bearer_token, dfs):
    boolAPICallResponse = False
    try:
        dfData = dfs
        # Data read from dataframe
        applicationId = str(dfData.loc[dfData.index[0], 'Application Id'])
        candidateID = str(dfData.loc[dfData.index[0], 'Candidate Id'])
        strBusinessKey = str(dfData.loc[dfData.index[0], 'business_key_'])
        BID = str(dfData.loc[dfData.index[0], 'BID'])
        strBID = BID.zfill(8)

        print(candidateID)
        print(applicationId)
        print(strBusinessKey)
        print(strBID)
        status=""
        if (len(candidateID) in canIdDigit and len(applicationId) in appIdDigit) and (candidateID != "" and applicationId != ""):
            print("Can id & App Id validation Pass")
            print(getCandidateDataAPI)

            # Application ID Check api call
            applicationIDCheckAPI = getCandidateDataAPI.format(candidateID)
            print(applicationIDCheckAPI)
            applicationIdLst = applicationIDCheckAPIfn(applicationIDCheckAPI, bearer_token)

            # Application Id Match check
            if int(applicationId) in applicationIdLst:
                print("Application ID Available")
                logging.info("Application ID Available in Intap :" + str(applicationId))
                print(str(dfData.loc[dfData.index[0], 'Indent']))
                indent=str(dfData.loc[dfData.index[0], 'Indent'])
                #strIndent = '' if (indent.lower() == "na") or indent.lower()=="none" else indent
                strRoleDesignation = str(dfData.loc[dfData.index[0], 'Role Designation'])
                strpuUnit = str(dfData.loc[dfData.index[0], 'Practice Unit'])
                strUnit = str(dfData.loc[dfData.index[0], 'Unit'])
                strCategory = str(dfData.loc[dfData.index[0], 'Category'])
                DOJ=str(dfData.loc[dfData.index[0], 'Date Of Joining'])
                strDOJ ="" if DOJ.lower()=="none" else DOJ
                strRelvExperience = str(dfData.loc[dfData.index[0], 'Experience'])
                strLOP = str(dfData.loc[dfData.index[0], 'Location Of Posting'])
                strCCMailId = str(dfData.loc[dfData.index[0], 'CC Employee Name'])
                GrossPerAnnums = str(dfData.loc[dfData.index[0], 'Salary Proposed'])
                strHiringTypes = str(dfData.loc[dfData.index[0], 'Hiring Type'])
                joiningBouns = str(dfData.loc[dfData.index[0], 'Joining Bonus'])
                a=0
                strJoiningBonus = a if (joiningBouns.lower() == "na") or joiningBouns.lower()=="none" else int(joiningBouns)
                strGrossPerAnnums = a if (GrossPerAnnums.lower() == "na") or GrossPerAnnums.lower() == "none" else int(
                    GrossPerAnnums)
                strStretch = str(dfData.loc[dfData.index[0], 'Stretch'])
                strT50 = str(dfData.loc[dfData.index[0], 'T50'])
                ae = str(dfData.loc[dfData.index[0], "AE"])
                strEstablishmentSubtypes = str(dfData.loc[dfData.index[0], 'Establishment Type'])
                strRWI = str(dfData.loc[dfData.index[0], "RWI"])
                strAccountName = str(dfData.loc[dfData.index[0], 'Account Name'])
                strDigitalTag = str(dfData.loc[dfData.index[0], 'Digital Tag'])
                strDigitalOffering = str(dfData.loc[dfData.index[0], 'Digital Offering'])

                if strEstablishmentSubtypes.lower() == 'sez1':
                    strEstablishmentSubtypes = "SEZ"
                else :
                    strEstablishmentSubtypes=strEstablishmentSubtypes.replace(" ","");

                strTechnology = str(dfData.loc[dfData.index[0], 'Technology'])
                strRecruiterName = str(dfData.loc[dfData.index[0], 'Recuriter Name Approval'])
                strRecruiterManagerName = str(dfData.loc[dfData.index[0], 'Recuriter Manager Name'])
                strHighestQualification = str(dfData.loc[dfData.index[0], 'Education Category'])
                strEducationQualification = str(dfData.loc[dfData.index[0], 'Education Qualification'])
                strRemoteWrkFrmHome = str(dfData.loc[dfData.index[0], 'Remote Work From Home'])
                strPermanentWrkFrmHome = str(dfData.loc[dfData.index[0], 'Permanent Work From Home'])
                strPartTimeEndDate = ""
                strRemoteWrkLocationName = str(dfData.loc[dfData.index[0], 'Remote Work Location'])
                strPartTimeEmployee = str(dfData.loc[dfData.index[0], 'Part Time Employee'])
                strRemoteWorkEndDate = str(dfData.loc[dfData.index[0], 'End Date'])
                no="No"
                ad =str(dfData.loc[dfData.index[0], 'ad'])
                strAD = no if (ad.lower() == "na") or ad.lower()=="none" else ad
                strAE = no if (ae.lower() == "na") or ae.lower()=="none" else ae
                #strComments = "Technology/Skill - " + strTechnology.title() + " : T50 - " + strT50.title() + " : 2% AD - " + strAD.title() + " : AE - " + strAE.title() + "% : Stretch - " + strStretch.title() + "% : RWI – " + strRWI.title() + " : Account Name – " + strAccountName.title() + ""
                #print(strComments)
                strRecruiterId = str(dfData.loc[dfData.index[0], 'Recuriter Id Approval'])
                strRecruiterManagerId = str(dfData.loc[dfData.index[0], 'Recruiter Manager Id'])
                leapStatusUpdateResult=""


                #Indent Validation for BEF and Sales Unit
                if strCategory.lower()=="bef" or strCategory.lower()=="sales":
                    digits = re.findall(r'\d{5}', indent)
                    strIndent= digits[0] if digits else ''
                    strComments = "Technology/Skill - " + strTechnology.title() + " : T50 - " + strT50.title() + " : 2% AD - " + strAD.title() + " : AE - " + strAE.title() + "% : Stretch - " + strStretch.title() + "% : RWI – " + strRWI.title() + " : Account Name – " + strAccountName.title() + " : Indent - " + indent.title()
                    print(strComments)
                else:
                    strIndent = '' if (indent.lower() == "na") or indent.lower() == "none" else indent
                    strComments = "Technology/Skill - " + strTechnology.title() + " : T50 - " + strT50.title() + " : 2% AD - " + strAD.title() + " : AE - " + strAE.title() + "% : Stretch - " + strStretch.title() + "% : RWI – " + strRWI.title() + " : Account Name – " + strAccountName.title() + ""
                    print(strComments)

                # Get Remote work location code fn call
                if strRemoteWrkLocationName != "" and strRemoteWrkLocationName != None :
                    strRemoteWrkLocationCode = getRemoteWorkLocationCode(strGetRemoteWorkLocationURL, bearer_token,strRemoteWrkLocationName)
                else:
                    strRemoteWrkLocationCode = None

                # Permanent work from home validation check
                if strRemoteWrkFrmHome.lower() == "yes" and strPermanentWrkFrmHome.lower() == "yes":
                    strRemoteWrkFrmHome = ""
                    strPermanentWrkFrmHome = ""
                    strRemoteWrkLocationName = ""
                elif strRemoteWrkFrmHome.lower() == "yes" or strPermanentWrkFrmHome.lower()=="yes":
                    strRemoteWrkFrmHome = "Y";
                    strPartTimeEmployeeCode = "N";
                    if strPermanentWrkFrmHome.lower()=="yes":
                        strRemoteWorkEndDate="2078-12-31"
                elif strPartTimeEmployee.lower() == "yes":
                    strPartTimeEmployeeCode = "Y";
                    strRemoteWrkFrmHome = "N";
                else:
                    strRemoteWrkFrmHome = "N";
                    strPartTimeEmployeeCode = "N";
                    strRemoteWorkEndDate=None
                    strRemoteWrkLocationCode=None

                # CategoryCode Check
                if strCategory != ""  :
                    strCategoryId = getCategoryCode(getCategoryURL, bearer_token, strCategory)
                else:
                    strCategoryId = None

                # Get Eduation Category Code
                if strHighestQualification != "":
                    url = strGetEduationCategoryURL.replace("{0}", applicationId)
                    strUrl = url.replace("{1}", candidateID)
                    print(strUrl)
                    strEducationCategoryCode = getEducationCategryCode(strUrl, bearer_token, strHighestQualification)
                else:
                    strEducationCategoryCode = None

                # Get Eduation Considered Code
                if strEducationQualification != "":
                    url = strGetEducationConsideredListAPI
                    strEducationConsideredCode = getEducationConsideredCode(url, bearer_token,strEducationQualification)
                else:
                    strEducationConsideredCode = None

                # Get PU Code Value
                if strUnit != "":
                    url = strGetPuCodeURL.replace("{0}", applicationId)
                    strUrl = url.replace("{1}", strUnit)
                    print(strUrl)
                    strPuCode = getPUCode(strUrl, bearer_token,strUnit)
                else:
                    strPuCode = None

                # LOP Id Find Call
                if strLOP != "":
                    url = strGetLOPCodeURL.replace("{0}", applicationId)
                    strUrl = url.replace("{1}", strLOP)
                    print(strUrl)
                    strLOPCode = getLOPCode(strUrl, bearer_token, strLOP)
                else:
                    strLOPCode = None

                # Get Role Designtion code fn call
                if strRoleDesignation != "":
                    strRoleCode = getRoleDesCode(strGetRoleCodeURL, bearer_token, strRoleDesignation)
                else:
                    strRoleCode = "2014ASRACO"

                # Joining Bouns and stretch value check
                if strJoiningBonus == "" or strJoiningBonus == 'None':
                    strJoiningBonus = None
                if strStretch =="" or strStretch == None or strStretch.lower()=='none':
                    strStretch = None
                else:
                    strStretch = float(strStretch)

                print(strRelvExperience)
                print(type(strRelvExperience))
                if strRelvExperience =="" or strRelvExperience == None or strRelvExperience=='None':
                    strRelvExperience =None
                else:
                    strRelvExperience =int(strRelvExperience)
                # Get Account Name code Fn Call
                accUrl = strGetCandidateDataAPI.format(applicationId)
                strAccountCode = getAccountCode(accUrl, bearer_token, applicationId, strAccountName)
                print("Account Code")
                print(strAccountCode)

                # Propose offer fun call
                payload = json.dumps({
                    "indents": strIndent,
                    "indentTypes": "",
                    "candidateIds": candidateID,
                    "activityIds": "600101",
                    "rolecapCodes": strRoleCode,
                    "puIds": strPuCode,
                    "categoryIds": strCategoryId,
                    "dojs": strDOJ,
                    "applicationIds": applicationId,
                    "relevantExperiences": strRelvExperience,
                    "postingLocations": strLOPCode,
                    "cc": strCCMailId,
                    "grossPerAnnums": strGrossPerAnnums,
                    "hiringTypes": strHiringTypes,
                    "digitalOfferinges": strDigitalOffering,
                    "digitalTags": strDigitalTag,
                    "joiningBonuss": strJoiningBonus,
                    "stretchs": strStretch,
                    "establishmentSubtypes": strEstablishmentSubtypes,
                    "flag": "Draft",
                    "comments": strComments,
                    "recruiter": strRecruiterId,
                    "recruiterManager": strRecruiterManagerId,
                    "educationCategorys": strEducationCategoryCode,
                    "degreeIds" : strEducationConsideredCode,
                    "careerLevel": None,
                    "programmingLevel": None,
                    "remoteWorkHome": strRemoteWrkFrmHome,
                    "remoteWorkLocation": strRemoteWrkLocationCode,
                    "remoteEndDate": strRemoteWorkEndDate,
                    "longTermBonus": None,
                    "partTimeEmp": strPartTimeEmployeeCode,
                    "workType": None,
                    "partTimeEnddate": strPartTimeEndDate,
                    "partTimeTerm": None,
                    "additionalAddendumId": strAccountCode
                })
                headers = {
                    'Content-Type': 'application/json',
                    'Authorization': 'Bearer {}'.format(bearer_token)
                }
                response = requests.request("POST", proposeOfferUrl, headers=headers, data=payload)
                print(response.text)
                logging.info("Propose offer API Json Input :" + payload)
                logging.info("Propose offer API Response status code :" + str(int(response.status_code)))
                if response.status_code == 200:
                    boolAPICallResponse = True
                    strResponseText = response.text
                    url = postSalaryUpdateAPI.replace("{0}", strResponseText).replace("{1}", str(strGrossPerAnnums))
                    boolSalaryUpdateResult = postSalaryUpdate(url, bearer_token)
                    print("Salary Update Status " + str(boolSalaryUpdateResult))
                    status = "preview"
                    reasonCode = "approved"
                    comments = "dataUpdated Successfully"
                    statusUpdateResult = statusUpdateInLeap(leapStatusUpdateUrl, strLeapToken, status, reasonCode,
                                                            comments,
                                                            strBID, strBusinessKey);
                    print(statusUpdateResult)
                else:
                    boolAPICallResponse = False
                    status = "exception"
                    reasonCode = "approved"
                    if response.status_code==500 and response.text.lower()=='improper data exception':
                        comments = "Input Json format is wrong"
                    else:
                        comments = "dataUpdated failed"
                    statusUpdateResult = statusUpdateInLeap(leapStatusUpdateUrl, strLeapToken, status, reasonCode,
                                                            comments,
                                                            strBID, strBusinessKey)
                    print(statusUpdateResult)

                if  statusUpdateResult==True:
                    leapStatusUpdateResult="Case ID Status Updated in leap"
                else:
                    leapStatusUpdateResult="Case Id Status updation Failed"

            else:
                print("Given Application ID not available in intap")
                boolAPICallResponse = False
                status = "exception"
                reasonCode = "exception"
                comments = "dataUpdated failed- Given Application ID is not available in intap"
                statusUpdateResult = statusUpdateInLeap(leapStatusUpdateUrl, strLeapToken, status, reasonCode, comments,
                                                        strBID, strBusinessKey)
                print(statusUpdateResult)
                if  statusUpdateResult==True:
                    leapStatusUpdateResult="Case ID Status Updated in leap"
                else:
                    leapStatusUpdateResult="Case Id Status updation Failed"

        else:
            print("Candidate ID or Application ID Validation Fail")
            boolAPICallResponse = False
            status = "exception"
            reasonCode = "exception"
            comments = "DataUpdation failed- Candidate ID or Application ID Validation Fail"
            statusUpdateResult = statusUpdateInLeap(leapStatusUpdateUrl, strLeapToken, status, reasonCode, comments,
                                                    strBID, strBusinessKey)
            print(statusUpdateResult)
            if statusUpdateResult == True:
                leapStatusUpdateResult = "Case ID Status Updated in leap"
            else:
                leapStatusUpdateResult = "Case Id Status updation Failed"

    except Exception as e:
        print(str(e))
        print(traceback.format_exc())
        logging.error("ERROR in Propose offer fn : " + str(e))
    return boolAPICallResponse,status,reasonCode,comments,leapStatusUpdateResult

# applies recursion to fetch count inside dictionaries and store into the count list
def myprint(globalInbox):
    try:
        for k, v in globalInbox.items():
            if isinstance(v, dict):
                myprint(v)
            else:
                # print(type(v))
                if isinstance(v, str):
                    try:
                        data = json.loads(v)
                        count[k] = data['count']
                        logging.info("count[k] : " + str(count[k]))
                        # print(k+" "+data['count'])
                    except:
                        logging.error("raise error")
                # print("{0} : {1}".format(k, v))
    except Exception as e:
        logging.error("ERROR in myprint method : " + str(e))


# Convert the Timedelta to seconds and format as hh:mm:ss
def format_timedelta(td):
  seconds = int(td.total_seconds())
  hours, remainder = divmod(seconds, 3600)
  minutes, seconds = divmod(remainder, 60)
  return f"{hours:02d}:{minutes:02d}:{seconds:02d}"



# Leap Authorization
strLeapToken = leapAuthorization()

# Intap authorization
bearer_token = intapAuthorization(intapAuthorizationUrl)

if  strLeapToken !=None and bearer_token !=None:
    # global case count api
    try:
        headers = {'Authorization': 'Bearer ' + strLeapToken, 'Content-type': 'application/json',
                   'Accept': 'application/json, text/plain, */*'}
        globalInbox = requests.request("GET", GlobalCaseUrl, headers=headers, data=payload)
        globalInbox_Status = globalInbox.status_code
        logging.info("Get global cases Inbox status code: " + str(int(globalInbox_Status)))
        if globalInbox_Status == 200:
            globalInbox = globalInbox.json()
            logging.info("Global Inbox: " + str(globalInbox))
        else:
            logging.error("Global inbox status is not 200")
    except Exception as e:
        logging.error("ERROR in global inbox count api : " + str(e))

    myprint(globalInbox)  # calling function

    # print(len(count))
    # print(count["Inbox"]["count"])

    # Inbox case id take method
    try:
        InboxContent = requests.request("GET", InboxCountUrl, headers=headers, data=payload)
        InboxContent_Status = InboxContent.status_code
        logging.info("Inbox Content status : " + str(InboxContent_Status))
        if InboxContent_Status == 200:
            InboxContent = InboxContent.json()
            logging.info("Type od InboxContent : " + str(type(InboxContent)))
            lstData = InboxContent['data']
            # print(lstData)
            logging.info("Type of InboxContent : " + str(type(lstData)))
            for i in range(len(lstData)):
                logging.info("ID" + str(lstData[i]['ID']))
                caseid = re.findall("[0-9]{8}", lstData[i]['ID'])
                caseid = ''.join(caseid)
                caseid = int(caseid.replace("^0x", ""))
                caseId.append(caseid)
            logging.info("Length of caseid: " + str(len(caseId)))
            logging.info("Case ID: " + str(caseId))
    except Exception as e:
        logging.error("ERROR in Inbox case id take method : " + str(e))

    # Datacount API Fn call
    try:
        DataSetCount = requests.request("GET", DataCountAPI, headers=headers, data=payload)
        DataSetCount_Status = DataSetCount.status_code
        logging.info("DatasetCount API Status Code: " + str(DataSetCount_Status))
        if DataSetCount_Status == 200:
            DataSetCount = DataSetCount.json()
            logging.info("Dataset Count : " + str(DataSetCount))
            totalIteration = math.ceil(DataSetCount / 10)
        else:
            logging.error("Dataset count status is not 200")
    except Exception as e:
        logging.error("ERROR in Datacount API method : " + str(e))

    try:
        caseIdMatchCount = 0
        # while len(caseId)!=0:
        if len(caseId) > 0:
            for i in range(totalIteration):
                DataSetAPI1 = DataSetAPI.format(i)
                logging.info("Data set API : " + DataSetAPI1)
                DataSet = requests.request("GET", DataSetAPI1, headers=headers, data=payload)
                DataSet_Status = DataSet.status_code
                logging.info("GET Dataset API Status code: " + str(DataSet_Status))
                if DataSet_Status == 200:
                    DataSet = DataSet.json()
                    # print("Dataset",DataSet)

                    for i in range(len(DataSet)):
                        logging.info("ID : " + str(DataSet[i]['BID']))
                        # print(DataSet[i]['BID'],caseId)
                        bid = int(DataSet[i]['BID'])
                        # print(type(bid))

                        if bid in caseId:
                            logging.info("ID in case ID : " + str(DataSet[i]['BID']))
                            # df=pd.json_normalize(DataSet[i])
                            lstDataSetData.append(DataSet[i])
                            caseIdMatchCount = caseIdMatchCount + 1
                            if caseIdMatchCount == len(caseId):
                                countMatchFlag = True;

                    if countMatchFlag == True:
                        logging.error("Breaking total iteration loop")
                        break
                else:
                    logging.info("Dataset API response in not 200")
        else:
            logging.info("Bot Queue inbox count is zero")
        logging.info("case id match count" + str(caseIdMatchCount))
        #logging.info("list of data set data : " + str(lstDataSetData))
    except Exception as e:
        logging.error("ERROR in dataset API Process: " + str(e))
    if caseIdMatchCount > 0:
        try:
            '''
            with open('data.json', 'w', encoding='utf-8') as f:
                json.dump(lstDataSetData, f, ensure_ascii=False, indent=4)
    
    
            pd.read_json("data.json").to_excel(xlsx_url)
    
            df = pd.read_excel(xlsx_url, index_col=0)
            
            '''

            df = pd.DataFrame(lstDataSetData)

            df.drop(["last_updated_date_", "tenant_id_", "proc_def_key_", "proc_inst_id_", "duration_", "task_name_",
                     "_extras", "task_def_key_", "due_date_", "state_", "created_date_", "requested_by", "status_", "_salt",
                     "source_",
                     "current_assignee_", "last_assignee_", "end_date_", "src", "employeeNameOne", "employeeIdOne",
                     "employeeRoleOne", "employeeNameTwo",
                     "employeeIdTwo", "employeeRoleTwo", "employeeNameThree", "employeeIdThree", "employeeRoleThree",
                     "issueCategory",
                     "offerReceivedDate", "processedOnDate", "offersToCB"], axis=1, inplace=True)

            df.rename(columns={"category": "Category", "dateOfJoining": "Date Of Joining", "indent": "Indent",
                               "locationOfPosting": "Location Of Posting",
                               "establishmentType": "Establishment Type", "hiringType": "Hiring Type",
                               "practiceUnit": "Practice Unit", "roleDesignation": "Role Designation",
                               "role": "Role", "jobLevel": "Job Level", "jobSubLevel": "Job Sub Level",
                               "personalSubLevel": "Personal Sub Level",
                               "personalLevel": "Personal Level", "remoteWorkFromHome": "Remote Work From Home",
                               "permanentWorkFromHome": "Permanent Work From Home",
                               "educationQualification": "Education Qualification",
                               "educationCategory": "Education Category",
                               "experience": "Experience",
                               "salaryProposed": "Salary Proposed", "joiningBonus": "Joining Bonus", "t50": "T50",
                               "ae": "AE",
                               "stretch": "Stretch",
                               "recuriterIdContactDetails": "Recuriter Id Contact Details",
                               "recruiterManagerId": "Recruiter Manager Id", "ccEmployeeId": "CC Employee Id",
                               "recuriterNameApproval": "Recuriter Name Approval", "endDate": "End Date",
                               "recuriterManagerName": "Recuriter Manager Name",
                               "unit": "Unit", "digitalOffering": "Digital Offering",
                               "remoteWorkLocation": "Remote Work Location",
                               "digitalTag": "Digital Tag", "candidateId": "Candidate Id",
                               "applicationId": "Application Id",
                               "roleAverage": "Role Average",
                               "short_description": "Subject Line", "recuriterIdApproval": "Recuriter Id Approval",
                               "offerType": "Offer Type", "partTimeEmployee": "Part Time Employee",
                               "ccEmployeeName": "CC Employee Name",
                               "reasonCode": "Reason Code",
                               "recuriterNameContactDetails": "Recuriter Name Contact Details",
                               "accountName": "Account Name", "rwi": "RWI", "tech": "Technology",
                               "candidateName": "Candidate Name"}, inplace=True)
            data = df

            '''
            excel = data.to_excel(str(path) + '\\data.xlsx')
            logging.info("Type of the excel : " + str(type(excel)))
    
            today = date.today()
            today = today.strftime("%d-%m-%Y")
            dateFolderPath = os.path.join(path, today)
            if os.path.exists(dateFolderPath):
                logging.info("Date folder path already exists")
            else:
                os.mkdir(dateFolderPath)
            attachments = []
            '''
            for row in range(len(df)):
                starttime = pd.to_datetime(datetime.now().strftime("%H:%M:%S"))
                print(row)
                print(len(df))
                # print(df.iat[row,46])
                # print(df.iat[row,'business_key_'])
                print(df.loc[df.index[row], 'business_key_'])
                # subFolderName = df.loc[row]['business_key_']

                subFolderName = df.loc[df.index[row], 'business_key_']
                subFolderName = subFolderName.replace(":", "-")
                '''
                subFolderPath = os.path.join(dateFolderPath, subFolderName)
                logging.info("Sub folder path : " + str(subFolderPath))
                if os.path.exists(subFolderPath):
                    logging.info("folder exists")
                else:
                    os.mkdir(subFolderPath)
                #print(data["business_key_"] == subFolderName.replace("-", ":"))
                '''
                sub_df = data[data["business_key_"] == subFolderName.replace("-", ":")]
                # print(sub_df)
                boolProposeOfferAPICallResponse,status,reasonCode,comments,leapStatusUpdateResult=proposeOfferAPICall(proposeOfferUrl, bearer_token, sub_df)

                '''
                if boolProposeOfferAPICallResponse== True:
                    Status = 'Pass'
                    Comments = ''
                else:
                    Status = 'Exception'
                    Comments = "Couldn't able to generate excel for this case id"
                    
                '''

                # creating output file
                logging.info("Output file generation - Started")
                current_date = datetime.now().strftime("%Y.%m.%d")
                #filename_op = f"output_{str(current_date)}_DatasetExtraction.xlsx"
                filename_op = f"output_{str(current_date)}_DatasetExtraction.csv"
                endTime=pd.to_datetime(datetime.now().strftime("%H:%M:%S"))

                time_diff=endTime-starttime
                formatted_time = format_timedelta(time_diff)

                # Sample data for demonstration
                data_1 = {
                    'Case ID': [subFolderName],
                    'Start Time': [starttime],
                    'End Time': [endTime],
                    'Run Time': [formatted_time],
                    'Status': [status],
                    'Reason Code':[reasonCode],
                    'Comments': [comments],
                    'LEAP Status Update Result':[leapStatusUpdateResult],
                    'User Name': [os.getlogin()]
                }

                # Create a DataFrame with the data
                df_1 = pd.DataFrame(data_1)

                # Append the data to the existing file or create a new file
                try:
                    # Try to read the existing file
                    #existing_data = pd.read_excel(filename_op)
                    existing_data = pd.read_csv(filename_op)

                    # Append the new data to the existing file
                    updated_data = pd.concat([existing_data, df_1], ignore_index=True)

                    # Write the updated data to the file
                    #updated_data.to_excel(filename_op, index=False)
                    updated_data.to_csv(filename_op, index=False)

                except FileNotFoundError:
                    # If the file doesn't exist, create a new file and write the data
                    df_1.to_csv(filename_op, index=False)
                logging.info("Output file generation Completed")
                body = """
                Hi,
    
                    Case ID details excel file's are generated in the processing folder. Please find the attached output file.
    
                Thanks,
                RPA Script"""

                attachments = [os.path.join(os.getcwd(), str(filename_op))]
                # logging.info("Attachment list length : "+str(len(attachments)))

                send_email(subject, body, recipients, attachments)

                '''
                filePath = subFolderPath + "\\CaseDetails-{0}.xlsx".format(subFolderName)
                logging.info("File Path : " + str(filePath))
                sub_df.to_excel(filePath)
                if os.path.exists(filePath):
                    Status='Pass'
                    Comments=''
                else:
                    Status='Exception'
                    Comments="Couldn't able to generate excel for this case id"
                
                #creating output file
                logging.info("Output file generation - Started")
                current_date = datetime.now().strftime("%Y.%m.%d")
                filename_op = f"output_{str(current_date)}_DatasetExtraction.xlsx"
    
                # Sample data for demonstration
                data_1 = {
                    'Case ID': [subFolderName],
                    'Start Time': [starttime],
                    'End Time': [datetime.now().strftime("%H:%M:%S")],
                    #'Run Time': [datetime.now().strftime("%H:%M:%S") - starttime],
                    'Status': [Status],
                    'Comments': [Comments],
                    'Machine Name':[socket.gethostname()]
                }
    
                # Create a DataFrame with the data
                df_1 = pd.DataFrame(data_1)
    
                # Append the data to the existing file or create a new file
                try:
                    # Try to read the existing file
                    existing_data = pd.read_excel(filename_op)
    
                    # Append the new data to the existing file
                    updated_data = pd.concat([existing_data, df_1], ignore_index=True)
    
                    # Write the updated data to the file
                    updated_data.to_excel(filename_op, index=False)
    
                except FileNotFoundError:
                    # If the file doesn't exist, create a new file and write the data
                    df_1.to_excel(filename_op, index=False)
                logging.info("Output file generation Completed")
                body = """
    Hi,
    
        Case ID details excel file's are generated in the processing folder. Please find the attached output file.
    
    Thanks,
    RPA Script"""
    
                attachments=[os.path.join(os.getcwd(), str(filename_op))]
                #logging.info("Attachment list length : "+str(len(attachments)))
    
            #send_email(subject, body, recipients, attachments)
    
        except Exception as e:
            logging.error("ERROR : " + str(e))
            print(str(e))
            '''
        except Exception as e:
            print(str(e))
            print(traceback.format_exc())
    elif caseIdMatchCount <= 0:
        print("Case ID Match count is Zero")
        logging.info("Case ID match count : " + str(caseIdMatchCount) + " so, we cannot generate a excel file")
        body = """
    Hi,
                
        Case ID match count is zero. so, we cannot generate a excel file in the processing folder.
    
    Thanks,
    RPA Script"""
        send_email(subject, body, recipients)
else:
    print("Intap or Leap Authorization Failed")
    logging.error("Intap or Leap Authorization Failed")
    body="Leap or Intap authorization issue, please check on VPN Connectivity in your system"
    send_email(subject, body, recipients)