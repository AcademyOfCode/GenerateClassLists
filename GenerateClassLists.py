import os
import csv
import datetime
import dateutil.parser
import requests
import time
import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import io
import sys
import openpyxl
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('apiKey')
parser.add_argument('testMode')
parser.add_argument('emailPassword')
args = parser.parse_args()

emailPassword = args.emailPassword
apiKey = args.apiKey
apiURL = 'https://api.squarespace.com/1.0/commerce/orders'
parentDirectory = 'C:/Users/Robert'
documentsDirectory = parentDirectory + '/Documents'
ordersDirectory = documentsDirectory + '/Class Orders/'
classListDirectory = documentsDirectory + '/Class Lists/'

startDate = '2019-01-07T11:20:00.000000Z'
endDate = datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%S.%fZ')

testMode = bool(args.testMode)

def GetDateTimeFromISO8601String(s):
    return dateutil.parser.parse(s)

def WriteLastGenerationDate():
    with open('LastClassListGenerationDate.txt', 'w') as file:
        file.write(endDate)

def ReadLastGenerationDate():
    with open('LastClassListGenerationDate.txt') as file:
        lastEndDate = GetDateTimeFromISO8601String(file.readlines()[0])
    startDate = lastEndDate + datetime.timedelta(microseconds=1)

def ExportAllOrders(startDate, endDate):
    pageNumber = 1
    nextPageCursor = ""
    responsePageList = []

    headers = {
        'Authorization': 'Bearer ' + apiKey,
    }

    params = (
        ('modifiedAfter', startDate),
        ('modifiedBefore', endDate),
    )

    response = requests.get(apiURL, headers=headers, params=params)
    print('Page ' + str(pageNumber) + ' Request Completed.')
    responsePageList.append(response)
    responseTextLines = response.text.split('\n')

    while 'true' in responseTextLines[len(responseTextLines)-3]:
        pageNumber += 1

        for line in responseTextLines:
            if 'nextPageCursor' in line:
                nextPageCursor = line.split(':')[1].replace('"', '').replace(',', '').replace(' ', '')

        params = (
            ('cursor', nextPageCursor),
        )

        response = requests.get(apiURL, headers=headers, params=params)
        responseTextLines = response.text.split('\n')
        print('Page ' + str(pageNumber) + ' Request Completed.')
        responsePageList.append(response)

    return responsePageList

def ExportIndividualOrders(allOrdersList):
    orderCount = 1
    responseList = []
    classTypeList = []
    fullYearList = []
    abortClassListGeneration = False

    headers = {
        'Authorization': 'Bearer ' + apiKey,
        'User-Agent' : 	'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36'
    }

    print('Exporting Individual Order...')
    for orderList in allOrdersList:

        lineNumber = 0
        lineNumberList = []

        for line in orderList.text.split('\n'):
            if 'orderNumber' in line:
                lineNumberList.append(lineNumber)
            lineNumber += 1

        for lineNum in lineNumberList:
            productFormError = False
            writeToProductFormErrorsFile = False

            orderID = orderList.text.split('\n')[lineNum-1].split(':')[1].replace(' ','').replace('"','').replace(',','')
            response = requests.get(apiURL + '/' + orderID, headers=headers)

            if 'Custom payment amount' in response.json()['lineItems'][0]['productName']:
                continue

            if 'Test' in response.json()['lineItems'][0]['customizations'][0]['value']:
                if testMode:
                    if 'Tech Club' in response.json()['lineItems'][0]['productName'].split('- ')[1]:
                        classType = (response.json()['lineItems'][0]['productName'].split()[0],response.json()['lineItems'][0]['productName'].split()[1], 'Tech Club')
                    else:
                        classType = (response.json()['lineItems'][0]['productName'].split()[0],response.json()['lineItems'][0]['productName'].split()[1])

                    if not classType[1].isdigit():
                        productFormError = True
                        abortClassListGeneration = True

                    if productFormError is False:
                        if classType not in classTypeList:
                            classTypeList.append(classType)
                        responseList.append(response.json())

                        if '30 Weeks' in response.json()['lineItems'][0]['variantOptions'][1]['value']:
                            classType = ('Spring', str(int(response.json()['lineItems'][0]['productName'].split()[1]) + 1), 'Next Term')

                            if classType not in classTypeList:
                                classTypeList.append(classType)

                            fullYearList.append(response.json())

                        print('Order ' + str(orderCount) + ' Completed.')
                        orderCount += 1
                    else:
                        if os.path.isfile('ProductFormErrors.txt'):
                            with open('ProductFormErrors.txt', 'r') as file:
                                if response.json()['lineItems'][0]['productName'] + '\n' not in file.readlines():
                                    writeToProductFormErrorsFile = True
                        else:
                            with open('ProductFormErrors.txt', 'w') as file:
                                file.write(response.json()['lineItems'][0]['productName'] + '\n')
            else:
                if 'Tech Club' in response.json()['lineItems'][0]['productName'].split('- ')[1]:
                    classType = (response.json()['lineItems'][0]['productName'].split()[0],
                                 response.json()['lineItems'][0]['productName'].split()[1], 'Tech Club')
                else:
                    classType = (response.json()['lineItems'][0]['productName'].split()[0],
                                 response.json()['lineItems'][0]['productName'].split()[1])

                if not classType[1].isdigit():
                    productFormError = True
                    abortClassListGeneration = True

                if productFormError is False:
                    if classType not in classTypeList:
                        classTypeList.append(classType)
                    responseList.append(response.json())

                    if '30 Weeks' in response.json()['lineItems'][0]['variantOptions'][1]['value']:
                        classType = ('Spring', str(int(response.json()['lineItems'][0]['productName'].split()[1]) + 1), 'Next Term')

                        if classType not in classTypeList:
                            classTypeList.append(classType)

                        fullYearList.append(response.json())

                    print('Order ' + str(orderCount) + ' Completed.')
                    orderCount += 1
                else:
                    if os.path.isfile('ProductFormErrors.txt'):
                        with open('ProductFormErrors.txt', 'r') as file:
                            if response.json()['lineItems'][0]['productName'] + '\n' not in file.readlines():
                                writeToProductFormErrorsFile = True
                    else:
                        with open('ProductFormErrors.txt', 'w') as file:
                            file.write(response.json()['lineItems'][0]['productName'] + '\n')

            if writeToProductFormErrorsFile:
                with open('ProductFormErrors.txt', 'a') as file:
                    file.write(response.json()['lineItems'][0]['productName'] + '\n')

    return responseList, classTypeList, fullYearList, abortClassListGeneration

# def ListCSVFilesInDirectory(directory):
#     return [f for f in os.listdir(directory) if os.path.isfile(directory + f) and f.__contains__('orders') and f.endswith('.csv')]
#
# def ReadCSVFile(filePath):
#     dataList = []
#
#     # with open('C:/Users/Robert/Documents/orders.json', encoding='utf-8') as f:
#     #     data = json.load(f)
#
#     with open(filePath, newline='',encoding='utf-8') as csvFile:
#         reader = csv.DictReader(csvFile)
#
#         for row in reader:
#             dataList.append(row)
#
#     return dataList
#
# def WriteCSVFile(directory, data):
#     variants = []
#     numClasses = 0
#     dataGapSize = 27
#
#     if not os.path.exists(directory):
#         os.makedirs(directory)
#
#     for d in data:
#         if '/' in d['Lineitem variant']:
#             d['Lineitem variant'] = d['Lineitem variant'].split('/')[0]
#         if d['Lineitem variant'].split(',')[0:2] not in variants:
#             variants.append(d['Lineitem variant'].split(',')[0:2])
#
#     seperateClassDataList = [[] for j in range(len(variants))]
#
#     for j in range(len(seperateClassDataList)):
#         for k in range(len(data)):
#             if data[k]['Lineitem variant'].split(',')[0:2] == variants[j]:
#                 seperateClassDataList[j].append(data[k])
#
#     for seperateClassData in seperateClassDataList:
#
#         dataLines = []
#         seperatedNameClassData = []
#
#         for l in range(len(seperateClassData)):
#             if len(seperateClassData[l]['Product Form: Student Name'].split()) > 1 and len(
#                     seperateClassData[l]['Product Form: Student 2 Name'].split()) > 1 and len(
#                 seperateClassData[l]['Product Form: Student 3 Name'].split()) == 0:
#                 seperatedNameClassData.append(seperateClassData[l])
#             elif len(seperateClassData[l]['Product Form: Student Name'].split()) > 1 and len(
#                     seperateClassData[l]['Product Form: Student 2 Name'].split()) == 0 and len(
#                 seperateClassData[l]['Product Form: Student 3 Name'].split()) > 1:
#                 seperatedNameClassData.append(seperateClassData[l])
#             elif len(seperateClassData[l]['Product Form: Student Name'].split()) > 1 and len(
#                     seperateClassData[l]['Product Form: Student 2 Name'].split()) > 1 and len(
#                 seperateClassData[l]['Product Form: Student 3 Name'].split()) > 1:
#                 seperatedNameClassData.append(seperateClassData[l])
#                 seperatedNameClassData.append(seperateClassData[l])
#             seperatedNameClassData.append(seperateClassData[l])
#
#         prevOrderID = 0
#         thirdOrder = False
#
#         studentNames = ['' for i in range(len(seperatedNameClassData))]
#         orderIDs = ['' for i in range(len(seperatedNameClassData))]
#         emails = ['' for i in range(len(seperatedNameClassData))]
#         billingNames = ['' for i in range(len(seperatedNameClassData))]
#         phoneNumbers = ['' for i in range(len(seperatedNameClassData))]
#         studentDOBs = ['' for i in range(len(seperatedNameClassData))]
#         studentSchoolsClasses = ['' for i in range(len(seperatedNameClassData))]
#         returningStudent = ['' for i in range(len(seperatedNameClassData))]
#         lastTerm = ['' for i in range(len(seperatedNameClassData))]
#         additionalSupport = ['' for i in range(len(seperatedNameClassData))]
#         photographyConsent = ['' for i in range(len(seperatedNameClassData))]
#         otherDetails = ['' for i in range(len(seperatedNameClassData))]
#         studentData = ['' for i in range(len(seperatedNameClassData))]
#
#         for m in range(len(seperatedNameClassData)):
#             if seperatedNameClassData[m]['Order ID'] == prevOrderID:
#                 if thirdOrder:
#                     studentNames[m] = seperatedNameClassData[m]['Product Form: Student 3 Name']
#                     studentDOBs[m] = seperatedNameClassData[m]['Product Form: Student 3 Date of Birth']
#                     thirdOrder = False
#                 else:
#                     studentNames[m] = seperatedNameClassData[m]['Product Form: Student 2 Name']
#                     studentDOBs[m] = seperatedNameClassData[m]['Product Form: Student 2 Date of Birth']
#                     thirdOrder = True
#             else:
#                 studentNames[m] = seperatedNameClassData[m]['Product Form: Student Name']
#                 studentDOBs[m] = seperatedNameClassData[m]['Product Form: Student Date of Birth']
#                 thirdOrder = False
#
#             prevOrderID = seperatedNameClassData[m]['Order ID']
#
#             if len(seperatedNameClassData[m]['Product Form: Student Name'].split()) == 1 and len(seperatedNameClassData[m]['Product Form: Student 2 Name'].split()) == 1:
#                 if len(seperatedNameClassData[m]['Product Form: Student 3 Name'].split()) == 1:
#                     studentNames[m] = seperatedNameClassData[m]['Product Form: Student Name'] + ' ' + seperatedNameClassData[m]['Product Form: Student 2 Name'] + ' ' + seperatedNameClassData[m]['Product Form: Student 3 Name']
#                 else:
#                     studentNames[m] = seperatedNameClassData[m]['Product Form: Student Name'] + ' ' + seperatedNameClassData[m]['Product Form: Student 2 Name']
#             elif len(seperatedNameClassData[m]['Product Form: Student Name'].split()) == 1 and len(seperatedNameClassData[m]['Product Form: Student 3 Name'].split()) == 1:
#                 if len(seperatedNameClassData[m]['Product Form: Student 2 Name'].split()) == 1:
#                     studentNames[m] = seperatedNameClassData[m]['Product Form: Student Name'] + ' ' + seperatedNameClassData[m]['Product Form: Student 2 Name'] + ' ' + seperatedNameClassData[m]['Product Form: Student 3 Name']
#                 else:
#                     studentNames[m] = seperatedNameClassData[m]['Product Form: Student Name'] + ' ' + seperatedNameClassData[m]['Product Form: Student 3 Name']
#
#             orderIDs[m] = seperatedNameClassData[m]['Order ID']
#             emails[m] = seperatedNameClassData[m]['Email']
#             billingNames[m] = seperatedNameClassData[m]['Billing Name']
#             phoneNumbers[m] = seperatedNameClassData[m]['Billing Phone']
#             studentSchoolsClasses[m] = seperatedNameClassData[m]['Product Form: Student\'s School and Class']
#             returningStudent[m] = seperatedNameClassData[m]['Product Form: Are you a returning Academy of Code student?']
#             lastTerm[m] = seperatedNameClassData[m]['Product Form: If you are a returning student, when was your last term with AoC?']
#             if 'Product Form: Additional support for your child' in seperatedNameClassData[m]:
#                 additionalSupport[m] = seperatedNameClassData[m]['Product Form: Additional support for your child']
#             photographyConsent[m] = seperatedNameClassData[m]['Product Form: Photography Consent']
#             if 'Product Form: Other details' in seperatedNameClassData[m]:
#                 otherDetails[m] = seperatedNameClassData[m]['Product Form: Other details']
#
#         term = data[0]['Lineitem name'].split(' ')[0]
#
#         venue = data[0]['Lineitem name'].split(' - ')[1].split(',')[0].replace(' ', '')
#
#         if venue not in openVenues:
#             numClasses = 12
#         else:
#             if term == 'Autumn':
#                 numClasses = 12
#             elif term == 'Spring':
#                 numClasses = 18
#
#         for n in range(len(seperatedNameClassData)):
#             studentData[n] = [studentNames[n],
#                           orderIDs[n],
#                           emails[n],
#                           billingNames[n],
#                           phoneNumbers[n],
#                           studentNames[n],
#                           studentDOBs[n],
#                           studentSchoolsClasses[n],
#                           returningStudent[n],
#                           lastTerm[n],
#                           additionalSupport[n],
#                           photographyConsent[n],
#                           otherDetails[n]
#                           ]
#             for num in range(numClasses+dataGapSize):
#                 studentData[n].insert(1, ' ')
#
#             dataLines.append(studentData[n])
#
#         day = seperatedNameClassData[0]['Lineitem variant'].split(',')[0][0:3]
#
#         if 'Spring 2019 - Muslim National School, Clonskeagh' in seperatedNameClassData[0]['Lineitem name']:
#             time = '14:50'
#             day = 'Tuesday'
#         elif 'Spring 2018 - Kildare Town Educate Together National School' in seperatedNameClassData[0]['Lineitem name']:
#             time = '14:10'
#         elif 'Spring 2019 - The Teresian School' in seperatedNameClassData[0]['Lineitem name']:
#             time = seperatedNameClassData[0]['Lineitem variant'].split('(')[1][0:5].replace(":", "")
#         elif 'Thursday 16:00 - 17:30' in seperatedNameClassData[0]['Lineitem variant']:
#             time = seperatedNameClassData[0]['Lineitem variant'].split(' ')[1].replace(":", "")
#         else:
#             time = seperatedNameClassData[0]['Lineitem variant'].split(',')[1][1:6].replace(":", "")
#
#         #shortVenueIndex = re.search(r'^([^A-Z]*[A-Z]){2}', venue).span()[1]
#         shortVenue = venue[0:10]
#
#         print(directory + day + shortVenue + time + '.csv')
#
#         classListTemplateLines = [['Date'],
#                          ['Student Name'],
#                          ['Class ' + str(classNum + 1) for classNum in range(numClasses)],
#                          [' ' for gap in range(dataGapSize-1)],
#                          ['Gender', 'Order ID', 'Email', 'Billing Name', 'Phone', 'Student Name(s)', 'Student Date(s) of Birth',
#                           'Student\'s School and Class', 'Are you a returning Academy of Code student?',
#                           'If you are a returning student, when was your last term with AoC?',
#                           'Additional support for your child', 'Photography Consent', 'Other details',
#                           'Other Teacher Notes']
#                          ]
#         classListTemplateLines[1] = classListTemplateLines[1] + classListTemplateLines[2] + classListTemplateLines[3] + classListTemplateLines[4]
#
#         for i in range(3):
#             del classListTemplateLines[2]
#
#         with open(directory + day + shortVenue + time + '.csv', 'w') as writeFile:
#             writer = csv.writer(writeFile,lineterminator = '\n')
#
#             writer.writerows(classListTemplateLines)
#             writer.writerows(dataLines)

def GoogleDriveAccess():
    print('Getting Access to Google Drive.')

    gauth = GoogleAuth()
    gauth.LoadCredentialsFile("credentials.txt")

    if gauth.credentials is None:
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        gauth.Refresh()
    else:
        gauth.Authorize()
    gauth.SaveCredentialsFile("credentials.txt")

    return GoogleDrive(gauth)

def DownloadClassListsFromGoogleDrive(drive, classListName):
    downloadMimeType = None

    mimeTypes = {
        'application/vnd.google-apps.document': 'application/pdf',
        'application/vnd.google-apps.spreadsheet': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }

    if classListName + '.xlsx' in os.listdir("."):
        os.remove(classListName + '.xlsx')

    fileList = drive.ListFile({'q': "'root' in parents and trashed=false"}).GetList()

    for file in fileList:
        if 'Class lists' in file['title']:
            folderContentList = drive.ListFile({'q': "'{}' in parents and trashed=false".format(file['id'])}).GetList()

    for file in folderContentList:
        if classListName in file['title']:
            if file['mimeType'] in mimeTypes:
                print('Downloading Class Lists from Google Drive.')
                downloadMimeType = mimeTypes[file['mimeType']]
                file.GetContentFile(file['title'], mimetype=downloadMimeType)
            break

    if classListName not in os.listdir("."):
        wb = openpyxl.Workbook()
        wb.save(classListName + '.xlsx')

def AppendDfToExcel(fileName, df, sheetName, book, startRow=None, truncateSheet=False, **toExcelKwargs):
    if 'engine' in toExcelKwargs:
        toExcelKwargs.pop('engine')

    writer = pd.ExcelWriter(fileName, engine='openpyxl')

    try:
        writer.book = book

        if startRow is None and sheetName in writer.book.sheetnames:
            startRow = writer.book[sheetName].max_row

        if truncateSheet and sheetName in writer.book.sheetnames:
            idx = writer.book.sheetnames.index(sheetName)
            writer.book.remove(writer.book.worksheets[idx])
            writer.book.create_sheet(sheetName, idx)

        writer.sheets = {ws.title:ws for ws in writer.book.worksheets}
    except FileNotFoundError:
        pass

    if startRow is None:
        startRow = 0

    df.to_excel(writer, sheetName, header=False, index=False, startrow=startRow, **toExcelKwargs)

    writer.save()

def CreateAndAppendClassLists(orderList, classListName):
    dataGapSize = 27

    if 'Autumn' in classListName:
        numClasses = 12
    elif 'Spring' in classListName:
        numClasses = 18
    else:
        numClasses = 5

    summaryTemplateLines = [['Venue','No. of Students']
                            ]

    classListTemplateLines = [['Date'],
                             ['Student Name'],
                             ['Class ' + str(classNum + 1) for classNum in range(numClasses)],
                             [' ' for gap in range(dataGapSize - 1)],
                             ['Gender', 'Order ID', 'Email', 'Billing Name', 'Phone', 'Student Name(s)',
                              'Student Date(s) of Birth',
                              'Student\'s School and Class', 'Are you a returning Academy of Code student?',
                              'If you are a returning student, when was your last term with AoC?',
                              'Additional support for your child', 'Photography Consent', 'Other Details',
                              'Other Teacher Notes']
                            ]

    classListTemplateLines[1] = classListTemplateLines[1] + classListTemplateLines[2] + classListTemplateLines[3] + classListTemplateLines[4]

    for i in range(3):
        del classListTemplateLines[2]

    print('Sorting Class Lists')

    for fileName in os.listdir("."):
        if classListName in fileName:
            if '.xlsx' not in fileName:
                os.rename(fileName, fileName + '.xlsx')
                fileName = fileName + '.xlsx'

            seperatedNameOrderList = []

            for order in orderList:
                if len(order['lineItems'][0]['customizations'][0]['value'].split()) > 1 and len(
                        order['lineItems'][0]['customizations'][1]['value'].split()) > 1 and len(
                    order['lineItems'][0]['customizations'][2]['value'].split()) == 0:
                    seperatedNameOrderList.append(order)
                elif len(order['lineItems'][0]['customizations'][0]['value'].split()) > 1 and len(
                        order['lineItems'][0]['customizations'][1]['value'].split()) == 0 and len(
                    order['lineItems'][0]['customizations'][2]['value'].split()) > 1:
                    seperatedNameOrderList.append(order)
                elif len(order['lineItems'][0]['customizations'][0]['value'].split()) > 1 and len(
                        order['lineItems'][0]['customizations'][1]['value'].split()) > 1 and len(
                    order['lineItems'][0]['customizations'][2]['value'].split()) > 1:
                    seperatedNameOrderList.append(order)
                    seperatedNameOrderList.append(order)
                seperatedNameOrderList.append(order)

            prevOrderID = 0
            thirdOrder = False

            book = load_workbook(fileName)
            writer = pd.ExcelWriter(fileName, engine='openpyxl')
            writer.book = book

            if 'Summary' not in book.sheetnames:
                summaryTemplateDataFrame = pd.DataFrame(data=summaryTemplateLines)
                summaryTemplateDataFrame.to_excel(writer, sheet_name='Summary', header=False, index=False)

                sheet = book['Summary']
                cell = sheet['A1']
                cell.font = Font(bold=True)
                cell = sheet['B1']
                cell.font = Font(bold=True)

                writer.save()

            if 'Sheet' in book.sheetnames:
                sheet = book.get_sheet_by_name('Sheet')
                book.remove_sheet(sheet)

            for order in seperatedNameOrderList:
                if order['orderNumber'] == prevOrderID:
                    if thirdOrder:
                        order['lineItems'][0]['customizations'][0]['value'] = order['lineItems'][0]['customizations'][2]['value']
                        order['lineItems'][0]['customizations'][3]['value'] = order['lineItems'][0]['customizations'][5]['value']
                        thirdOrder = False
                    else:
                        order['lineItems'][0]['customizations'][0]['value'] = order['lineItems'][0]['customizations'][1]['value']
                        order['lineItems'][0]['customizations'][3]['value'] = order['lineItems'][0]['customizations'][4]['value']
                        thirdOrder = True
                else:
                    thirdOrder = False

                prevOrderID = order['orderNumber']

                if len(order['lineItems'][0]['customizations'][0]['value'].split()) == 1 and len(order['lineItems'][0]['customizations'][1]['value'].split()) == 1:
                    if len(order['lineItems'][0]['customizations'][2]['value'].split()) == 1:
                        d = [[order['lineItems'][0]['customizations'][0]['value'] + ' ' +
                              order['lineItems'][0]['customizations'][1]['value'] + ' ' +
                              order['lineItems'][0]['customizations'][2]['value'],
                              order['orderNumber'],
                              order['customerEmail'],
                              order['billingAddress']['firstName'] + ' ' +
                              order['billingAddress']['lastName'],
                              order['billingAddress']['phone'],
                              order['lineItems'][0]['customizations'][0]['value'],
                              order['lineItems'][0]['customizations'][3]['value'],
                              order['lineItems'][0]['customizations'][6]['value'],
                              order['lineItems'][0]['customizations'][7]['value'],
                              order['lineItems'][0]['customizations'][8]['value'],
                              order['lineItems'][0]['customizations'][10]['value'],
                              order['lineItems'][0]['customizations'][11]['value'],
                              '']
                             ]
                    else:
                        d = [[order['lineItems'][0]['customizations'][0]['value'] + ' ' +
                              order['lineItems'][0]['customizations'][1]['value'],
                              order['orderNumber'],
                              order['customerEmail'],
                              order['billingAddress']['firstName'] + ' ' +
                              order['billingAddress']['lastName'],
                              order['billingAddress']['phone'],
                              order['lineItems'][0]['customizations'][0]['value'],
                              order['lineItems'][0]['customizations'][3]['value'],
                              order['lineItems'][0]['customizations'][6]['value'],
                              order['lineItems'][0]['customizations'][7]['value'],
                              order['lineItems'][0]['customizations'][8]['value'],
                              order['lineItems'][0]['customizations'][10]['value'],
                              order['lineItems'][0]['customizations'][11]['value'],
                              '']
                            ]
                elif len(order['lineItems'][0]['customizations'][0]['value'].split()) == 1 and len(order['lineItems'][0]['customizations'][2]['value'].split()) == 1:
                    if len(order['lineItems'][0]['customizations'][1]['value'].split()) == 1:
                        d = [[order['lineItems'][0]['customizations'][0]['value'] + ' ' +
                              order['lineItems'][0]['customizations'][1]['value'] + ' ' +
                              order['lineItems'][0]['customizations'][2]['value'],
                              order['orderNumber'],
                              order['customerEmail'],
                              order['billingAddress']['firstName'] + ' ' +
                              order['billingAddress']['lastName'],
                              order['billingAddress']['phone'],
                              order['lineItems'][0]['customizations'][0]['value'],
                              order['lineItems'][0]['customizations'][3]['value'],
                              order['lineItems'][0]['customizations'][6]['value'],
                              order['lineItems'][0]['customizations'][7]['value'],
                              order['lineItems'][0]['customizations'][8]['value'],
                              order['lineItems'][0]['customizations'][10]['value'],
                              order['lineItems'][0]['customizations'][11]['value'],
                              '']
                             ]
                    else:
                        d = [[order['lineItems'][0]['customizations'][0]['value'] + ' ' +
                              order['lineItems'][0]['customizations'][2]['value'],
                              order['orderNumber'],
                              order['customerEmail'],
                              order['billingAddress']['firstName'] + ' ' +
                              order['billingAddress']['lastName'],
                              order['billingAddress']['phone'],
                              order['lineItems'][0]['customizations'][0]['value'],
                              order['lineItems'][0]['customizations'][3]['value'],
                              order['lineItems'][0]['customizations'][6]['value'],
                              order['lineItems'][0]['customizations'][7]['value'],
                              order['lineItems'][0]['customizations'][8]['value'],
                              order['lineItems'][0]['customizations'][10]['value'],
                              order['lineItems'][0]['customizations'][11]['value'],
                              '']
                             ]
                else:
                    d = [[order['lineItems'][0]['customizations'][0]['value'],
                          order['orderNumber'],
                          order['customerEmail'],
                          order['billingAddress']['firstName'] + ' ' +
                          order['billingAddress']['lastName'],
                          order['billingAddress']['phone'],
                          order['lineItems'][0]['customizations'][0]['value'],
                          order['lineItems'][0]['customizations'][3]['value'],
                          order['lineItems'][0]['customizations'][6]['value'],
                          order['lineItems'][0]['customizations'][7]['value'],
                          order['lineItems'][0]['customizations'][8]['value'],
                          order['lineItems'][0]['customizations'][10]['value'],
                          order['lineItems'][0]['customizations'][11]['value'],
                          '']
                         ]

                for k in range(numClasses+27):
                    d[0].insert(1, '')

                orderDataFrame = pd.DataFrame(data=d)

                if 'st.' in order['lineItems'][0]['productName'].split('- ')[1].replace(' ', '').split(',')[0].lower():
                    if '\'' in order['lineItems'][0]['productName'].split('- ')[1].replace(' ', '').split(',')[0].lower():
                        venue = order['lineItems'][0]['productName'].split('- ')[1].replace(' ', '').split(',')[0].lower().split('st.')[1].split('\'')[0]
                    else:
                        venue = order['lineItems'][0]['productName'].split('- ')[1].replace(' ', '').split(',')[0].lower().split('st.')[1][0:4]
                elif 'the' in order['lineItems'][0]['productName'].split('- ')[1].replace(' ', '').split(',')[0].lower():
                    venue = order['lineItems'][0]['productName'].split('- ')[1].replace(' ', '').split(',')[0].lower().split('the')[1][0:4]
                else:
                    venue = order['lineItems'][0]['productName'].split('- ')[1].replace(' ', '').split(',')[0][0:4].lower()
                day = order['lineItems'][0]['variantOptions'][0]['value'][0:3]
                time = order['lineItems'][0]['variantOptions'][0]['value'].split(',')[1].replace(' ','')[0:5].replace(':','_')

                sheetName = venue + '_' + day + '_' + time

                if sheetName not in book.sheetnames:
                    writer = pd.ExcelWriter(fileName, engine='openpyxl')
                    writer.book = book

                    templateDataFrame = pd.DataFrame(data=classListTemplateLines)
                    templateDataFrame.to_excel(writer, sheet_name=sheetName, header=False, index=False)

                    sheet = book[sheetName]
                    cell = sheet['A1']
                    cell.font = Font(bold=True)

                    for col in range(1, sheet.max_column+1):
                        colLetter = get_column_letter(col)
                        cell = sheet[colLetter+'2']
                        cell.font = Font(bold=True)
                    writer.save()

                AppendDfToExcel(fileName, orderDataFrame, sheetName, book)

            book = load_workbook(fileName)

            SortWorkSheets(book)

            row = 2
            totalStudents = 0

            for worksheet in book.worksheets:
                currentCell = book.get_sheet_by_name('Summary').cell(row=row, column=1)
                if worksheet.title != 'Summary':
                    currentCell.value = worksheet.title
                    currentCell.hyperlink = '#%s!%s' % (worksheet.title, 'A1')

                    currentCell = book.get_sheet_by_name('Summary').cell(row=row, column=2)
                    currentCell.value = worksheet.max_row-2

                    totalStudents += worksheet.max_row-2
                    row += 1

            currentCell = book.get_sheet_by_name('Summary').cell(row=row+1, column=1)
            currentCell.value = 'Total'
            currentCell.font = Font(bold=True)
            currentCell = book.get_sheet_by_name('Summary').cell(row=row+1, column=2)
            currentCell.value = totalStudents
            currentCell.font = Font(bold=True)

            book.save(fileName)

def SortWorkSheets(book):
    book._sheets.sort(key=lambda ws: ws.title)
    summarySheetIndex = book.worksheets.index(book.get_sheet_by_name('Summary'))
    newOrder = [i for i in range(len(book.worksheets))]
    del newOrder[summarySheetIndex]
    newOrder.insert(0, summarySheetIndex)
    book._sheets = [book._sheets[i] for i in newOrder]

def DeleteOldFileFromGoogleDrive(drive, classListName):
    fileList = drive.ListFile({'q': "'root' in parents and trashed=false"}).GetList()

    for file in fileList:
        if 'Class lists' in file['title']:
            folderContentList = drive.ListFile({'q': "'{}' in parents and trashed=false".format(file['id'])}).GetList()

    for file in folderContentList:
        if classListName in file['title']:
            file.Delete()
            print('Deleting ' + file['title'] + ' from Google Drive.')
            break

def UploadToGoogleDrive(drive, classListName):
    fileList = drive.ListFile({'q': "'root' in parents and trashed=false"}).GetList()

    for file in fileList:
        if 'Class lists' in file['title']:
            folderId = file['id']
            fileMetaData = {'title': classListName + '.xlsx', "parents": [{"id": folderId, "kind": "drive#childList"}]}
            folder = drive.CreateFile(fileMetaData)
            folder.SetContentFile(classListName + '.xlsx')  # The contents of the file
            print('Uploading ' + file['title'] + ' from Google Drive.')
            folder.Upload({'convert': True})
            break

def SendErrorEmails():
    print('Sending error emails.')

    subject = "Error Generating Class Lists"
    body = "This is an automated email that was sent to inform you of an error in generating the class lists.\n\nThis is possibly due to a problem with the product form.\n\nPlease see the attachment for more information"
    senderEmail = "robert@theacademyofcode.com"
    receiverEmail = "robert@theacademyofcode.com"

    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = senderEmail
    message["To"] = receiverEmail
    message["Subject"] = subject
    message["Bcc"] = receiverEmail  # Recommended for mass emails

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    filename = "ProductFormErrors.txt"  # In same directory as script

    # Open PDF file in binary mode
    with open(filename, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )

    # Add attachment to message and convert message to string
    message.attach(part)
    text = message.as_string()

    # Log in to server using secure context and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(senderEmail, emailPassword)
        server.sendmail(senderEmail, receiverEmail, text)

def main():
    start = time.time()

    ReadLastGenerationDate()

    allOrdersList = ExportAllOrders(startDate, endDate)
    individualOrdersList, classTypeList, fullYearList, abortClassListGeneration = ExportIndividualOrders(allOrdersList)

    if not abortClassListGeneration:
        if len(individualOrdersList) > 0:
            drive = GoogleDriveAccess()
            for classType in classTypeList:
                if classType[0] == 'Summer':
                    classListName = 'Summer Camps ' + classType[1]
                else:
                    if len(classType) > 2 and classType[2] == 'Tech Club':
                        classListName = classType[2] + classType[0] + ' ' + classType[1]
                    else:
                        classListName = 'Evening&Weekends ' + classType[0] + ' ' + classType[1]

                if testMode:
                    classListName = 'Test ' + classListName

                DownloadClassListsFromGoogleDrive(drive, classListName)
                if len(classType) > 2 and classType[2] == 'Next Term':
                    CreateAndAppendClassLists(fullYearList, classListName)
                else:
                    CreateAndAppendClassLists(individualOrdersList, classListName)
                DeleteOldFileFromGoogleDrive(drive, classListName)
                UploadToGoogleDrive(drive, classListName)

        WriteLastGenerationDate()
    else:
        print('Aborting class list generation.')
        SendErrorEmails()

    # orderCSVFiles = ListCSVFilesInDirectory(ordersDirectory)
    #
    # for file in orderCSVFiles:
    #     data = ReadCSVFile(ordersDirectory + file)

        #WriteCSVFile(classListDirectory, data)

    end = time.time()

    print(str(round(end - start, 2)) + ' secs')

if __name__ == '__main__':
    main()