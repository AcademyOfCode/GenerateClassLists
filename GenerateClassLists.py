import os
import datetime
import dateutil.parser
import requests
import time
import argparse
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

parser = argparse.ArgumentParser()
parser.add_argument('apiKey')
parser.add_argument('testMode')
args = parser.parse_args()

apiKey = args.apiKey
apiURL = 'https://api.squarespace.com/1.0/commerce/orders'

startDate = '2019-02-07T11:20:00.000000Z'
endDate = datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%S.%fZ')

testMode = bool(args.testMode)

def GetDateTimeFromISO8601String(s):
    return dateutil.parser.parse(s)

def WriteLastGenerationDate():
    with open(os.path.dirname(os.path.abspath(__file__)) + '\LastClassListGenerationDate.txt', 'w') as file:
        file.write(endDate)

def ReadLastGenerationDate():
    with open(os.path.dirname(os.path.abspath(__file__)) + '\LastClassListGenerationDate.txt') as file:
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
        ('modifiedBefore', endDate)
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

                    if '30 Weeks' in response.json()['lineItems'][0]['variantOptions'][1]['value']:
                        classType = ('Spring', str(int(response.json()['lineItems'][0]['productName'].split()[1]) + 1), 'Next Term')
                        fullYearList.append(response.json())

                    if classType not in classTypeList:
                        classTypeList.append(classType)

                    responseList.append(response.json())
                    print('Order ' + str(orderCount) + ' Completed.')
                    orderCount += 1

            else:
                if 'Tech Club' in response.json()['lineItems'][0]['productName'].split('- ')[1]:
                    classType = (response.json()['lineItems'][0]['productName'].split()[0],
                                 response.json()['lineItems'][0]['productName'].split()[1], 'Tech Club')
                else:
                    classType = (response.json()['lineItems'][0]['productName'].split()[0],
                                 response.json()['lineItems'][0]['productName'].split()[1])

                if '30 Weeks' in response.json()['lineItems'][0]['variantOptions'][1]['value']:
                    classType = ('Spring', str(int(response.json()['lineItems'][0]['productName'].split()[1]) + 1), 'Next Term')
                    fullYearList.append(response.json())

                if classType not in classTypeList:
                    classTypeList.append(classType)

                responseList.append(response.json())
                print('Order ' + str(orderCount) + ' Completed.')
                orderCount += 1

    return responseList, classTypeList, fullYearList

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
    gauth.SaveCredentialsFile(os.path.dirname(os.path.abspath(__file__))+ "credentials.txt")

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

    for fileName in os.listdir(os.path.dirname(os.path.abspath(__file__))):
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

                sheetName = day + '_' + venue.capitalize() + '_' + time

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
    dayList = ['Mon','Tue','Wed','Thu','Fri','Sat', 'Sun']
    seperateDayList = []
    newOrder = []

    for day in dayList:
        sheetList = []
        for sheet in book.sheetnames:
            if day == sheet[0:3]:
                sheetList.append(sheet)
        if sheetList:
            seperateDayList.append(sheetList)

    for day in seperateDayList:
        day.sort()
    newClassOrderList = [j for i in seperateDayList for j in i]

    for sheetName in newClassOrderList:
        newOrder.append(book.worksheets.index(book.get_sheet_by_name(sheetName)))

    summarySheetIndex = book.worksheets.index(book.get_sheet_by_name('Summary'))
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

def main():
    start = time.time()

    ReadLastGenerationDate()

    allOrdersList = ExportAllOrders(startDate, endDate)
    individualOrdersList, classTypeList, fullYearList = ExportIndividualOrders(allOrdersList)

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

    end = time.time()

    print(str(round(end - start, 2)) + ' secs')

if __name__ == '__main__':
    main()