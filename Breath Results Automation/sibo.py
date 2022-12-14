import openpyxl
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image


answerH2 = """ ANSWER """

answerCH4 = """ ANSWER """

answerBoth = """ ANSWER """

answerNone = """ ANSWER """


def validInput(msg, options):
    while True:
        entered_value = input(msg)
        if entered_value in options:
            break
        else:
            print('Wrong input. Please try again')
            continue
    return entered_value

def greetings(file):
    global H2, CH4, sheet, wb, filename, label1, label2, label3
    try:
        filename = file
        wb = openpyxl.load_workbook(f'{filename}')
        sheet = wb['Lactulose SIBO']
        H2 = False
        CH4 = False
        print(f"Opening {filename}!\n")
        print("\n---------------------------------------------------------------------------------------------------------\n")
    except openpyxl.utils.exceptions.InvalidFileException:
        print("Openpyxl error")
    except NameError:
        print("Wrong file, not a xlsx file")
    except TypeError:
        print("Wrong file, not a xlsx file")

def getH2():
    global H2
    global minH2
    global maxH2
    indexListC = []
    for i in range(18,26):
        value = str('C' + str(i))
        indexListC.append(value)

    minH2 = sheet['C18'].value
    for index in indexListC: # find min H2 value
        try:
            if sheet[index].value < minH2:
                minH2 = sheet[index].value
        except TypeError:
            print("Value Error, some H2 data is missing")
            print(f"in index: {index}")
            print(f"sheet[index].value is {sheet[index].value}")
            print(f"minH2 is {minH2}\n")

    maxH2 = sheet['C18'].value
    for index in indexListC: # find max H2 value
        try:
            if sheet[index].value > maxH2:
                maxH2 = sheet[index].value
        except TypeError:
            print("Value Error, some H2 data is missing")
            print(f"in index: {index}")
            print(f"sheet[index].value is {sheet[index].value}")
            print(f"maxH2 is {maxH2}\n")

    if maxH2 - minH2 > 20:
        H2 = True

def getCH4():
    global CH4
    global minCH4
    global maxCH4
    indexListD = []
    for i in range(18,26):
        value = str('D' + str(i))
        indexListD.append(value)

    minCH4 = sheet['D18'].value
    for index in indexListD: # find min CH4 value
        try:
            if sheet[index].value < minCH4:
                minCH4 = sheet[index].value
        except TypeError:
            print("Value Error, some CH4 data is missing")
            print(f"in index: {index}")
            print(f"sheet[index].value is {sheet[index].value}")
            print(f"minCH4 is {minCH4}\n")

    maxCH4 = sheet['D18'].value
    for index in indexListD: # find max CH4 value
        try:
            if sheet[index].value > maxCH4:
                maxCH4 = sheet[index].value
        except TypeError:
            print("Value Error, some CH4 data is missing")
            print(f"in index: {index}")
            print(f"sheet[index].value is {sheet[index].value}")
            print(f"maxCH4 is {maxCH4}\n")

    if maxCH4 > 10: # or maxCH4 - minCH4 >= 15:
        CH4 = True

def checkValues():
    global pdfName
    if H2 and CH4:
        print("Changing the answer to 'High Value H2 & CH4'\n")
        sheet['A28'] = answerBoth
        sheet['A28'].alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
    elif H2:
        print("Changing the answer to 'High Value H2'\n")
        sheet['A28'] = answerH2
        sheet['A28'].alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
    elif CH4:
        print("Changing the answer to 'High Value CH4'\n")
        sheet['A28'] = answerCH4
        sheet['A28'].alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
    else:
        print("Changing the answer to normal.\n")
        sheet['A28'] = answerNone
        sheet['A28'].alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
    # wb.save(filename=filename)


def addImage():
    sheet = wb['Lactulose SIBO']
    img = Image("signature.jpg")
    img.width = 500
    img.height = 300
    sheet.add_image(img, 'B60')
    wb.save(filename=filename)

def generalKnowledge():
    patientID = str(sheet['B11'].value)
    patientName = str(sheet['B10'].value)
    dateOf = str(sheet['H12'].value)
    patientSex = str(sheet['B13'].value)
    print(f"Patient name is: {patientName}\n"
          f"ID: {patientID}\n"
          f"Sex: {patientSex}\n\n"
          f"Date of the Examination: {dateOf}\n\n"
          f"H2 Values:\n"
          f"minimum value: {minH2}\n"
          f"maximum value: {maxH2}\n"
          f"H2: {H2}\n\n"
          f"CH4 Values:\n"
          f"minimum value: {minCH4}\n"
          f"maximum value: {maxCH4}\n"
          f"CH4: {CH4}\n")

def validate():
    if sheet['C17'].value != 'H2':
        print("Bad values, contact Noam")
        print(f"Scanning for {sheet['C17'].value} instead of H2!")
        exit()
    if sheet['D17'].value != 'CH4':
        print("Bad values, contact Noam")
        print(f"Scanning for {sheet['D17'].value} instead of CH4!")
        exit()



def sibo(file):
    greetings(file)
    try:
        validate()
        getH2()
        getCH4()
        generalKnowledge()
        checkValues()
        addImage()
    except Exception as e:
        print(f"Error - {e}")
        print(f"Contact Noam")


if __name__ == "__main__":
    sibo()
