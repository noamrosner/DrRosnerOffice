
from bdikot import *

def getMaxRow():
    global maxRow
    maxRow = 2
    for i in range(2, sheet.max_row):
        if type(sheet[f"D{i}"].value) == str:
            maxRow = maxRow + 1
        else:
            if type(sheet[f"D{i+1}"].value) == str:
                maxRow = maxRow + 1
                continue
            if type(sheet[f"D{i+2}"].value) == str:
                maxRow = maxRow + 2
                continue
            else:
                break
    return maxRow

def formatFile():
    maxRow = getMaxRow()
    for i in range(2, maxRow):
        sheet[f"A{i}"].value = sheet["A2"].value
    for i in range(2, maxRow):
        sheet[f"B{i}"].value = sheet["B2"].value
    for i in range(2, maxRow):
        if type(sheet[f"D{i}"].value) != str:
            continue
        try:
            sheet[f"C{i}"].value = sheet[f"C{i}"].value.replace(" ", '')
            print(f"Changed C{i}")
        except TypeError:
            print(f"C{i} ; Spaces replacement gone wrong")
        except AttributeError:
            print(f"C{i} ; Spaces replacement gone wrong")
        try:
            sheet[f"G{i}"].value = sheet[f"G{i}"].value.replace("נייד: ", '')
        except AttributeError:
            print(f"G{i} ; Cell number replacement gone wrong")
    wb.save(filename)

def maccabiAskFile():
    global sheet
    global wb
    global filename
    global maxRow
    while True:
        try:
            filename = askopenfilename()
            wb = openpyxl.load_workbook(f'{filename}')
            sheet = wb['מכבי']
            print(f"Opening {filename}!\n")
            print(
                "\n--------------------------------------------------------------------------------------------------------"
                "-\n")
            return 1
        except openpyxl.utils.exceptions.InvalidFileException:
            print("Abort!")
            return 0
        except NameError:
            print("Wrong file, not a xlsx file")
        except TypeError:
            print("Wrong file, not a xlsx file")

def maacabiSendToList():
    global sheet
    global wb
    global filename
    global maxRow
    for i in range(2, maxRow):
        if sheet[f"E{i}"].value == "תור טלפוני":
            try:
                msg = phoneMessage(sheet[f"A{i}"].value.strftime("%d/%m/%Y"), sheet[f"B{i}"].value,
                                     sheet[f"C{i}"].value, sheet[f"D{i}"].value)
                validPhone(f'{sheet[f"F{i}"].value}')
                sendMessage(validPhone(sheet[f"G{i}"].value), msg)
                sheet[f"I{i}"].value = "Sent successfully"
            except Exception as e:
                print(e)
                sheet[f"I{i}"].value = "NOT SENT"
        else:
            try:
                msg = maccabiMessage(sheet[f"A{i}"].value.strftime("%d/%m/%Y"), sheet[f"B{i}"].value,
                                     sheet[f"C{i}"].value, sheet[f"D{i}"].value)
                validPhone(f'{sheet[f"F{i}"].value}')
                sendMessage(validPhone(sheet[f"G{i}"].value), msg)
                sheet[f"I{i}"].value = "Sent successfully"
            except Exception as e:
                print(e)
                sheet[f"I{i}"].value = "NOT SENT"
    wb.save(filename)

def maccabiMessage(date, day, time, name):
        msg = (f"\n"
               f" שלום {name},\n"
               f"זו תזכורת מהמרפאה של ד״ר רוזנר לתור "
               f"שנקבע לך ליום {day} {date} בשעה {time}\n"
               f"כתובת המרפאה: בניין רסיטל לשעבר, מנחם בגין 156 קומה  22 תל אביב.\n"
               f"\n*אין חניה בבניין!*"
               f"\n"
               f"אפשר לחנות בחניון עזריאלי טאון בכתובת מנחם בגין 146 או ברכבת, בחניון ארלוזרוב ולהגיע ברגל.\n"
               f"יש להגיע עם כרטיס מגנטי\n"
               f"\n להלן קישור למיקום המרפאה בוייז,נסיעה עם Waze:"
               f"\n"
               f"https://waze.com/ul/hsv8wrzcqj"
               f"\n"
               f"\n אודה לאישור התור"
               f"\n"
               f"\n"
               f"בברכה,\n"
               f"מרפאתו של ד״ר רוזנר ")
        return msg

def phoneMessage(date, day, time, name):
    msg = (f"\n"
           f" שלום {name},\n"
           f"זו תזכורת מהמרפאה של ד״ר רוזנר לתור טלפוני "
           f"שנקבע לך ליום {day} {date} בשעה {time}\n"
           f"\n"
           f"\n התור הינו תור טלפוני"
           f""
           f"\n אודה לאישור התור"
           f"\n"
           f"\n"
           f"בברכה,\n"
           f"מרפאתו של ד״ר רוזנר ")
    return msg
    
def office():
    if maccabiAskFile() == 1:
        formatFile()
        maacabiSendToList()
