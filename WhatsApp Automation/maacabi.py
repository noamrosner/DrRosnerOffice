
from bdikot import *

def fillList():
    sheet2 = wb['אחרי']
    j = 2
    try:
        for i in range(1, sheet.max_row-3, 6):
            lst1 = sheet[f"A{i}"].value.split()
            print(lst1)
            lst2 = sheet[f"A{i+1}"].value.split()
            print(lst2)
            lst3 = sheet[f"A{i + 2}"].value.split()
            print(lst3)
            sheet2[f"A{j}"].value = f"{lst1[0]}"
            sheet2[f"B{j}"].value = sheet2[f"B2"].value
            sheet2[f"C{j}"].value = sheet2[f"C2"].value
            sheet2[f"D{j}"].value = f"{lst1[1]}"
            sheet2[f"E{j}"].value = f"{lst1[2]}"
            sheet2[f"F{j}"].value = f"{lst1[-1]}"
            sheet2[f"G{j}"].value = f"{lst2[1]} {lst2[2]}"
            sheet2[f"H{j}"].value = f"{lst3[1]}"
            j = j+1
            print(j)
    except AttributeError:
        print(f"error with index {i}")
    wb.save(filename)


def formatFile():
    # fill first date, day manually
    for i in range(2, sheet.max_row):
        sheet[f"A{i}"].value = sheet["A2"].value
    for i in range(2, sheet.max_row):
        sheet[f"B{i}"].value = sheet["B2"].value
    for i in range(2, sheet.max_row):
        try:
            if sheet[f"C{i}"].value[0] == ' ':
                sheet[f"C{i}"].value = sheet[f"C{i}"].value[1:]
                print(f"Changed C{i}")
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
    while True:
        try:
            filename = askopenfilename()
            wb = openpyxl.load_workbook(f'{filename}')
            sheet = wb['לפני']
            print(f"Opening {filename}!\n")
            print(
                "\n--------------------------------------------------------------------------------------------------------"
                "-\n")
            break
        except openpyxl.utils.exceptions.InvalidFileException:
            print("Fix this")
        except NameError:
            print("Wrong file, not a xlsx file")
        except TypeError:
            print("Wrong file, not a xlsx file")

def maacabiSendToList():
    global sheet
    global wb
    global filename
    for i in range(2, sheet.max_row):
        try:
            msg = maccabiMessage(sheet[f"A{i}"].value.strftime("%d/%m/%Y"), sheet[f"B{i}"].value, sheet[f"C{i}"].value.strftime("%H:%M")
                                 , sheet[f"D{i}"].value)
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
           f"זו תזכורת מהמרפאה של ד״ר רוזנר לתור "
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

def cleanList():
    try:
        for i in range(2, sheet.max_row):
            val = sheet[f"G{i}"].value.split()
            sheet[f"G{i}"].value = val[1]
    except AttributeError:
        print(f"Error with index {i}")
        print(sheet.max_row)
    except Exception as e:
        print(e)
    wb.save(filename)


def compareList(list1, list2):
    print()
    
def office():
    maccabiAskFile()
    formatFile()
    #maacabiSendToList()

"""if __name__ == "__main__":
    office()"""
