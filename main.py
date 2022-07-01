import datetime
import Bimail

today = datetime.date.today()
yesterday = today - datetime.timedelta(days=1)

#whitePath = "C:/CMD/KST_CHECK/KST_EXPORT_" + yesterday.strftime('%Y_%m_%d') + ".csv"
#importPath = "C:/LX_Import/" + yesterday.strftime('%Y-%m-%d') + "/"
#aufKSTPath = importPath + "Import_Aufteilung_Kostenstellen_" + yesterday.strftime('%Y-%m-%d') + ".csv"
#lohnartenPath = importPath + "Import_Lohnarten_" + yesterday.strftime('%Y-%m-%d') + ".csv"
#zeitenPath = importPath + "Import_Zeiten_" + yesterday.strftime('%Y-%m-%d') + ".csv"

whitePath = ''
importPath = ''
aufKSTPath = ''
lohnartenPath = ''
zeitenPath = ''

recipients = []

subj = "moonImport - KST-überprüfung"

whiteList = []

files = {}

#files = {"aufKST": [aufKSTPath, 4, 0, 1, 3],
#         "lohnarten": [lohnartenPath, 14, 0, 1, 12],
#         "zeiten": [zeitenPath, 8, 0, 1, 3]}

msgs = []
msg = ""
teiler = "\n\n---------------------------------------------------------------\n\n"

readConfSuccess = True
readWhiteSucess = True


def checkImport(file, name):
    global msg

    new_rows = []

    print("---------------------" + name + "-------------------------------")

    with open(file[0], 'r') as f:
        csvlines = f.readlines()

        msg += "<br>Datei " + file[0] + "gefunden"

    for row in range(0, len(csvlines)):

        new_row = csvlines[row]
        elems = csvlines[row].split(';')

        #Erste Zeile (Header) Auslassen
        if row != 0:

            if elems[file[1]] not in whiteList:
                print(elems[file[1]] + " is NOT whitelisted")
                msgs.append(name + "  -- Firma: " + elems[file[2]] + " PersNR: " + elems[file[3]] + " Tag: " + elems[file[4]] + " -- KST " + elems[file[1]] + " nicht gefunden! <br>")

                count = 0
                new_row = ""

                #Zeile neu basteln, nur fündiges ersetzten durch DUMMY
                for e in elems:
                    if count != file[1]:
                        new_row = new_row + e.rstrip()
                    else:
                        new_row = new_row + "DUMMY"

                    count = count + 1

                    if count != len(elems):
                        new_row = new_row + ";"
                    else:
                        new_row = new_row + "\n"

        new_rows.append(new_row)

    with open(file[0], 'w') as f:
        for row in new_rows:
            f.write(row)


def readWhitelist():

    global msg
    global readWhiteSucess

    msg += "Versuche Whitlist zu laden - " + whitePath + "\n"

    try:
        with open(whitePath) as f:
            csvlines = f.readlines()

            for x in range(1, len(csvlines)):
                row = csvlines[x].split(';')

                whiteList.append(row[0])

            msg += "Whitelist erfolgreich gelesen - " + whitePath + "\n"
            msg += f.read()
            msg += teiler

    except Exception as e:
        msg += "Whitelist konnte nicht gelesen werden!\n Programm wird beendet"
        msg + str(e)
        readWhiteSucess = False


def sendEmail():
    mymail = Bimail.Bimail(subj, recipients)
    text = ""

    for msg in msgs:
        text = text + msg

    text = text + "<br><br>Alle KST durch 'DUMMY' ersetzt."

    print(text)

    mymail.htmladd(text)

    mymail.send()

# LEGACY
# def loadEmailRecipients():
#    with open("C:/CMD/moonKSTcheckMail.csv") as f:
#        csvlines = f.readlines()
#        for x in range(0, len(csvlines)):
#            recipients.append(csvlines[x].rstrip())
#
#    print(recipients)

def loadConfig():
    global recipients
    global whitePath
    global importPath
    global aufKSTPath
    global lohnartenPath
    global zeitenPath
    global readConfSuccess
    global msg

    today = datetime.date.today()

    msg += "Versuche config-File zu lesen - ./config.csv\n\n"

    try:
        with open('./config.csv', 'r') as f:
            lines = f.readlines()
            for x in range(len(lines)):
                lineElems = lines[x].split(';')

                #EmailEmpfänger einlesen - ganze Zeile, adressen getrennt durch ;
                if x == 0:
                    recipients = lineElems
                    msg += str(recipients) + '\n'

                if x == 1:
                    day = today - datetime.timedelta(days=-1*int(lineElems[2]))
                    whitePath = lineElems[1] + day.strftime(lineElems[3]) + lineElems[4]
                    msg += whitePath + '\n'

                if x == 2:
                    day = today - datetime.timedelta(days=-1*int(lineElems[2]))
                    importPath = lineElems[1] + day.strftime(lineElems[3]) + lineElems[4]
                    msg += importPath + '\n'

                if x == 3:
                    day = today - datetime.timedelta(days=-1*int(lineElems[2]))
                    aufKSTPath = importPath + lineElems[1] + day.strftime(lineElems[3]) + lineElems[4]
                    msg += aufKSTPath + '\n'

                if x == 4:
                    day = today - datetime.timedelta(days=-1*int(lineElems[2]))
                    lohnartenPath = importPath + lineElems[1] + day.strftime(lineElems[3]) + lineElems[4]
                    msg += lohnartenPath + '\n'

                if x == 5:
                    day = today - datetime.timedelta(days=-1*int(lineElems[2]))
                    zeitenPath = importPath + lineElems[1] + day.strftime(lineElems[3]) + lineElems[4]
                    msg += zeitenPath

                #if x == 6: überschrift für Colums

                if x == 7 or x == 8 or x == 9:
                    files[lineElems[0]] = [lineElems[1], lineElems[2], lineElems[3], lineElems[4]]

                # nur für msg DOKU
                if x == 9:
                    msg += str(files) + '\n'

        msg += "Config geladen"
        msg += teiler

    except Exception as e:

        msg += "Konnte Config nicht in Variablen laden\n"
        msg += str(e)
        msg += teiler
        readConfSuccess = False

def main():

    loadConfig()

    if readConfSuccess:
        readWhitelist()

    if readWhiteSucess and readConfSuccess:
        checkImport(files["aufteilungFile"], "Aufteilung Kostenstellen")
        checkImport(files["lohnartenFile"], "Lohnarten")
        #checkImport(files["zeitenFile"], "Zeiten")

    #sendEmail()

    print(msg)

main()
