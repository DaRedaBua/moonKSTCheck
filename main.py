import datetime
import Bimail

today = datetime.date.today()
yesterday = today - datetime.timedelta(days=1)

whitePath = "C:/CMD/KST_CHECK/KST_EXPORT_" + yesterday.strftime('%Y_%m_%d') + ".csv"
importPath = "C:/LX_Import/" + yesterday.strftime('%Y-%m-%d') + "/"
aufKSTPath = importPath + "Import_Aufteilung_Kostenstellen_" + yesterday.strftime('%Y-%m-%d') + ".csv"
lohnartenPath = importPath + "Import_Lohnarten_" + yesterday.strftime('%Y-%m-%d') + ".csv"
zeitenPath = importPath + "Import_Zeiten_" + yesterday.strftime('%Y-%m-%d') + ".csv"

recipients = []

subj = "KST-Fehler bei Import"

whiteList = []

files = {"aufKST": [aufKSTPath, 4, 0, 1, 3],
         "lohnarten": [lohnartenPath, 14, 0, 1, 12],
         "zeiten": [zeitenPath, 8, 0, 1, 3]}

msgs = []

def checkImport(file, name):
    new_rows = []

    print("---------------------" + name + "-------------------------------")

    with open(file[0], 'r') as f:
        csvlines = f.readlines()

    for row in range(0, len(csvlines)):

        new_row = csvlines[row]
        elems = csvlines[row].split(';')

        #Erste Zeile (Header) Auslassen
        if row != 0:
            # Nicht auf Whitelist
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
    with open(whitePath) as f:
        csvlines = f.readlines()

    for x in range(1, len(csvlines)):
        row = csvlines[x].split(';')

        whiteList.append(row[0])

def sendEmail():
    mymail = Bimail.Bimail(subj, recipients)
    text = ""

    for msg in msgs:
        text = text + msg

    text = text + "<br><br>Alle KST durch 'DUMMY' ersetzt."

    print(text)

    mymail.htmladd(text)

    mymail.send()


def loadEmailRecipients():
    with open("C:/CMD/moonKSTcheckMail.csv") as f:
        csvlines = f.readlines()
        for x in range(0, len(csvlines)):
            recipients.append(csvlines[x].rstrip())

    print(recipients)

def main():
    loadEmailRecipients()

    readWhitelist()

    checkImport(files["aufKST"], "Aufteilung Kostenstellen")
    checkImport(files["lohnarten"], "Lohnarten")
    checkImport(files["zeiten"], "Zeiten")

    sendEmail()

main()
