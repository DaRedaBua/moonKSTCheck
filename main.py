import datetime
import Bimail
import sys, os
import traceback

today = datetime.date.today()
yesterday = today - datetime.timedelta(days=1)

whitePath = ''
importPath = ''
aufKSTPath = ''
lohnartenPath = ''
zeitenPath = ''

outAufPath = []
outLohnartenPath = []
outZeitenPath = []

recipients = []

subj = "moonImport - KST-überprüfung"

whiteList = []

files = {}

msgs = []
msg = ""
teiler = "<br><br>---------------------------------------------------------------<br><br>"

readConfSuccess = True
readWhiteSucess = True


def checkImport(path, cols, name, outpath):
    global msg

    msg += path + '<br>'
    msg += str(cols) + '<br>'
    msg += name + '<br>'
    msg += str(outpath) + '<br>'

    new_rows = []

    opensuccess = False
    replacesuccess = False

    msgs.append("<br>")

    if not cols[4] == "IGNORE":

        try:
            with open(path, 'r') as f:
                csvlines = f.readlines()

                msg += "<br>" + name + "-Datei erfolgreich geladen: <br>      " + path + '<br>'
                msgs.append("<br>" + name + "-Datei erfolgreich geladen: <br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp" + path + '<br>')
                opensuccess = True

        except Exception as e:
            msg += "Konnte Datei " + name + " nicht öffnen - " + path + "<br>"
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            msg += str(exc_type, fname, str(exc_tb.tb_lineno)) + '<br>'

        msg += "Opensuccess: " + str(opensuccess) + '<br>'

        #try:
        if opensuccess:
            msg += "Starte Ersetzungsvorgang" + "<br>"
            msgs.append("&nbsp&nbsp&nbsp&nbsp&nbsp&nbspStarte Ersetzunngsvorgang<br>")

            for row in range(1, len(csvlines)):

                new_row = csvlines[row]
                #msg += new_row + '<br>'

                if cols[4] == 'CHECK':

                    elems = csvlines[row].split(';')

                    if elems[cols[0]] not in whiteList:
                        msgs.append(name + "  -- Firma: " + elems[cols[1]] + " PersNR: " + elems[cols[2]] + " Tag: " + elems[cols[3]] + " -- KST " + elems[cols[0]] + " nicht gefunden - KST nicht in Whitelist! <br>")
                        msg += name + "  -- Firma: " + elems[cols[1]] + " PersNR: " + elems[cols[2]] + " Tag: " + elems[cols[3]] + " -- KST " + elems[cols[0]] + " nicht gefunden - KST nicht in Whitelist! <br>"

                        count = 0
                        new_row = ""

                        #Zeile neu basteln, nur fündiges ersetzten durch DUMMY
                        for e in elems:
                            if count != cols[0]:
                                new_row = new_row + e.rstrip()
                            else:
                                new_row = new_row + "DUMMY"

                            count = count + 1

                            if count != len(elems):
                                new_row = new_row + ";"
                            else:
                                new_row = new_row + "<br>"



                new_rows.append(new_row)

            msg += "<br><br>Ersetzungsvorgang für " + name + " abgeschlossen"
            replacesuccess = True

        msg += "<br>Replacesuccess: " + str(replacesuccess) + '<br>'

        if replacesuccess and (cols[4] == "CHECK" or cols[4] == "COPY"):
            try:
                msg += "Schreibe datei in outpath - " + outpath + "<br>"

                with open(outpath , 'w') as f:
                    for row in new_rows:
                        f.write(row)

                f.close()

                msgs.append("&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp" + name + "-Datei erfolgreich gespeichert:<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp" + outpath + '<br>')
                msg += name + "-Datei erfolgreich gespeichert:<br>      " + outpath + '<br>'

            except Exception as e:
                print(outpath)
                msg += "Konnte Datei nicht speichern - ECHT " + outpath + '<br>'
                msg += traceback.format_exc()

            try:
                msg += "Schreibe Daten zurück in ouriginaldatei - " + path + '<br>'

                with open(path, 'w') as f:
                    for row in new_rows:
                        f.write(row)

                f.close()

                msgs.append("&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp" + name + "-Datei erfolgreich zurückgeschrieben:<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp" + path + '<br>')
                msg += name + "-Datei erfolgreich zurückgeschrieben:<br>      " + path + '<br>'

            except Exception as e:
                msg += "Konnte Datei nicht speichern - ORIGINALDATEI " + path + '<br>'
                exc_type, exc_obj, exc_tb = sys.exc_info()
                fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                msg += str(exc_type, fname, str(exc_tb.tb_lineno)) + '<br>'

            msg += teiler


def readWhitelist():

    global msg
    global readWhiteSucess

    msg += "Versuche Whitlist zu laden - " + whitePath + "<br>"

    try:
        with open(whitePath) as f:
            csvlines = f.readlines()

            for x in range(1, len(csvlines)):
                row = csvlines[x].split(';')

                whiteList.append(row[0])

            msg += "Whitelist erfolgreich gelesen - " + whitePath + "<br>"
            msgs.append("Whitelist erfolgreich gelesen - " + whitePath + "<br>")
            msg += str(whiteList)
            msg += f.read()

    except Exception as e:
        msg += "Whitelist konnte nicht gelesen werden!<br> Programm wird beendet"
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        msg += str(exc_type, fname, str(exc_tb.tb_lineno)) + '<br>'
        readWhiteSucess = False

    finally:
        msg += teiler


def sendEmail():
    global msg

    try:
        msg += "Sende Email an " + str(recipients) + "..."

        mymail = Bimail.Bimail(subj, recipients)
        text = ""

        for smsg in msgs:
            text = text + smsg

        msg = text + teiler + msg

        mymail.htmladd(msg)

        mymail.send()

    except Exception as e:
        msg += "Konnte E-Mail nicht versenden..."
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        msg += str(exc_type, fname, str(exc_tb.tb_lineno)) + '<br>'

    finally:
        msg += teiler


def loadConfig():
    global recipients
    global whitePath
    global importPath
    global aufKSTPath
    global lohnartenPath
    global zeitenPath
    global readConfSuccess
    global outLohnartenPath
    global outZeitenPath
    global outAufPath
    global msg

    today = datetime.date.today()

    msg += "Versuche config-File zu lesen - ./config.csv<br><br>"

    try:
        with open('./config.csv', 'r') as f:
            lines = f.readlines()
            for x in range(len(lines)):

                lineElems = str.rstrip(lines[x]).split(';')

                #EmailEmpfänger einlesen - ganze Zeile, adressen getrennt durch ;
                if x == 0:
                    recipients = lineElems
                    recipients = list(filter(None, recipients))
                    msg += str(recipients) + '<br>'

                if x == 1:
                    day = today - datetime.timedelta(days=-1*int(lineElems[2]))
                    whitePath = lineElems[1] + day.strftime(lineElems[3]) + lineElems[4]
                    msg += whitePath + '<br>'

                if x == 2:
                    day = today - datetime.timedelta(days=-1*int(lineElems[2]))
                    importPath = lineElems[1] + day.strftime(lineElems[3]) + lineElems[4]
                    msg += importPath + '<br>'

                if x == 3:
                    day = today - datetime.timedelta(days=-1*int(lineElems[2]))
                    aufKSTPath = importPath + lineElems[1] + day.strftime(lineElems[3]) + lineElems[4]
                    msg += aufKSTPath + '<br>'

                if x == 4:
                    day = today - datetime.timedelta(days=-1*int(lineElems[2]))
                    lohnartenPath = importPath + lineElems[1] + day.strftime(lineElems[3]) + lineElems[4]
                    msg += lohnartenPath + '<br>'

                if x == 5:
                    day = today - datetime.timedelta(days=-1*int(lineElems[2]))
                    zeitenPath = importPath + lineElems[1] + day.strftime(lineElems[3]) + lineElems[4]
                    msg += zeitenPath

                # if x == 6: überschrift für Colums

                if x == 7 or x == 8 or x == 9:
                    files[lineElems[0]] = [int(lineElems[1]), int(lineElems[2]), int(lineElems[3]), int(lineElems[4]), lineElems[5]]

                # nur für msg DOKU
                if x == 9:
                    msg += str(files) + '<br>'

                # if x == 10: überschrift für outFiles

                if x == 11:
                    day = today - datetime.timedelta(days=-1 * int(lineElems[2]))
                    outAufPath = lineElems[1] + '/' + lineElems[0] + day.strftime(lineElems[3]) + lineElems[4]
                    msg += outAufPath

                if x == 12:
                    day = today - datetime.timedelta(days=-1 * int(lineElems[2]))
                    outLohnartenPath = lineElems[1] + '/' + lineElems[0] + day.strftime(lineElems[3]) + lineElems[4]
                    msg += outLohnartenPath

                if x == 13:
                    day = today - datetime.timedelta(days=-1 * int(lineElems[2]))
                    outZeitenPath = lineElems[1] + '/' + lineElems[0] + day.strftime(lineElems[3]) + lineElems[4]
                    msg += outZeitenPath

        msg += "Config geladen"
        msgs.append("Config File geladen - ./config.csv <br><br>")

        msg += "recipients - " + str(recipients)
        msg += "whitelistPath - " + whitePath
        msg += "importPath - " + importPath
        msg += "aufKSTPath - " + aufKSTPath
        msg += "lohnartenPath - " + lohnartenPath
        msg += "zeitenPath - " + zeitenPath + '<br>'
        msg += str(files) + '<br>'
        msg += str(outAufPath) + '<br>'
        msg += str(outLohnartenPath) + '<br>'
        msg += str(outZeitenPath) + '<br>'

        msg += teiler


    except Exception as e:

        msg += "Konnte Config nicht in Variablen laden<br>"
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        msg += str(exc_type, fname, str(exc_tb.tb_lineno)) + '<br>'
        msg += teiler
        readConfSuccess = False


def main():

    global msg

    loadConfig()

    if readConfSuccess:
        readWhitelist()

    if readWhiteSucess and readConfSuccess:
        checkImport(aufKSTPath, files["aufteilungKSTs"], "Aufteilung Kostenstellen", outAufPath)
        checkImport(lohnartenPath, files["lohnarten"], "Lohnarten", outLohnartenPath)
        checkImport(zeitenPath, files["zeiten"], "Zeiten", outZeitenPath)

    sendEmail()

    msg = msg.replace('<br>', '\n')
    print(msg)

main()
