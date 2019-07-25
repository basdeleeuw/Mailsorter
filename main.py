import os
import win32com.client as win32

MAILPATH = "./Mails"
HEADERPATH = "./headers.txt"
SAVEPATH = "./mail.oft"
TABLEPATH = "./Table"

MAILTO = "oracledboper@icts.utwente.nl"
MAILSUBJECT = "Daily Check"

HEADERLIST_END = "==END=="
NOREPORT_HEADERS = [
    4, 5, 6
]  # Headers to be neglected. Note that the counting starts with 0, so the 4th header is actually number 3 here!


# Reads the headers from HEADERPATH and returns them as a matrix[headernr][linenr]
# OUT:   headers: a list ot the error-headers as a [headernr][2]-matrix
def readheaders():
    headerfile = open(HEADERPATH, 'r')
    headers = []
    line = headerfile.readline()

    while line:
        if str(line) != str("\n"):
            headers.append([line, headerfile.readline()])
            line = headerfile.readline()
        line = headerfile.readline()
    headerfile.close
    headers.append(HEADERLIST_END)
    return headers


# Reads the actual errors from an e-mail and returns them as a matrix[headernr][linenr]
# IN:    filename: The path to the file from which the errors are to be extracted
#        headerlist: A list of the error headers in a [headernr][2]-matrix
# OUT:   errors: a matrix[headernr][linenr] of the errors in the e-mail
def geterrors(filename, headerlist):
    file = open(MAILPATH + "/" + filename, 'r')
    errors = [[] for x in range(len(headerlist))]
    n_header = 0

    line = file.readline()
    while line and line != HEADERLIST_END:
        if str(line) == str(headerlist[n_header][0]):
            while ((str(line) != str(headerlist[n_header + 1][0])) and line):
                if ((str(line) != str("\n"))
                        and (str(line) != str(HEADERLIST_END))
                        and (str(line) != str(headerlist[n_header][1]))
                        and (str(line) != str(headerlist[n_header][0]))):
                    errors[n_header].append(line)
                line = file.readline()
            n_header = n_header + 1
        else:
            line = file.readline()

    file.close()
    return errors


# Checks if there are errors to be reported
def checkerrors(errors):
    reportable_errors = False
    for i in range(len(errors)):
        if ((i not in NOREPORT_HEADERS) and errors[i]):
            reportable_errors = True
    return reportable_errors


def tablerow(file, errors, headerlist):
    errorcell = ""
    for i in range (len(errors)):
        if ((i not in NOREPORT_HEADERS) and errors[i]):
            errorcell = errorcell + headerlist[i][0] + "<br>" + headerlist[i][1] + "<br><br>"
            errorhtml = errors[i]
            print(errorhtml)
            for j in range (len(errorhtml)):
                errorhtml[j] = errorhtml[j].replace("\n", "<br>")
                errorhtml[j] = errorhtml[j].replace(" ", "&nbsp;")
            print(errorhtml)
            errorhtml = "".join(errorhtml)
            print(errorhtml)
            errorcell = errorcell + errorhtml + "<br><br>"
    filename = file[:-8]
    row = '''<tr>
                <td class="tg-table">''' + filename + '''</td>
                <td class="tg-table">''' + errorcell + '''</td>
             </tr>'''
    return row

def composetables(filelist, errorlist, headerlist):
    tables = [""] * 2
    emptytable = open(TABLEPATH+"/EmptyTable.html").read()
    tableheader = open(TABLEPATH+"/TableHeader.html").read()

    tables[0] = '<table class="tg">' + tableheader
    for i in range (len(filelist)):
        if checkerrors(errorlist[i]):
            tables[0] = tables[0] + tablerow(filelist[i],errorlist[i],headerlist)
    tables[0] = tables[0] + "</table>"

    tables[1] = emptytable

    return tables

def composebody(tables):
    style = open(TABLEPATH+"/Style.html").read()

    body = style + "Hallo,<br><br>" + "Bij deze de meldingen van vandaag.<br><br>" + "<h2>Daily logs</h2>" + tables[0] + "<h2>Non-daily logs</h2>" + tables[1] + "<br><br>Groeten,<br> "
    return body

# Writes the errors to an e-mail
# IN:    filelist: A list of the file names of the errorlogs
#       errorlist: A matrix[headernr][linenr] of the errors in the e-mail
#       headerlist: A list of the error headers in a [headernr][2]-matrix
def mail_errors(body):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.to = MAILTO
    mail.subject = MAILSUBJECT
    mail.HTMLbody = body
    mail.Display(True)


# Prints the errors to the screen in a semi-readable way, mainly for debugging
# IN:    errors: a matrix[headernr][linenr] of the errors in the e-mail
def printerrors(errors):
    for i in range(len(errors)):
        for j in range(len(errors[i])):
            print(errors[i][j])


#####################################################################
##############################MAIN LOOP##############################
#####################################################################
files = os.listdir(MAILPATH)
headers = readheaders()
errors = []

for i in range(len(files)):
    errors.append(geterrors(files[i], headers))
    # printerrors(errorslist[i])

mailtables = composetables(files, errors, headers)

mailbody = composebody(mailtables)

mail_errors(mailbody)
