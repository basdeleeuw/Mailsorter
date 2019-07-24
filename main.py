import os
import win32com.client as win32

MAILPATH = "./Mails"
HEADERPATH = "./headers.txt"
SAVEPATH = "./mail.oft"

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


# Checks of there are errors to be reported
def checkerrors(errors):
    reportable_errors = False
    for i in range(len(errors)):
        if ((i not in NOREPORT_HEADERS) and errors[i]):
            reportable_errors = True
    return reportable_errors


# Writes the errors to an e-mail
# IN:    filelist: A list of the file names of the errorlogs
#       errorlist: A matrix[headernr][linenr] of the errors in the e-mail
#       headerlist: A list of the error headers in a [headernr][2]-matrix
def mail_errors(filelist, errorlist, headerlist):
    print("You forgot to implement the save_errors routine you dummy")
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.to = MAILTO
    mail.subject = MAILSUBJECT
    mail.HTMLbody = '''<style type="text/css">
.tg  {border-collapse:collapse;border-spacing:0;}
.tg td{font-family:Arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:black;}
.tg th{font-family:Arial, sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;border-color:black;}
.tg .tg-5ua9{font-weight:bold;text-align:left}
.tg .tg-s268{text-align:left}
</style>
<table class="tg">
  <tr>
    <th class="tg-5ua9">Database</th>
    <th class="tg-5ua9">Melding</th>
  </tr>
  <tr>
    <td class="tg-s268"></td>
    <td class="tg-s268"></td>
  </tr>
</table>'''  # TODO: Fix the body
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

mail_errors(files, errors, headers)
