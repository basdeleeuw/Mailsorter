import os
import win32com.client as win32

MAILPATH = "./Mails"
HEADERPATH = "./headers.txt"
SAVEPATH = "./mail.oft"

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
def save_errors(filelist, errorlist, headerlist):
    print("You forgot to implement the save_errors routine you dummy")


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

save_errors(files, errors, headers)
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
