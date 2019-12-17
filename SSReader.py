# SSReader
# Author: Low, Jimmy (lemoning.low@gmail.com)
# Date Created: 28 November 2019
# Last Edit: 13 December 2019
#
# This script reads an Excel Spreadsheet with a set of key column and value column, and convert them into IOS string
# resource file 'Localizable.strings' or Android string resource file 'strings.xml'.
# The script allows customized placeholders and replace them with the IOS / Android acceptable format (%@ / %s).
# This script takes in multiple necessary parameters such as file path, column of key, column of value, target type.
#
# Limitation: The translation is done one sheet at a time. If there are multiple sheet to be translated in a single
# spreadsheet file, the result must be concatenated manually.

import getopt, sys, os
import xlrd


def main():
    # TODO: module argparse to parse command line arguments
    # prepare to accept cmd arguments
    fullCmdArguments = sys.argv
    argumentList = fullCmdArguments[1:]
    unixOptons = "c:r:x:t:f:s:p:"
    gnuOptions = ["row=", "columnKey=", "columnValue=", "target=", "file=", "sheet=", "placeholder="]
    try:
        arguments, values = getopt.getopt(argumentList, unixOptons, gnuOptions)
    except getopt.error as error:
        print (str(error))
        sys.exit(2)

    # TODO: put the file handling code into one function
    # Default column and row index
    columnKey = 'a'
    row = 1
    columnValue = 'b'
    target = "android"
    file = ""
    sheetIndex = 0
    initialPlaceHolder = "{}"

    for currentArgument, currentValue in arguments:
        if currentArgument in ("-c", "--columnKey"):
            columnKey = currentValue
        elif currentArgument in ("-r", "--row"):
            row = currentValue
        elif currentArgument in ("-x", "--columnValue"):
            columnValue = currentValue
        elif currentArgument in ("-t", "--target"):
            target = currentValue
        elif currentArgument in ("-f", "--file"):
            file = currentValue
        elif currentArgument in ("-s", "--sheet"):
            sheetIndex = currentValue
        elif currentArgument in ("-p", "--placeholder"):
            initialPlaceHolder = currentValue
    if len(arguments) < 1:
        print("""
            -h / --help: Display list of available arguments

            -r<arg> / --row=<arg>: Pass the row of the first key. Must be > 1. Default is 1 (Example: 1)

            -c<arg> / --columnKey=<arg>: Pass the column of the first key. Default is A (Example A, e)

            -x<arg> / --columnValue=<arg>: Pass the column of the first value: Default is A (Example A, e)

            -t<arg> / --target=<arg>: Pass the type of the target. Default is Android (IOS / Android)

            -f<arg> / --file=<arg>: Pass the name of the spreadsheet file. This argument is compulsory

            -s<arg> / --sheet=<arg>: Pass the sheet index of the file. Default is 0. (Example: 0)

            -p<arg> / --placeholder<arg>: Pass the placeHolder. Default is {} (Example: {}, <>)
            """)
        return

    #  check input.
    try:
        if not columnKey.isalpha():
            raise Exception("Error: columnKey should only be made out of alphabets.")
        if not columnValue.isalpha():
            raise Exception("Error: columnValue should only be made out of alphabets.")
        row = int(row) - 1
        if row < 0:
            raise Exception("Error: row must be bigger or equal 1.")
        sheetIndex = int(sheetIndex)
        target = target.lower()
        if target != "android" and target != "ios":
            raise Exception("Error: target only accept input 'IOS' or 'Android' (Not case sensitive).")
        if len(file) <= 0:
            raise Exception("Error: you must provide a file path.")
    except Exception as error:
        print(error)
        sys.exit(2)

    convert_file(columnKey, row, columnValue, target, file, sheetIndex, initialPlaceHolder)
    print("Conversion Completed!")


def convert_file(columnKey, row, columnValue, target, file, sheetIndex, initialPlaceHolder):

    # convert alphabets to sheet column index in int.
    columnValue, columnKey = ord(columnValue.lower()) - 97, ord(columnKey.lower()) - 97
    finalPlaceHolder = ""

    # TODO: is there better way to get the path?
    current_directory = os.path.dirname(os.path.realpath(__file__))
    file = (current_directory + "/" + file)
    workbook = xlrd.open_workbook(file)
    sheet = workbook.sheet_by_index(sheetIndex)

    if target == "ios":
        iosFile = open("Localizable.strings", 'w')
        finalPlaceHolder = "%@"
    elif target == "android":
        androidFile = open("strings.xml", 'w')
        androidFile.write("<resources>\n")
        finalPlaceHolder = "%s"

    for rowIndex in range(row, sheet.nrows):
        if sheet.cell_value(rowIndex, columnKey) != "":

            # if the cell contains number, convert to string
            value = str(sheet.cell_value(rowIndex, columnValue))
            value = special_character_replacement(value, target)

            if '{' in value:
                value = placeholder_replacement(sheet.cell_value(rowIndex, columnValue), initialPlaceHolder,
                                                finalPlaceHolder)

            # TODO: use the .format() function instead
            if target == "ios":
                writeValue = '\"' + sheet.cell_value(rowIndex, columnKey) + "\" = \"" + value + "\";\n"
                iosFile.write(writeValue)
            if target == "android":
                androidFile.write("\t<string name=\"" + sheet.cell_value(rowIndex, columnKey) + "\">"
                                  + value + "</string>\n")

    if target == "ios":
        iosFile.close()
    elif target == "android":
        androidFile.write("</resources>\n")
        androidFile.close()


def placeholder_replacement(message, initialPlaceHolder, finalPlaceHolder):
    initialPlaceHolderOpen = initialPlaceHolder[0]
    initialPlaceHolderClose = initialPlaceHolder[1]
    openIndex = -1
    iterator = 0
    while iterator < len(message):
        if message[iterator] == initialPlaceHolderOpen:
            openIndex = iterator
            iterator += 1
        # check if there is already a placeholder opening
        elif message[iterator] == initialPlaceHolderClose and openIndex != -1:
            closeIndex = iterator
            message = message.replace(message[openIndex:closeIndex + 1], finalPlaceHolder)
            # reset iterator after replacement to openIndex so nothing to be missed
            iterator = openIndex + 1
            openIndex = -1
        else:
            iterator += 1
    return message


def special_character_replacement(message, target):
    # TODO: Use regex instead of the iteration
    if target == "android":
        if '&' in message:
            message = message.replace('&', '&amp;')
        if '\'' in message:
            message = message.replace('\'', "\\\'")
    return message


# run the main function
if __name__ == "__main__":
     main()
