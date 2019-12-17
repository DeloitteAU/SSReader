# README

This script reads an Excel Spreadsheet with a set of key column and value column, and convert them into IOS string
resource file 'Localizable.strings' or Android string resource file 'strings.xml'. The script allows customized
placeholders and replace them with the IOS / Android acceptable format (%@ / %s).

This script takes in multiple necessary parameters such as file path, column of key, column of value, target type.

# Limitation
The translation is done one sheet at a time. If there are multiple sheet to be translated in a single spreadsheet file,
the result must be concatenated manually.

# Requirement
- User has Python 3.x installed.

# Procedure

1. Install Python 3.x, suggest to use HomeBrew to install so pip3 wil be installed along.
2. Migrate to the folder where the repository is cloned.
3. Install xlrd by using the following command:
pip3 install -r requirements.txt
4. Run the script without any argument to see the help list:
python3 SSReader.py
5. Use the arguments to flexibly adjust all the attributes and convert your file!
