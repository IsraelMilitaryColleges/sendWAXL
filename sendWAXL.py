# import load_workbook
from openpyxl import load_workbook
import urllib.parse
import webbrowser

def message_file_content():
     with open('message.txt', 'r') as txt_file:
          return txt_file.read()

def generate_message_from_file(name, user, password):
     message = message_file_content()
     message_parameters = {"NAME": name, "USER": user, "PASSWORD": password}
     for parameter in message_parameters:
          message = non_formatted_message.replace(parameter, message_parameters[parameter])
     return message

def generate_message(name, user, password):
     message = "שלום " + name + "!\n" + 'במידה ולא הצלחת להיכנס לסביבת הלמידה האזרחית, מצ"ב פרטי הגישה.' + "\n\n" + "user: " + user + "\n" + "password: " + password + "\n" + "טכנולוגיות למידה"
     return message

def send_whatsapp(phone_number, message):
     url_encoded_message = urllib.parse.quote(message)
     url = 'https://wa.me/' + phone_number + '?text=' + url_encoded_message
     webbrowser.open_new_tab(url)

def main():
    SPREADSHEET_FILEPATH = "organized_spreadsheet.xlsx"
    sheet = load_workbook(SPREADSHEET_FILEPATH).active
    MAX_ROW = sheet.MAX_ROW

    NAME_COLUMN_INDEX = 1
    USER_COLUMN_INDEX = 2
    PASSWORD_COLUMN_INDEX = 3
    PHONE_COLUMN_INDEX = 4

    for line in range(1, MAX_ROW + 1):
         name = sheet.cell(row=line, column=NAME_COLUMN_INDEX).value
         user = sheet.cell(row=line, column=USER_COLUMN_INDEX).value
         password = sheet.cell(row=line, column=PASSWORD_COLUMN_INDEX).value
         phone_number = str(sheet.cell(row=line, column=PHONE_COLUMN_INDEX).value)

         # message = generate_message_from_file(name, user, password)
         message = generate_message(name, user, password)
         send_whatsapp(phone_number, message)

         print(phone_number + '\n')

if __name__ == '__main__':
    main()
