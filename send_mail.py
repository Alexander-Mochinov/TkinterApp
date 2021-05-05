from openpyxl import load_workbook
import sys
import smtplib
import re


#======================== GLOBAL VARIABLE ==============
NAME_TEXT_FILE = 'text_for_send.txt' 
#=======================================================

def send_mail(HOST,PORT,email, password, message, to_email) -> bool:
    """Отправка сообщения"""
    try:
        server = smtplib.SMTP(HOST, PORT)
        server.starttls()
        server.login(email, password)
        server.sendmail(email, to_email, message.encode('utf-8'))
        server.quit()
        return True
    except:
        return False
def set_message(list_params, message):
    """Замена сообщения текста в файле"""
    surname  = list_params[0]
    name = list_params[1]
    patronymic = list_params[2]
    to_the_DO_system = list_params[3]
    to_the_testing_system = list_params[4]
    to_MS_Teams = list_params[3]
    password = list_params[5]
    replace_values_list = {
        '{{surname}}' : surname,
        '{{name}}' : name,
        '{{patronymic}}' : patronymic,
        '{{to_the_DO_system}}' : to_the_DO_system,
        '{{to_the_testing_system}}' : to_the_testing_system,
        '{{to_MS_Teams}}' : to_MS_Teams,
        '{{password}}' : password,
    }
    for key in replace_values_list:
        message = re.sub(key, replace_values_list[key], message)
    return message

def read_file(filename, range):
    """Чтение самого документа"""
    list_value = []
    range = range.split(':')
    wb = load_workbook(filename)
    name_list_file = wb.get_sheet_names()
    sheet = wb.active
    for cellObj in sheet[range[0]:range[1]]:
        list_value.append([cell.value for cell in cellObj])
    return list_value



def get_text_for_send(name_file):
    """Получение текста из файла"""
    f = open(name_file, encoding='utf-8')
    text = f.read()
    f.close()
    return text