from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter import messagebox as mb
from openpyxl import load_workbook
import os
import shutil
import tkinter as tk
import tkinter.ttk as ttk
import re
try:
    import configparser
except ImportError:
    import ConfigParser as configparser
import sys
import smtplib


#======================== GLOBAL VARIABLE ==============
NAME_TEXT_FILE = 'text_for_send.txt' 
#=======================================================

#=============== Для компилирования в исполняемый файл .exe пришлось переместить в главный файл 
def send_mail(HOST,PORT,email, password, message, to_email) -> bool:
    """Отправка сообщения"""
    try:
        server = smtplib.SMTP(HOST, PORT)
        server.starttls()
        server.login(email, password)

        server.sendmail(email, to_email, message.encode('utf-8'))
        server.quit()
        
        return True
    except Exception as e:
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
        '{{surname}}' : str(surname),
        '{{name}}' : str(name),
        '{{patronymic}}' : str(patronymic),
        '{{to_the_DO_system}}' : str(to_the_DO_system),
        '{{to_the_testing_system}}' : str(to_the_testing_system),
        '{{to_MS_Teams}}' : str(to_MS_Teams),
        '{{password}}' : str(password),
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
#===========================================================================
CHECK_EMAIL = '^[\w]+([^\s]*)+\@{1}[\w]+\.{1}[\w]+$'
CHECK_RANGE = '^([A-Z]{1}\d{1}){1}$'

class Configure:
    def __init__(self, path):
        self.config = configparser.ConfigParser()
        try:
            self.config.add_section("Settings")
        except Exception as e:
            print(e)
        self.path = path
        self.config.read(path)
        self.host = self.config.get("Settings", 'host')
        self.port = self.config.get("Settings", 'port')
        self.email = self.config.get("Settings", 'email')
        self.password = self.config.get("Settings", 'password')


    def get_config(self):
        """
        Returns the config object
        """
        if not os.path.exists(self.path):
            create_config(self.path)

        config = configparser.ConfigParser()
        config.read(self.path)
        return config

    def update_setting(self, section, setting, value):
        """
        Update a setting
        """
        self.config = self.get_config()
        self.config.set(section, setting, value)
        with open(self.path, "w") as config_file:
            self.config.write(config_file)

    def chek_file_configure(self):
        if not os.path.exists(self.path):
            self.createConfig()
            return False
        else:
            try:
                self.config.get("Settings", "host")
                self.config.get("Settings", "port")
                self.config.get("Settings", "email")
                self.config.get("Settings", "password")
            except Exception as e:
                return False
            return True

    def createConfig(self):
        """
        Create a config file
        """
        with open(self.path, "w") as config_file:
            self.config.write(config_file)


class Main_Window:
    def __init__(self, master, conf, path):
        self.conf = conf
        self.path = path
        self.name_dir = str(os.getcwd()) + str('\\file\\')

        self.master = master
        self.frame = tk.Frame(self.master)
        self.frame.pack()

        self.main_menu = Menu()

        self.file_menu = Menu()
        self.file_menu.add_command(label="Настройки почтового клиента", command=self.new_window)
        # self.file_menu.add_command(label="Save")
        # self.file_menu.add_command(label="Open")
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Выход", command=self.close_windows)

        self.main_menu.add_cascade(label="Опции", menu=self.file_menu)
        self.master.config(menu=self.main_menu)

        self.lable_one = Entry(self.frame, width=5)
        self.lable_two = Entry(self.frame, width=5)
        self.set_values = tk.Button(self.frame, text='Установить диапазон и отправить письмо(а)', width=300, command=self.get_lable_values)
        self.lable_one.pack(side=LEFT, padx=5, pady=0)
        self.lable_two.pack(side=LEFT, padx=5, pady=0)
        self.set_values.pack()

        self.text = Text()
        self.text.pack(side=LEFT)
        self.scroll = Scrollbar()

        self.scroll.pack(side=LEFT, fill=Y)
        self.text.config(yscrollcommand=self.scroll.set)

    def check_field_string(self, pattern, string) -> bool:
        string = re.match(pattern, string)
        if string:
            return True
        else:
            return False


    def get_lable_values(self):

        val_one = self.lable_one.get()
        val_two = self.lable_two.get()
        check_one = False # Такое гавнище конечно но мне лень писать !!!!
        check_two = False
        if self.check_field_string(CHECK_RANGE, val_one) == False:
            mb.showerror(
                "Ошибка",
                f"Установленный диапозон {val_one} не верный !"
            )
        else:
            check_one = True

        if self.check_field_string(CHECK_RANGE, val_two) == False:
            mb.showerror(
                "Ошибка",
                f"Установленный диапозон {val_two} не верный !"
            )
        else: 
            check_two = True
        if check_one and check_two:
            self.filename = askopenfilename()
            list_value = read_file(self.filename, str(val_one) + ':' + str(val_two))
            text = get_text_for_send(NAME_TEXT_FILE)
            # try:
            path = "settings.ini"
            conf = Configure(path)
            answeare = conf.chek_file_configure()

            config = configparser.ConfigParser()
            config.read(path)
            host = config.get("Settings", 'host')
            port = config.get("Settings", 'port')
            email = config.get("Settings", 'email')
            password = config.get("Settings", 'password')

            if not answeare:
                mb.showerror(
                    "Ошибка",
                    f"Ошибка файла конфигурации \nпроверьте ввод настроик почтового клиента и попробуйте снова !\n\n{e}"
                )

            for index, value in enumerate(list_value):
                message = set_message(list_value[index], text)
                parametrs_in_list = list_value[index]
                if self.check_field_string(CHECK_EMAIL, parametrs_in_list[6]) == True:
                    send_mail(host, port, email, password, message, parametrs_in_list[6])
                    self.text.insert(1.0, f"Success full:  Письмо на адрес {parametrs_in_list[6]} было отправленно ! \n")
                else:
                    self.text.insert( f"{parametrs_in_list[6]} , Данная почта не соответвстует типу, например example@test.ru !\n")
            # except Exception as e:
            #     mb.showerror(
            #         "Ошибка",
            #         f"Ошибка файла\nпроверьте файл на ошибки и попробуйте снова !\n\n{e}"
            #     )
    def close_windows(self):
        self.master.destroy()

    def new_window(self):
        self.newWindow = tk.Toplevel(self.master)
        self.newWindow.geometry('300x240')
        Set_Settings(self.newWindow, self.path)
        


class Set_Settings:
    def __init__(self, master, path):
        self.conf = Configure(path)
        self.master = master
        self.group_1 = LabelFrame(self.master, padx=15, pady=10, text="Персональная информация")
        self.group_1.pack(padx=10, pady=5)

        Label(self.group_1, text="Хост").grid(row=0)
        Label(self.group_1, text="Порт").grid(row=1)
        Label(self.group_1, text="E-mail").grid(row=2)
        Label(self.group_1, text="Пароль").grid(row=3)
        self.lable_host = Entry(self.group_1)
        self.lable_host.insert(0, self.conf.host)
        self.lable_host.grid(row=0, column=1, sticky=W)
        self.lable_port = Entry(self.group_1)
        self.lable_port.insert(0, self.conf.port)
        self.lable_port.grid(row=1, column=1, sticky=W)
        self.lable_email = Entry(self.group_1)
        self.lable_email.insert(0, self.conf.email)
        self.lable_email.grid(row=2, column=1, sticky=W)
        self.lable_password = Entry(self.group_1)
        self.lable_password.insert(0, self.conf.password)
        self.lable_password.grid(row=3, column=1, sticky=W)

        self.btn_submit = Button(self.master, text="Сохранить", command=self.set_smtp_parametrs)
        self.btn_submit.pack(padx=10, pady=10, side=RIGHT)

    def set_smtp_parametrs(self):
        host = self.lable_host.get()
        port = self.lable_port.get()
        email = self.lable_email.get()
        password = self.lable_password.get()

        if host:
            chek_host = True
        else:
            chek_host = False
        if port:
            chek_port = True
        else:
            chek_port = False
        if email:
            chek_email = True
        else:
            chek_email = False
        if password:
            chek_password = True
        else:
            chek_password = False

        try:
            if chek_host:
                self.conf.update_setting("Settings", "host", host)

            if chek_port:
                self.conf.update_setting("Settings", "port", port)

            if chek_email:
                self.conf.update_setting("Settings", "email", email)

            if chek_password:
                self.conf.update_setting("Settings", "password", password)

            self.close_windows()
        except:
            mb.showerror(
                "Ошибка",
                "Должно быть введены значения"
            )

    def close_windows(self):
        self.master.destroy()


def main():
    root = tk.Tk()
    path = "settings.ini"
    conf = Configure(path)
    answeare = conf.chek_file_configure()
    if not answeare:
        app = Set_Settings(root, path)
        root.geometry('345x350')
        root.mainloop()
    else:
        app = Main_Window(root, conf, path)
        root.geometry('345x350')
        root.mainloop()


if __name__ == '__main__':
    main()
