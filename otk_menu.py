
import os
import sys
import time
import msvcrt
import shutil 
import subprocess
from win32gui import SetForegroundWindow, ShowWindow
import json #для сохранения значений переменных по умолчанию в файле в формате json
from gurux_dlms.objects import GXDLMSClock, GXDLMSData, GXDLMSRegister, GXDLMSDisconnectControl
import datetime
from gurux_dlms.objects import GXDLMSDisconnectControl, GXDLMSRegister, GXDLMSProfileGeneric
from datetime import datetime, timedelta
from gurux_dlms.objects import GXDLMSClock
from gurux_dlms.enums import DataType, ObjectType
from gurux_dlms import GXDLMSClient, GXTime, GXDateTime
from gurux_dlms.GXDLMSException import *
from pathlib import Path

from libs.otkLib import *
from otk_opto import getAboutOtkOpto, ExchangeBetweenPrograms, optoRunVarRead, \
    getLocalStatistic, getDefaultValue

from otk_opto_getPW import connectionMeterSetup, getMeterPassDefault, changeNumOfMeters

from otk_printStatisticSUTP import getDefaultValueStatistic

from otk_saveLinkCurFolder import *

from colorama import init

init()

def getAboutOtkMenu():
    version = "07.04.2024 20:57"
    descript = "Главное меню для запуска проверок ПУ и модемов"
    return [version, descript]


def readDefaultValue(value_dict={}):
    keys_list=list(value_dict.keys())
    for key in keys_list:
        globals()[key]=value_dict[key]



def toCheckModuleVersions():
    ret = 1
    module_vers_ok_dict = {"otkLib": "05.04.2024 12:45",
                           "otk_opto": "05.04.2024 12:46"}
    otk_opto_vers = getAboutOtkOpto()
    otkLib_vers = getAboutOtkLib()
    format_seconds = "%d.%m.%Y %H:%M"
    for i in module_vers_ok_dict:
        var1 = f"{i}_vers"
        var1_val = locals()[var1][0]
        var1_seconds = int(datetime.strptime(
            var1_val, format_seconds).timestamp())
        control_seconds = int(datetime.strptime(
            module_vers_ok_dict[i], format_seconds).timestamp())
        if var1_seconds < control_seconds:
            txt1 = f"Версия модуля '{i}' {var1_val} не совместима с программой.\n" + \
                f"Требуется версия не ниже {module_vers_ok_dict[i]}."
            print(f"{bcolors.FAIL}{txt1}{bcolors.ENDC}")
            ret = 0
    return ret


def writeDefaultValue(dict={}):
    global default_value_dict

    if len(dict) == 0:
        dict = default_value_dict
    keys_list = list(dict.keys())
    for key in keys_list:
        if key in globals():
            dict[key] = globals()[key]
    return dict


def meterPasswordSearch(source="otk_opto",workmode="эксплуатация"):
    global meter_pw_default     #пароль подключения к ПУ по умолчанию
    global meter_pw_default_descript  #описание пароля по умолчанию
    global meter_pw_encrypt     #зашифрованный пароль подключения к ПУ
    global default_value_dict

    default_value_dict = optoRunVarRead()

    meter_pw_default=default_value_dict['meter_pw_default']
    meter_pw_default_descript=default_value_dict['meter_pw_default_descript']

    res = ExchangeBetweenPrograms(operation="read", recipient="otk_menu")
    if res[0]=="1":
        rec_dict=res[2]
        dt_second =float(rec_dict.get("dateTime"))
        source_1 = rec_dict.get("source")
        operation = rec_dict.get("operation")
        pc_second=toformatNow()[3]
        if pc_second-dt_second < 5 and source_1 == source and \
                operation == "changePassword":
            meter_pass_default=getMeterPassDefault()
            res = readGonfigValue("meter_pass.json", [], meter_pass_default)
            if res[0] != "1":
                    return ["4", "Не удалось прочитать данные "
                            "о пароле из файла", "", ""]
            meter_pw_default_dict = res[2]
            meter_pw_default_old=meter_pw_default
            meter_pw_default_descript_old = meter_pw_default_descript
            keys_list = list(meter_pw_default_dict.keys())
            for key in keys_list:
                meter_pw_default = meter_pw_default_dict[key]
                meter_pw_default_descript = key
                res=cryptStringSec("зашифровать",meter_pw_default)
                meter_pw_encrypt=res[2]
                default_value_dict = writeDefaultValue(
                    default_value_dict)
                saveConfigValue('opto_run.json', default_value_dict)
                _, ans2, file_name = getUserFilePath('otk_check_opto.py',
                    workmode=workmode)
                if file_name == "":
                    return ["4", f"Ошибка в ПП getUserFilePath(): {ans2}"]
                code_exit = os.system(f"python {file_name}")
                res_1 = ExchangeBetweenPrograms(operation="read",
                    recipient="otk_menu_result")
                if res_1[0] == "1":
                    rec_dict_1 = res_1[2]
                    dt_second = float(rec_dict_1.get("dateTime"))
                    source_1 = rec_dict_1.get("source")
                    operation = rec_dict_1.get("operation")
                    result = rec_dict_1.get("result")
                    pc_second = toformatNow()[3]
                    if pc_second-dt_second < 30 and source_1 == "otk_check_opto" and \
                            operation == "checkPassword" and result=="1":
                        os.system("CLS")
                        txt1 = f"{bcolors.OKGREEN}Найден пароль для подключения к ПУ - " \
                            f"'{meter_pw_default_descript}'.{bcolors.ENDC}\n" \
                            f"{bcolors.OKBLUE}Установить его? 0-нет, 1-да.{bcolors.ENDC}"
                        key1 = ["0","1"]
                        oo = questionSpecifiedKey("", txt1, key1)
                        if oo=="0":
                            meter_pw_default=meter_pw_default_old
                            meter_pw_default_descript = meter_pw_default_descript_old
                            default_value_dict = writeDefaultValue(
                                default_value_dict)
                            saveConfigValue('opto_run.json',
                                            default_value_dict)
                            return ["3","Пользователь отказался изменить пароль"]
                        else:
                            return ["1","Подобран новый пароль",meter_pw_default, \
                                    meter_pw_default_descript]
            else:
                txt1 = f"{bcolors.WARNING}Подходящий пароль не найден.{bcolors.ENDC}\n" \
                    f"{bcolors.OKBLUE}Нажмите Enter.{bcolors.ENDC}"
                key1 = ["\r"]
                oo = questionSpecifiedKey("", txt1, key1,"",1)
                meter_pw_default=meter_pw_default_old
                meter_pw_default_descript = meter_pw_default_descript_old
                default_value_dict = writeDefaultValue(default_value_dict)
                saveConfigValue('opto_run.json', default_value_dict)
                return ["2","Подходящего пароля не нашли в списке"]
        else:
            return ["0","Не найдена запись с заданным ключем"]



def displayStatisticSUTPInWindow(interpreter_mode="py",workmode="эксплуатация"):

    title_search = "Статистика по проверке ПУ"
    res = searchTitleWindow(title_search)
    if res[0] == "1":
        hwnd = res[2]
        actionsSelectedtWindow([], hwnd,
            "показать+активировать")
        return ["1", "Окно со статистикой отобразили."]
    file_name='otk_printStatisticSUTP.bat'
    if interpreter_mode=="exe":
        file_name='otk_printStatisticSUTP_exe.bat'
    _, ans2, file_path = getUserFilePath(file_name,
        workmode=workmode)
    if file_path == "":
        return ["0", f"Ошибка в ПП getUserFilePath(): {ans2}"]
    txt1 = f"{bcolors.OKGREEN}Пожалуйста подождите, идет загрузка " \
        f"программы для отображения статистики проверки ПУ...{bcolors.ENDC}\n"
    print(txt1)
    subprocess.Popen(f"start {file_path}", shell=True)
    return ["1", "Окно со статистикой отобразили."]


def mainMenuOLD():

    global meter_pw_default     # пароль доступа к ПУ по умолчанию
    global meter_pw_default_descript    #описание пароля доступа к ПУ
    global com_opto                 # COM-порт оптопорта
    global com_rs485                #COM-порт для RS-485
    global com_config_opto      # COM-порт, к которому подключен оптопорт
    global com_config_rs485     # COM-порт, к которому подключен RS-485
    global default_value_dict       # список со зн-ями по умолчанию
    global workmode                 #метка режима работы программы "тест" - режим теста1, 

    run_mode_dict={'interpreterMode':"py"}
    res = readGonfigValue("run_mode.json",[],run_mode_dict)
    if res[0] != "1":
        print(f"\n{bcolors.WARNING}При формировании перменных "
              f"конфигурации программы возникла ошибка.{bcolors.ENDC}")
        txt1 = f"{bcolors.OKBLUE}Нажмите Ввод.{bcolors.ENDC}"
        questionSpecifiedKey("", txt1, ["\r", "", 1])
        sys.exit()
    run_mode_dict = res[2]
    interpreter_mode=run_mode_dict["interpreterMode"]
    
    cicl=True
    while cicl:

        default_value_dict = optoRunVarRead()
        readDefaultValue(default_value_dict)

        workmode=default_value_dict["workmode"]

        res=updateFilesFromList("auto_upd_file.json", True, 
            workmode)
        
        txt_pass=f"{bcolors.FAIL}Пароль отсутствует в списке.{bcolors.ENDC}"
        pass_set=False
        if meter_pw_default!="":
            res=meterPasswordInfo(password=meter_pw_default,
                err_msg_set="off")
            txt_pass = f"{bcolors.FAIL}Пароль не указан.{bcolors.ENDC}"
            if res[0]=="1":
                descript_pass = res[3]
                txt_pass=f"{bcolors.OKGREEN}Установлен пароль по умолчанию с именем " \
                    f"'{descript_pass}'.{bcolors.ENDC}\n"
                pass_set=True
            
        print()
        os.system("CLS")
        
        prog_last_upd_num=0
        prog_actual_upd_num=0
        prog_actual_upd_moment = ""
        prog_last_upd_moment = ""
        res=getVersUpdProgrFiles(workmode=workmode)
        if res[0]=="0":
            print(f"{bcolors.WARNING}Не удалось получить информацию о новых обновлениях.{bcolors.ENDC}")

        prog_last_upd_num=res[2]
        prog_actual_upd_num=res[3]
        prog_last_upd_moment=res[5]
        prog_actual_upd_moment=res[6]

        
        print(txt_pass)

        com_opto_on=False
        res=checkComPort(com_opto=com_opto,print_msg="")
        if res[0]=="1":
            com_opto_on=True

        txt1="Выберите пункт меню"
        list_txt=["Настройка соединения с ПУ","Запуск полного теста ПУ", \
            "Показать статистику по проверке ПУ в отдельном окне", 
            "Показать версию программы проверки ПУ"]
        list_id=["настройка соединения","полный тест","статистика","версия"]
        if not com_opto_on:
            list_txt=["Настройка соединения с ПУ", "Показать статистику по проверке ПУ в отдельном окне",
            "Показать версию программы проверки ПУ"]
            list_id=["настройка соединения","статистика","версия"]
        if prog_last_upd_num<prog_actual_upd_num:
            print (f"{bcolors.WARNING}Последнее установленное обновление № {prog_last_upd_num} от " 
                   f"{prog_last_upd_moment}.{bcolors.ENDC}\n"
                   f"{bcolors.WARNING}Необходимо обновить ПО до версии № {prog_actual_upd_num} "
                   f"от {prog_actual_upd_moment}.{bcolors.ENDC}")
            list_txt=[f"Обновление файлов программы"]
            list_id=["обновление"]
        cur_id=""
        spec_list=["Выход"]
        spec_keys=["/"]
        spec_id=["выход"]           
        oo = questionFromList(bcolors.OKBLUE, txt1, list_txt, list_id, cur_id, \
                    spec_list=spec_list, spec_keys=spec_keys, spec_id=spec_id)
        print()
        os.system("CLS")

        if oo=="выход":
            return ["выход","Плановый выход из программы"]
        
        elif oo == "обновление":
            _, ans2, file_name = getUserFilePath('otk_progr_update.py')
            if file_name == "":
                return ["ошибка", f"Ошибка в ПП getUserFilePath(): {ans2}"]
            txt1=f"{bcolors.OKGREEN}Пожалуйста подождите, идет загрузка программы обновления файлов...{bcolors.ENDC}\n"
            print(txt1)
            code_exit = os.system(f"python {file_name}")
            return ["успешно","Программа обновления файлов была запущена."]

        elif oo=="версия":
            print()
            os.system("CLS")
            print(f"{bcolors.OKGREEN}Версия программы проверки ПУ № " \
                f"{prog_last_upd_num} от {prog_last_upd_moment}." \
                f"{bcolors.ENDC}")
            txt1 = f"\n{bcolors.OKBLUE}Для возврата в меню нажмите Enter.{bcolors.ENDC}"
            spec_keys=["\r"]
            oo=questionSpecifiedKey(
                colortxt="", txt=txt1, specified_keys_in=spec_keys, file_name_mp3="",
                specified_keys_only=1)
            return ["успешно", "Версию ПО отобразили."]


        elif oo == "статистика":
            res = displayStatisticSUTPInWindow(interpreter_mode=interpreter_mode)
            if res[0]=="0":
                print()
                os.system("CLS")
                print (f"{bcolors.WARNING}Возникла ошибка при " \
                    f"отображении статистики из СУТП.{bcolors.ENDC}\n")
                getLocalStatistic(date_filter="", fio_filter="")
                txt1 = f"\n{bcolors.OKBLUE}Для возврата в меню нажмите Enter.{bcolors.ENDC}"
                spec_keys=["\r"]
                oo=questionSpecifiedKey(
                    colortxt="", txt=txt1, specified_keys_in=spec_keys, file_name_mp3="",
                    specified_keys_only=1)
                return ["ошибка", "Возникла ошибка при " \
                    "создании окна со статистикой из СУТП"]
            else:
                return ["успешно", "Окно со статистикой отобразили."]

        elif oo=="полный тест":
            cicl1=True
            while cicl1:

                title_search = "Статистика по проверке ПУ"
                res = searchTitleWindow(title_search)
                if res[0] == "2":
                    txt1 = f"{bcolors.OKBLUE}Открыть окно для отображения " \
                        f"статистики по проверке ПУ? (0-нет, 1-да){bcolors.ENDC}"
                    spec_keys = ["0","1"]
                    oo = questionSpecifiedKey(colortxt="", txt=txt1, 
                        specified_keys_in=spec_keys, file_name_mp3="",
                        specified_keys_only=1)
                    print()
                    if oo=="1":
                        res = displayStatisticSUTPInWindow(interpreter_mode=interpreter_mode)
                        if res[0]=="0":
                            txt1 = f"{bcolors.WARNING}При открытии окна для отображения " \
                                f"статистики о проверке ПУ возникла ошибка.{bcolors.ENDC}" \
                                f"{bcolors.OKBLUE}Нажмите Enter.{bcolors.ENDC}\n"
                            spec_keys = ["\r"]
                            oo = questionSpecifiedKey(colortxt="", txt=txt1,
                                specified_keys_in=spec_keys, file_name_mp3="",
                                specified_keys_only=1)

                _, ans2, file_name = getUserFilePath('otk_opto_getPW.py')
                if file_name == "":
                    return ["ошибка", f"Ошибка в ПП getUserFilePath(): {ans2}"]
                txt1=f"{bcolors.OKGREEN}Пожалуйста подождите, идет загрузка программы для проверки ПУ...{bcolors.ENDC}\n"
                print(txt1)
                code_exit = os.system(f"python {file_name}")
                
                break

            return ["успешно","полный тест ПУ прошли"]

        elif oo == "ПО модема":
            cicl1=True
            while cicl1:
                file_name = 'otk_modem_vers_PO'
                _, ans2, file_path = getUserFilePath(file_name, only_dir="1")
                if file_path == "":
                    return ["ошибка", f"Ошибка в ПП getUserFilePath(): {ans2}"]
                txt1=f"{bcolors.OKGREEN}Пожалуйста подождите, идет загрузка программы чтения версии ПО модуля связи ...{bcolors.ENDC}\n"
                print(txt1)
                if interpreter_mode=="py":
                    file_path = os.path.join(file_path, file_name)
                    file_path = f"{file_path}.py"
                    code_exit = os.system(f"python {file_path}")
                else:
                    file_path = os.path.split(file_path)[0]
                    file_path = os.path.split(file_path)[0]
                    file_path = f"{file_path}\\{file_name}.exe"
                    code_exit = os.system(f"{file_path}")
                res=meterPasswordSearch(source="otk_modem_vers_PO")
                if res[0]=="0":
                    break
                elif res[0]=="1":
                    print("\nСохраняю выбранный пароль.")
                    meter_pw_default=res[2]
                    meter_pw_default_descript = res[3]
            return ["успешно","Чтение версии ПО модема выполнено."]

        elif oo == "настройка соединения":
            cicl1=True
            while cicl1:
                print()
                os.system("CLS")
                txt_pass=f": {bcolors.FAIL}Пароль по умолчанию отсутствует в списке.{bcolors.ENDC}"
                if meter_pw_default!="":
                    res=meterPasswordInfo(password=meter_pw_default,
                        err_msg_set="off")
                    if res[0]=="1":
                        descript = res[3]
                        txt_pass=f"{bcolors.OKGREEN}{descript}{bcolors.ENDC}"
                else:
                    txt_pass=f"{bcolors.FAIL}Пароль по умолчанию не указан.{bcolors.ENDC}"

                comports_list=[com_opto, com_rs485, com_config_opto, 
                                com_config_rs485]

                comports_txt_list=[]
                for com_cur in comports_list:
                    a_txt=f"{bcolors.FAIL}отсутствует COM-порт{bcolors.ENDC}"
                    res=checkComPort(com_opto=com_cur,print_msg="no")
                    if res=="1":
                        a_txt=f"{bcolors.OKGREEN}{com_cur}{bcolors.ENDC}"
                    comports_txt_list.append(a_txt)

                print()
                os.system("CLS")
                txt1 = "Выберите пункт меню"
                list_txt = [f"Выбор пароля высокого уровня по умолчанию для подключения к ПУ: {txt_pass}", 
                    f"Автоматическое подключение оптопорта для " \
                        f"аппаратной проверки ПУ: {comports_txt_list[0]}",
                    f"Автоматическое подключение преобразователя RS-485 для " \
                        f"аппаратной проверки ПУ: {comports_txt_list[1]}",
                    f"Автоматическое подключение оптопорта для проверки " \
                        f"конфигурации ПУ: {comports_txt_list[2]}",
                    f"Автоматическое подключение RS-485 для проверки " \
                        f"конфигурации ПУ: {comports_txt_list[3]}"]
                list_id = ["пароль", "com-порт", "com-порт RS-485", 
                            "com-config-opto", "com-config-RS485"]
                cur_id = ""
                spec_list=["Выход"]
                spec_keys=["/"]
                spec_id=["выход"]           
                oo = questionFromList(bcolors.OKBLUE, txt1, list_txt, list_id, cur_id, \
                            spec_list=spec_list, spec_keys=spec_keys, spec_id=spec_id)

                if oo == "выход":
                    break

                if oo == "пароль":
                    print()
                    os.system("CLS")
                    meter_pw_default_old = meter_pw_default
                    meter_pw_default_descript_old = meter_pw_default_descript
                    meter_pass_default = getMeterPassDefault()
                    res=readGonfigValue("meter_pass.json",[],meter_pass_default)
                    if res[0]!="1":
                            return ["0","Не удалось прочитать данные " \
                                "о пароле по умолчанию из файла","",""]
                    meter_pw_default_dict=res[2]
                    cicl2 = True
                    while cicl2:
                        print()
                        os.system("CLS")
                        txt1 = "Укажите пароль по умолчанию для подключения к ПУ"
                        list_txt = []
                        keys_list = list(meter_pw_default_dict.keys())
                        for key in keys_list:
                            password = meter_pw_default_dict[key]
                            list_txt.append(f"{key}: {password}")
                            list_id.append(password)
                        spec_list = ["ok","отмена"]
                        spec_keys = ["\r","/"]
                        spec_id = spec_list
                        list_id = keys_list
                        oo = questionFromList(bcolors.OKBLUE, txt1, list_txt, list_id, 
                            meter_pw_default_descript, spec_list, spec_keys, spec_id)
                        if oo == "отмена":
                            meter_pw_default=meter_pw_default_old
                            meter_pw_default_descript = meter_pw_default_descript_old
                            break
                        elif oo=="ok":
                            default_value_dict = writeDefaultValue(
                                default_value_dict)
                            saveConfigValue('opto_run.json',
                                            default_value_dict)
                            break
                        else:
                            meter_pw_default_descript = oo
                            meter_pw_default = meter_pw_default_dict[oo]

                
                elif oo in ["com-порт", "com-порт RS-485", 
                            "com-config-opto", "com-config-RS485"]:
                    a_com_dic = {"com-порт": ["оптопорт", "com_opto",
                        "Настройка оптопорта для аппаратной проверки ПУ."],
                        "com-порт RS-485": ["преобразователь RS-485", 
                                            "com_rs485",
                        "Настройка RS-485 для аппаратной проверки ПУ."],
                        "com-config-opto": ["оптопорт", "com_config_opto",
                        "Настройка оптопорта для проверки конфигурации ПУ."],
                        "com-config-RS485": ["преобразователь RS-485", 
                                              "com_config_rs485",
                        "Настройка RS-485 для проверки конфигурации ПУ."]}
                    res=getAutoCOMPort(a_com_dic[oo][0],"0",a_com_dic[oo][2])
                    a = a_com_dic[oo][1]
                    if res[0]=="1":
                        globals()[a_com_dic[oo][1]]= res[2]
                        default_value_dict = writeDefaultValue(default_value_dict)
                        saveConfigValue('opto_run.json', default_value_dict)



def mainMenu():

    global meter_pw_default     # пароль доступа к ПУ по умолчанию
    global meter_pw_default_descript    #описание пароля доступа к ПУ
    global number_of_meters     #количество подключаемых ПУ
    global com_config_opto      # COM-порт, к которому подключен оптопорт
    global com_config_rs485     # COM-порт, к которому подключен RS-485
    global multi_com_opto_dic   #словарь со списками используемых COM-портов
    global multi_com_rs485_dic   #словарь со списками используемых COM-портов
    global multi_com_config_opto_dic   #словарь со списками используемых COM-портов
    global multi_com_config_rs485_dic   #словарь со списками используемых COM-портов
    global com_config_eqv_com   #метка использования для проверки конфигурации ПУ
    global default_value_dict       # список со зн-ями по умолчанию
    global workmode                 #метка режима работы программы "тест" - режим теста1, 
 
    
    run_mode_dict={'interpreterMode':"py"}
    res = readGonfigValue("run_mode.json",[],run_mode_dict)
    if res[0] != "1":
        print(f"\n{bcolors.WARNING}При формировании перменных "
              f"конфигурации программы возникла ошибка.{bcolors.ENDC}")
        txt1 = f"{bcolors.OKBLUE}Нажмите Ввод.{bcolors.ENDC}"
        questionSpecifiedKey("", txt1, ["\r", "", 1])
        sys.exit()
    run_mode_dict = res[2]
    interpreter_mode=run_mode_dict["interpreterMode"]
    
    window_title = GetWindowText(GetForegroundWindow())

    cicl=True
    while cicl:

        default_value_dict = optoRunVarRead()
        readDefaultValue(default_value_dict)

        workmode=default_value_dict["workmode"]
        number_of_meters=default_value_dict["number_of_meters"]
        if number_of_meters=="" or number_of_meters<1:
            number_of_meters=1
            default_value_dict = writeDefaultValue(default_value_dict)
            saveConfigValue('opto_run.json', default_value_dict)

        multi_com_opto_dic=default_value_dict["multi_com_opto_dic"]
        multi_com_rs485_dic=default_value_dict["multi_com_rs485_dic"]
        multi_com_config_opto_dic=default_value_dict["multi_com_config_opto_dic"]
        multi_com_config_rs485_dic=default_value_dict["multi_com_config_rs485_dic"]

        
        res=updateFilesFromList("auto_upd_file.json", True, 
            workmode)
        
        txt_pass=f"{bcolors.FAIL}Пароль отсутствует в списке.{bcolors.ENDC}"
        pass_set=False
        if meter_pw_default!="":
            res=meterPasswordInfo(password=meter_pw_default,
                err_msg_set="off")
            txt_pass = f"{bcolors.FAIL}Пароль не указан.{bcolors.ENDC}"
            if res[0]=="1":
                descript_pass = res[3]
                txt_pass=f"{bcolors.OKGREEN}Установлен пароль по умолчанию с именем " \
                    f"'{descript_pass}'.{bcolors.ENDC}\n"
                pass_set=True
            
        print()
        os.system("CLS")
        
        prog_last_upd_num=0
        prog_actual_upd_num=0
        prog_actual_upd_moment = ""
        prog_last_upd_moment = ""
        res=getVersUpdProgrFiles(workmode=workmode)
        if res[0]=="0":
            print(f"{bcolors.WARNING}Не удалось получить информацию о новых обновлениях.{bcolors.ENDC}")

        prog_last_upd_num=res[2]
        prog_actual_upd_num=res[3]
        prog_last_upd_moment=res[5]
        prog_actual_upd_moment=res[6]

        print(txt_pass)


        txt1="Выберите пункт меню"
        list_txt=[f"Изменить количество подключаемых ПУ: {number_of_meters}",
            "Запуск полного теста ПУ", "Настройка соединения с ПУ",
            "Показать статистику по проверке ПУ в отдельном окне", 
            "Показать версию программы проверки ПУ"]
        list_id=["изменить количество ПУ", "полный тест", "настройка мульти соединения",
             "статистика","версия"]

                
        if prog_last_upd_num<prog_actual_upd_num:
            print (f"{bcolors.WARNING}Последнее установленное обновление № {prog_last_upd_num} от " 
                   f"{prog_last_upd_moment}.{bcolors.ENDC}\n"
                   f"{bcolors.WARNING}Необходимо обновить ПО до версии № {prog_actual_upd_num} "
                   f"от {prog_actual_upd_moment}.{bcolors.ENDC}")
            list_txt=[f"Обновление файлов программы"]
            list_id=["обновление"]
        
        cur_id=""
        spec_list=["Выход"]
        spec_keys=["/"]
        spec_id=["выход"]           
        oo = questionFromList(bcolors.OKBLUE, txt1, list_txt, list_id, cur_id, \
                    spec_list=spec_list, spec_keys=spec_keys, spec_id=spec_id)
        print()
        os.system("CLS")

        if oo=="выход":
            return ["выход","Плановый выход из программы"]
        
    
        elif oo == "изменить количество ПУ":
            changeNumOfMeters()

        elif oo == "обновление":
            _, ans2, a_file_path = getUserFilePath('otk_progr_update.bat')
            if a_file_path == "":
                return ["ошибка", f"Ошибка в ПП getUserFilePath(): {ans2}"]
            
            txt1="Пожалуйста подождите, идет загрузка " \
                "программы обновления файлов...\n"
            printGREEN(txt1)
            subprocess.Popen(f"start {a_file_path}", shell=True)

            sys.exit()
            

        elif oo=="версия":
            print()
            os.system("CLS")
            print(f"{bcolors.OKGREEN}Версия программы проверки ПУ № " \
                f"{prog_last_upd_num} от {prog_last_upd_moment}." \
                f"{bcolors.ENDC}")
            txt1 = f"\n{bcolors.OKBLUE}Для возврата в меню нажмите Enter.{bcolors.ENDC}"
            spec_keys=["\r"]
            oo=questionSpecifiedKey(
                colortxt="", txt=txt1, specified_keys_in=spec_keys, file_name_mp3="",
                specified_keys_only=1)
            return ["успешно", "Версию ПО отобразили."]

        elif oo == "статистика":
            res = displayStatisticSUTPInWindow(interpreter_mode=interpreter_mode)
            if res[0]=="0":
                print()
                os.system("CLS")
                print (f"{bcolors.WARNING}Возникла ошибка при " \
                    f"отображении статистики из СУТП.{bcolors.ENDC}\n")
                getLocalStatistic(date_filter="", fio_filter="")
                txt1 = f"\n{bcolors.OKBLUE}Для возврата в меню нажмите Enter.{bcolors.ENDC}"
                spec_keys=["\r"]
                oo=questionSpecifiedKey(
                    colortxt="", txt=txt1, specified_keys_in=spec_keys, file_name_mp3="",
                    specified_keys_only=1)
                return ["ошибка", "Возникла ошибка при " \
                    "создании окна со статистикой из СУТП"]
            else:
                return ["успешно", "Окно со статистикой отобразили."]

        elif oo=="полный тест":
            cicl1=True
            while cicl1:

                title_search = "Статистика по проверке ПУ"
                res = searchTitleWindow(title_search)
                if res[0] == "2":
                    txt1 = f"{bcolors.OKBLUE}Открыть окно для отображения " \
                        f"статистики по проверке ПУ? (0-нет, 1-да){bcolors.ENDC}"
                    spec_keys = ["0","1"]
                    oo = questionSpecifiedKey(colortxt="", txt=txt1, 
                        specified_keys_in=spec_keys, file_name_mp3="",
                        specified_keys_only=1)
                    print()
                    if oo=="1":
                        res = displayStatisticSUTPInWindow(interpreter_mode=interpreter_mode)
                        if res[0]=="0":
                            txt1 = f"{bcolors.WARNING}При открытии окна для отображения " \
                                f"статистики о проверке ПУ возникла ошибка.{bcolors.ENDC}" \
                                f"{bcolors.OKBLUE}Нажмите Enter.{bcolors.ENDC}\n"
                            spec_keys = ["\r"]
                            oo = questionSpecifiedKey(colortxt="", txt=txt1,
                                specified_keys_in=spec_keys, file_name_mp3="",
                                specified_keys_only=1)

                _, ans2, file_name = getUserFilePath('otk_opto_getPW.py')
                if file_name == "":
                    return ["ошибка", f"Ошибка в ПП getUserFilePath(): {ans2}"]
                txt1=f"{bcolors.OKGREEN}Пожалуйста подождите, идет загрузка программы для проверки ПУ...{bcolors.ENDC}\n"
                print(txt1)
                code_exit = os.system(f"python {file_name}")
                
                break

            return ["успешно","полный тест ПУ прошли"]

        elif oo == "ПО модема":
            cicl1=True
            while cicl1:
                file_name = 'otk_modem_vers_PO'
                _, ans2, file_path = getUserFilePath(file_name, only_dir="1")
                if file_path == "":
                    return ["ошибка", f"Ошибка в ПП getUserFilePath(): {ans2}"]
                txt1=f"{bcolors.OKGREEN}Пожалуйста подождите, идет загрузка программы чтения версии ПО модуля связи ...{bcolors.ENDC}\n"
                print(txt1)
                if interpreter_mode=="py":
                    file_path = os.path.join(file_path, file_name)
                    file_path = f"{file_path}.py"
                    code_exit = os.system(f"python {file_path}")
                else:
                    file_path = os.path.split(file_path)[0]
                    file_path = os.path.split(file_path)[0]
                    file_path = f"{file_path}\\{file_name}.exe"
                    code_exit = os.system(f"{file_path}")
                res=meterPasswordSearch(source="otk_modem_vers_PO")
                if res[0]=="0":
                    break
                elif res[0]=="1":
                    print("\nСохраняю выбранный пароль.")
                    meter_pw_default=res[2]
                    meter_pw_default_descript = res[3]
            return ["успешно","Чтение версии ПО модема выполнено."]

        elif oo == "настройка мульти соединения":
            connectionMeterSetup()

                            


if  __name__ ==  '__main__' : 

    title_new="Аппаратная проверка ПУ"
    res = searchTitleWindow(title_new)
    if res[0] == "1":
        hwnd=res[2]
        res=actionsSelectedtWindow([], hwnd, "закрыть")
        
            
    
    res = replaceTitleWindow("", title_new)

    res = toCheckModuleVersions()
    if res != 1:
        txt1 = "Нажмите любую клавишу"
        questionOneKey(bcolors.OKBLUE, txt1)
        sys.exit()
 
    os.system("CLS")

    saveLinkCurFolder()

    statistic_default_value_dict=getDefaultValueStatistic()
    res = readGonfigValue("print_statistic_run.json",[],
        statistic_default_value_dict)
    if res[0] != "1":
        print(f"\n{bcolors.WARNING}При формировании перменных "
              f"конфигурации программы возникла ошибка.{bcolors.ENDC}")
        txt1 = f"{bcolors.OKBLUE}Нажмите Ввод.{bcolors.ENDC}"
        questionSpecifiedKey("", txt1, ["\r", "", 1])
        sys.exit()
    print_statistic_config_value_dict = res[2]
    print_statistic_config_value_dict["programUser"]="инженер"
    res = saveConfigValue("print_statistic_run.json",
        print_statistic_config_value_dict)
    if res[0]=="0":
        print(f"\n{bcolors.WARNING}При записи изменений в " \
                f"конфигурационный файл возникла ошибка.{bcolors.ENDC}")
        txt1=f"{bcolors.OKBLUE}Нажмите Ввод.{bcolors.ENDC}"
        questionSpecifiedKey("",txt1,["\r","",1])
        sys.exit()

    while True:
        res=mainMenu()
        if res[0]=="ошибка" or res[0]=="выход":
            exit()
