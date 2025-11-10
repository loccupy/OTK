# программа проверки ПУ

def getAboutOtkOpto():
    version = "04.05.2024 07:02"
    descript = "Программа для проверки ПУ"
    return [version, descript]


import os
import sys
import time
import datetime
import msvcrt
import shutil 
import webbrowser
import segno #pip install segno для печати qr-кода
import pyperclip #pip install pyperclip для записи данных в буфер Windows
import docxtpl #pip install docxtpl для обработки *.docx-файлов
import win32print #pip install pywin32 для печати паспорта напрямую в принтер
import win32api #pip install pywin32 для печати паспорта напрямую в принтер
import json #для сохранения значений переменных по умолчанию в файле в формате json
from transliterate import translit      #pip install transliterate для транслитерации текста, записываемое в формате json

from gtts import gTTS  # pip install gTTS для синтеза речи с пом.Google через интернет
from playsound import playsound  # pip install playsound==1.2.2 для воспроизведения звукового файла
from tqdm import tqdm   #pip install tqdm   для отображения progress bar
from prettytable import PrettyTable

from gurux_dlms.objects import GXDLMSClock, GXDLMSData, GXDLMSRegister, GXDLMSDisconnectControl
from gurux_dlms.objects import GXDLMSDisconnectControl, GXDLMSRegister, GXDLMSProfileGeneric
from datetime import datetime, timedelta
from gurux_dlms.objects import GXDLMSClock
from gurux_dlms.enums import DataType, ObjectType
from gurux_dlms import GXDLMSClient, GXTime, GXDateTime, GXTimeZone
from gurux_dlms.GXDLMSException import *
from docxtpl import DocxTemplate, InlineImage, RichText
from pathlib import Path


from libs.sutpLib import getAboutSutpLib, savetToSUTP2, getNameEmployee, findNameMeterStatus, \
    getInfoAboutDevice,request_sutp,getDeviceHistory,getDeviceRepayHistory, preChecksToGhangeStatusMeter, \
    getInfoAboutDockedMC, getMeterAllSN
from libs.otkLib import *

init()


def toCheckModuleVersions():
    ret = 1
    module_vers_ok_dict = {
        "sutpLib": "13.04.2024 07:09", "otkLib": "14.04.2024 18:54"}
    sutpLib_vers = getAboutSutpLib()
    otkLib_vers = getAboutOtkLib()
    format_seconds="%d.%m.%Y %H:%M"
    for i in module_vers_ok_dict:
        var1 = f"{i}_vers"
        var1_val = locals()[var1][0]
        var1_seconds = int(datetime.strptime(var1_val, format_seconds).timestamp())
        control_seconds = int(datetime.strptime(module_vers_ok_dict[i], format_seconds).timestamp())
        if var1_seconds < control_seconds:
            txt1 = f"Версия модуля '{i}' {var1_val} не совместима с программой.\n" + \
                f"Требуется версия не ниже {module_vers_ok_dict[i]}."
            print(f"{bcolors.FAIL}{txt1}{bcolors.ENDC}")
            ret = 0
    return ret



def disconnectControlStatus():
    disconnect_control = GXDLMSDisconnectControl("0.0.96.3.10.255")
    a1=reader.read(disconnect_control, 1)
    a2=reader.read(disconnect_control, 2)
    a3 = reader.read(disconnect_control, 3)
    a4=reader.read(disconnect_control, 4)
    print ("физ.состояние реле (1): "+str(a1))
    print("программное состояние реле (2): "+str(a2))
    print("режим работы реле (3): "+str(a3))
    print(" (4): "+str(a4))


def testSwitchDisable():
    ret=3
    disconnect_control = GXDLMSDisconnectControl("0.0.96.3.10.255")
    cicl=True
    while cicl:
        try:
            disconnect_output_state = reader.read(disconnect_control, 2) #физическое состояние реле: True-замкнуто, False-разомкнуто
            disconnect_control_state = reader.read(disconnect_control, 3) #программное состояние реле: 0-разомкнуто, 1-замкнуто, 2-готово замкнуться при нажатии кн.МЕНЮ
            break
        except Exception as e:
            oo = communicationTimoutError("Определяли состояние реле при левом положении переключателя блокировки: ", e.args[0])
            if oo=="0" or oo=="-1":
                return 4
    if disconnect_control_state == 1:
        try:
            print(f"    подаю команду на размыкание реле...")
            reader.relay_disconnect()
            cicl = True
            while cicl:
                try:
                    disconnect_output_state = reader.read(disconnect_control, 2)
                    break
                except Exception as e:
                    oo = communicationTimoutError("Определяли состояние реле после попытки размокнуть контакты реле: ", e.args[0])
                    if oo=="0" or oo=="-1":
                        return 4
            if disconnect_output_state==False:
                print(f"{bcolors.FAIL}    Переключатель не сработал. Реле разомкнулось.{bcolors.ENDC}")
                ret=3
                oo=questionRetry()
                if oo==True:
                    ret=2
                return ret            
        except GXDLMSException:
            cicl = True
            while cicl:
                try:
                    disconnect_output_state = reader.read(disconnect_control, 2) #физическое состояние реле
                    disconnect_control_state = reader.read(disconnect_control, 3) #программное состояние реле
                    break
                except Exception as e:
                    oo = communicationTimoutError("Определяли состояние реле после попытки размокнуть контакты реле: ", e.args[0])
                    if oo=="0" or oo=="-1":
                        return 4
            if disconnect_output_state==True and disconnect_control_state == 1:
                print(f"    ok: реле осталось в замкнутом состоянии...")
                ret=1
                return ret
            else:
                print(f"{bcolors.FAIL}Проверка работы блокировки реле (блок \"testSwitchDisable\"): программная ошибка")
                ret=2
                return ret
        except Exception as e:
            oo = communicationTimoutError("Подавали команду на размыкание контактов реле: ", e.args[0])
            if oo=="0" or oo=="-1":
                return 4
    elif disconnect_control_state==0 or disconnect_control_state == 2:
        try:
            print(f"    подаю команду на замыкание реле...")
            reader.relay_reconnect()
            cicl = True
            while cicl:
                try:
                    disconnect_output_state = reader.read(disconnect_control, 2) #физическое состояние реле: True-замкнуто, False-разомкнуто
                    disconnect_control_state = reader.read(disconnect_control, 3) #программное состояние реле: 0-разомкнуто, 1-замкнуто, 2-готово замкнуться при нажатии кн.МЕНЮ
                    disconnect_control_mode = reader.read(disconnect_control, 4) #режим работы реле
                    break
                except Exception as e:
                    oo = communicationTimoutError(
                        "Определяли состояние реле после попытки замкнуть контакты реле: ", e.args[0])
                    if oo == "0" or oo == "-1":
                        return 4
            if disconnect_control_state == 2:
                questionOneKeyPause(
                    bcolors.OKBLUE, "    Зажмите кнопку МЕНЮ на 5 секунд", "    Нажмите любую клавишу", 5)
                cicl = True
                while cicl:
                    try:
                        disconnect_output_state = reader.read(disconnect_control, 2) #физическое состояние реле
                        disconnect_control_state = reader.read(disconnect_control, 3) #программное состояние реле
                        break
                    except Exception as e:
                        oo = communicationTimoutError(
                            "Определяли состояние реле после попытки замкнуть контакты реле с помощью кн.МЕНЮ: ", e.args[0])
                        if oo == "0" or oo == "-1":
                            return 4
                if disconnect_output_state and disconnect_control_state == 1:
                    print(f"{bcolors.FAIL}    Переключатель не сработал. Реле замкнулось.{bcolors.ENDC}")
                    ret=3
                    oo=questionRetry()
                    if oo==True:
                        ret=2
                    return ret
                elif disconnect_output_state==False and disconnect_control_state == 2:
                    print(f"\n    ok: реле осталось в готовности замкнуться...")
                    ret=1
                    return ret
                else:
                    txt1_1="разомкнуто"
                    txt1_2="разомкнуто"
                    if disconnect_output_state:
                        txt1_1="замкнуто"
                    if disconnect_control_state==1:
                        txt1_2="замкнуто"
                    elif disconnect_control_state==2:
                        txt1_2="готово замкнуться при зажатии кн.МЕНЮ на 5 сек."
                    print(f"{bcolors.FAIL}Проверка работы блокировки реле при левом положении "
                     f"переключателя (блок \"testSwitchDisable\").\n Физическое состояние реле: {txt1_1}\n"
                     f"Программное состояние реле: {txt1_2}")
                    ret=2
                    return ret 
            if disconnect_output_state and disconnect_control_state==1:
                print(f"{bcolors.FAIL}    Переключатель не сработал. Реле замкнулось.{bcolors.ENDC}")
                ret=3
                oo=questionRetry()
                if oo==True:
                    ret=2
                return ret
            else:
                txt1_1 = "разомкнуто"
                txt1_2 = "разомкнуто"
                if disconnect_output_state:
                    txt1_1 = "замкнуто"
                if disconnect_control_state == 1:
                    txt1_2 = "замкнуто"
                elif disconnect_control_state == 2:
                    txt1_2 = "готово замкнуться при зажатии кн.МЕНЮ на 5 сек."
                print(f"{bcolors.FAIL}Проверка работы блокировки реле при левом положении " 
                        f"переключателя (блок \"testSwitchDisable\").\n Физическое состояние реле: {txt1_1}\n" 
                        f"Программное состояние реле: {txt1_2}")
                ret=2
                return ret 
        except GXDLMSException:
            disconnect_output_state = reader.read(disconnect_control, 2) #физическое состояние реле
            if disconnect_output_state==False :
                print(f"    ok: реле осталось в разомкнутом состоянии...")
                ret=1
                return ret
            else:
                txt1_1 = "разомкнуто"
                txt1_2 = "разомкнуто"
                if disconnect_output_state:
                    txt1_1 = "замкнуто"
                if disconnect_control_state == 1:
                    txt1_2 = "замкнуто"
                elif disconnect_control_state == 2:
                    txt1_2 = "готово замкнуться при зажатии кн.МЕНЮ на 5 сек."
                print(f"{bcolors.FAIL}Проверка работы блокировки реле при левом положении " 
                        f"переключателя (блок \"testSwitchDisable\").\n Физическое состояние реле: {txt1_1}\n"
                        f"Программное состояние реле: {txt1_2}")
                ret=2
                return ret
        except Exception as e:
            oo = communicationTimoutError("Подавали команду на замыкание контактов реле: ", e.args[0])
            if oo=="0" or oo=="-1":
                return 4
    

def testSwitchEnable():
    ret=3
    cicl=True
    while cicl:
        disconnect_control = GXDLMSDisconnectControl("0.0.96.3.10.255")
        try:
            disconnect_output_state = reader.read(disconnect_control, 2) #физическое состояние реле: True-замкнуто, False-разомкнуто
            disconnect_control_state = reader.read(disconnect_control, 3) #программное состояние реле: 0-разомкнуто, 1-замкнуто, 2-готово замкнуться при нажатии кн.МЕНЮ
            disconnect_control_mode = reader.read(disconnect_control, 4)  # режим работы реле
            break
        except Exception as e:
            oo = communicationTimoutError("Определяли состояние реле при правом положении переключателя блокировки:", e.args[0])
            if oo=="0" or oo=="-1":
                return 4
    if disconnect_control_state==1:
        cicl=True
        while cicl:
            print(f"    подаю команду на размыкание реле...")
            try:
                reader.relay_disconnect()
                disconnect_output_state = reader.read(disconnect_control, 2) #физическое состояние реле: True-замкнуто, False-разомкнуто
                if disconnect_output_state==False:
                    print(f"    ok: реле разомкнулось...")
                    ret=1
                    return ret           
            except GXDLMSException:
                disconnect_output_state = reader.read(disconnect_control, 2) #физическое состояние реле
                if disconnect_output_state:
                    print(f"{bcolors.FAIL}    Переключатель не сработал - реле осталось замкнуто.{bcolors.ENDC}")
                    ret=3
                    oo=questionRetry()
                    if oo==True:
                        ret=2
                    return ret 
                else:
                    txt1_1 = "разомкнуто"
                    txt1_2 = "разомкнуто"
                    if disconnect_output_state:
                        txt1_1 = "замкнуто"
                    if disconnect_control_state == 1:
                        txt1_2 = "замкнуто"
                    elif disconnect_control_state == 2:
                        txt1_2 = "готово замкнуться при зажатии кн.МЕНЮ на 5 сек."
                    print(f"{bcolors.FAIL}Проверка работы блокировки реле при левом положении "
                            f"переключателя (блок \"testSwitchDisable\").\n Физическое состояние реле: {txt1_1}\n"
                            f"Программное состояние реле: {txt1_2}")
                    ret=2
                    return ret
            except Exception as e:
                oo = communicationTimoutError(
                    "При правом положении переключателя блокировки пытались разомкнуть реле:", e.args[0])
                if oo == "0" or oo == "-1":
                    return 4
    elif disconnect_control_state==0:
        cicl=True
        while cicl:
            try:
                print(f"    подаю команду на замыкание реле...")
                reader.relay_reconnect()
                disconnect_output_state = reader.read(disconnect_control, 2) #физическое состояние реле
                disconnect_control_state = reader.read(disconnect_control, 3) #программное состояние реле
                if disconnect_output_state and disconnect_control_state != 2:
                    print(f"    ok: реле замкнулось...")
                    ret=1
                    return ret
                if disconnect_control_state == 2:
                    break
            except GXDLMSException:
                disconnect_output_state = reader.read(disconnect_control, 2) #физическое состояние реле
                if disconnect_output_state==False:
                    print(f"{bcolors.FAIL}    Переключатель не сработал - реле осталось разомкнуто.{bcolors.ENDC}")
                    ret=3
                    oo=questionRetry()
                    if oo==True:
                        ret=2
                    return ret
                else:
                    txt1_1 = "разомкнуто"
                    txt1_2 = "разомкнуто"
                    if disconnect_output_state:
                        txt1_1 = "замкнуто"
                    if disconnect_control_state == 1:
                        txt1_2 = "замкнуто"
                    elif disconnect_control_state == 2:
                        txt1_2 = "готово замкнуться при зажатии кн.МЕНЮ на 5 сек."
                    print(f"{bcolors.FAIL}Проверка работы блокировки реле при левом положении "
                            f"переключателя (блок \"testSwitchDisable\").\n Физическое состояние реле: {txt1_1}\n"
                            f"Программное состояние реле: {txt1_2}")
                    ret=2
                    return ret
            except Exception as e:
                oo = communicationTimoutError(
                    "При правом положении переключателя блокировки пытались замкнуть реле:", e.args[0])
                if oo == "0" or oo == "-1":
                    return 4
    if disconnect_control_state==2:
        cicl=True
        while cicl:
            questionOneKeyPause(bcolors.OKBLUE,"    Зажмите кнопку МЕНЮ на 5 секунд. В течение этого времени реле должно замкнуться.", "    Нажмите любую клавишу", 5)
            try:
                disconnect_output_state = reader.read(disconnect_control, 2) #физическое состояние реле
            except Exception as e:
                oo = communicationTimoutError(
                    "При правом положении переключателя блокировки определяли фактическое положение реле после нажатия кн.МЕНЮ для замыкания реле:", e.args[0])
                if oo == "0" or oo == "-1":
                    return 3
            if disconnect_output_state:
                print(f"\n    ok: реле замкнулось...")
                ret=1
                return ret 
            else:
                print(f"{bcolors.FAIL}    Реле осталось разомкнуто.{bcolors.ENDC}")
                ret=3
                oo=questionRetry()
                if oo==False:
                    return ret


def testDisplayRele():
    disconnect_control = GXDLMSDisconnectControl("0.0.96.3.10.255")
    disconnect_control_state = reader.read(disconnect_control, 3) #программное состояние реле: 0-разомкнуто, 1-замкнуто, 2-готово замкнуться при нажатии кн.МЕНЮ
    disconnect_output_state = reader.read(disconnect_control, 2) #физическое состояние реле: True-замкнуто, False-разомкнуто
    ret=""
    if disconnect_output_state:
        txt1_1="    Проверьте на ЖКИ наличие символа замкнутого реле. Символ имеется? 0-нет, 1-да"
        oo=questionSpecifiedKey(bcolors.OKBLUE,txt1_1,["0","1"])
        if oo=="0":
            txt1_1="    Вы уверены, что на ЖКИ нет символа замкнутого реле? 0-нет, 1-да"
            oo_ref=questionSpecifiedKey(bcolors.WARNING,txt1_1,["0","1"])
            if oo_ref=="1":
                txt1="Не отображается символ замкнутого реле при замкнутом положении реле."
                print(f"\n    {bcolors.FAIL}Тест не пройден. {txt1}{bcolors.ENDC}")
                ret=txt1
    elif disconnect_control_state==0:
        txt1_1="    Проверьте на ЖКИ наличие символа разомкнутого реле. Символ имеется? 0-нет, 1-да"
        oo=questionSpecifiedKey(bcolors.OKBLUE,txt1_1,["0","1"])
        if oo=="0":
            txt1_1="    Вы уверены, что на ЖКИ нет символа разомкнутого реле? 0-нет, 1-да"
            oo_ref=questionSpecifiedKey(bcolors.WARNING,txt1_1,["0","1"])
            if oo_ref=="1":
                txt1="Не отображается символ разомкнутого реле при отключенном положении реле."
                print(f"\n    {bcolors.FAIL}Тест не пройден. {txt1}{bcolors.ENDC}")
                ret=txt1
    elif disconnect_control_state==2:
        txt1_1="    Проверьте на ЖКИ наличие мигающего символа разомкнутого реле. Символ имеется? 0-нет, 1-да"
        oo=questionSpecifiedKey(bcolors.OKBLUE,txt1_1,["0","1"])
        if oo=="0":
            txt1_1="    Вы уверены, что на ЖКИ не мигает символ разомкнутого реле? 0-нет, 1-да"
            oo_ref=questionSpecifiedKey(bcolors.WARNING,txt1_1,["0","1"])
            if oo_ref=="1":
                txt1="Не мигает символ разомкнутого реле при готовности реле к включению."
                print(f"{bcolors.FAIL}\n    Тест не пройден. {txt1}{bcolors.ENDC}")
                ret=txt1
    return ret



def toSaveResultExtInspection(meter_serial_number_1,meter_tech_number,
    external_condition, employees_name,employee_id, 
    default_filename_ext, default_dirname, work_dirname,
    dirname_sos, filename1, sutp_to_save, data_exchange_sutp, 
    save_mode="0", meter_next_step="0", workmode="эксплуатация", 
    reestr_dic={}, rep_copy_public="0", meter_config_check="3",
    config_send_mail="1", rep_err_send_mail="1",
    no_data_in_SUTP_send_mail="1", rep_remark_txt=""):

    res=checkVarProgrAvailable("sutp_to_save", sutp_to_save, "1")
    if res[0]!="1":
        return ["0", res[1]]

    meter_typegroup=""

    sheet_name = "Product1"
    if meter_serial_number_1!="":
        res = toGetProductInfo2(meter_serial_number_1, sheet_name)
        if res[0]=="1":
            meter_typegroup=res[4]
            meter_phase=res[18]
    res=toformatNow()
    dt=res[0]
    dt_sec=str(res[3])
    employee_ip_adr=getLocalIP()
    reestr_dic["id_rec"]=f"{dt_sec}_{employee_ip_adr}"

    res=getSutpToSaveDescript(sutp_to_save)
    sutp_to_save_descript=res[2]
    sutp_to_save_color=res[3]

    sutp_transm=sutp_to_save_descript

    meter_type=reestr_dic.get("meter_model_sutp", "")
    if meter_type=="":
        meter_type=meter_typegroup

    meter_config_param_filename=reestr_dic.get("meter_config_param_filename", "")
    meter_config_filename=reestr_dic.get("meter_config_filename", "")
    
    if save_mode=="0":
        txt = f"1. Дата и время проведения внешней проверки: {dt}\n"+ \
            f"2. Серийный номер ПУ: {meter_serial_number_1}\n"+ \
            f"3. Технический номер ПУ: {meter_tech_number}\n"+ \
            f"4. Вид ПУ: {meter_typegroup}\n"+ \
            f"5. Замечания к состоянию ПУ:\n{external_condition}\n\n"+ \
            f"Проверка проведена: {employees_name}\n"
        
        a_external_condition=external_condition.replace("\n", ",")
        reestr_dic['reestr_clipboard_err_txt']=a_external_condition

        fileWriter(default_filename_ext, "a", "", txt+ "\n", 
            "Сохранение в отчет результата внешней проверки")
    
    reestr_dic['employees_name']=employees_name
    reestr_dic['meter_serial_number']=meter_serial_number_1
    reestr_dic['meter_tech_number']=meter_tech_number
    reestr_dic['meter_type']=meter_typegroup
    
    
    if external_condition!="":
        pyperclip.copy(external_condition)
        print(f"{bcolors.WARNING}\nCписок замечаний записан в буфер " \
            f"обмена Windows:\n'{external_condition}'{bcolors.ENDC}\n")
    

    while True:
        if data_exchange_sutp!="0":
            if save_mode=="2" and meter_next_step=="1":
                txt1="Т.к. проводилась проверка только конфигурации ПУ, то " \
                    "результаты проверки в СУТП не переданы."
                printWARNING(txt1)

                fileWriter(default_filename_ext, "a", "", f"{txt1}\n", 
                    "Сохранение в отчет информации об обмене с СУТП.")
                
                break
            
            
            if sutp_to_save=="0" and save_mode=="2" and meter_next_step=="0":
                txt1="Выявлены замечания к конфигурации ПУ.\n" \
                    f"Отправить ПУ на ремонт? (0-нет, 1-да):"
                oo=questionSpecifiedKey(bcolors.OKBLUE, txt1, ["0", "1"], "", 1)
                print()
                if oo=="0":
                    txt1="Информация об отрицательном результате проверки " \
                        "конфигурации ПУ в СУТП не передана."
                    printWARNING(txt1)
                    fileWriter(default_filename_ext, "a", "", f"{txt1}\n", 
                        "Сохранение в отчет информации об обмене с СУТП.")
                    
                    break
    

            res="0"
            sutp_transmitted=False
            sutp_transm=sutp_to_save_descript
            txt1=f"{sutp_to_save_color}{sutp_to_save_descript}{bcolors.ENDC}"

            if sutp_to_save[0]=="2" or (sutp_to_save[0]=="3" and 
                    meter_next_step=="0"):
                if employee_id!="0":
                    res = savetToSUTP2(meter_tech_number, employee_id, meter_next_step,21,
                                    external_condition)

                    txt1=f"{bcolors.WARNING}Ошибка при передаче данных на сервер СУТП.{bcolors.ENDC}\n"
                    sutp_transm="Ошибка при передаче данных на сервер СУТП."
                    if res=="1":
                        sutp_transmitted=True
                        sutp_transm="Данные на сервер СУТП переданы."
                        txt1=f"{bcolors.OKGREEN}{sutp_transm}{bcolors.ENDC}"
                    
            print(f"{txt1}")
            fileWriter(default_filename_ext, "a", "", f"{sutp_transm}\n", 
                "Сохранение в отчет информации о передаче в СУТП информации "
                "о результате проверки")
            
            if (sutp_transmitted == False and (sutp_to_save[0]=="2" or \
                    (sutp_to_save[0]=="3" and meter_next_step=="0"))) or \
                    sutp_to_save[0]=="1" :
                print(f"{bcolors.WARNING}Внесите результаты проверки ПУ в СУТП вручную.{bcolors.ENDC}\n"
                    f"{bcolors.OKBLUE}После окончания нажмите Enter.{bcolors.ENDC}")
                webbrowser.open('http://sutp.promenergo.local/section/14', new=0)
                oo=questionSpecifiedKey( "", "", ["\r"], "", 1)
            
            
            if sutp_to_save[0]=="2" or (sutp_to_save[0]=="3" and 
                    meter_next_step=="0"):
                print ("Запрашиваю статус ПУ в СУТП...")
                a_mode="1"
                a_pw_encrypt=""
                res=getInfoAboutDevice(meter_tech_number, workmode, employee_id, 
                                        a_pw_encrypt, a_mode)
                txt1=f"{bcolors.WARNING}Постконтроль статуса ПУ в СУТП: не удалось " \
                    f"получить из СУТП информацию о новом статусе ПУ.{bcolors.ENDC}"
                sutp_transm=f"Постконтроль статуса ПУ в СУТП: не удалось получить " \
                    f"из СУТП информацию о новом статусе ПУ."
                sutp_transmitted=False
                if res[0]!="0":
                    a_status=res[7]
                    a_color=bcolors.OKGREEN
                    a_status_dic={"0": "Дефект", "1": "ОТК пройден"}
                    sutp_transm=f"Статус ПУ в СУТП: {a_status}."
                    sutp_transmitted=True
                    if a_status not in a_status_dic[meter_next_step]:
                        a_color=bcolors.WARNING
                        sutp_transm=f"Постконтроль статуса ПУ в СУТП: ошибка " \
                            f"при изменении статуса ПУ. Текущий статус ПУ в " \
                            f"СУТП: {a_status}."
                        sutp_transmitted=False
                    txt1=f"{a_color}{sutp_transm}{bcolors.ENDC}"

                print(f"{txt1}")

                if sutp_transmitted==False:
                    txt1="Выберите дальнейшее действие:"
                    menu_item_list=["Повторно отправить информацию в СУТП", 
                            "Закончить проверку ПУ"]
                    menu_id_list=["повторить", "закончить"]
                    spec_list=[]
                    spec_keys=[]
                    spec_id_list=[]
                    oo=questionFromList(bcolors.OKBLUE, txt1, menu_item_list, menu_id_list,
                        "", spec_list, spec_keys, spec_id_list, 1, start_list_num=1)
                    print()
                    if oo=="повторить":
                        continue

                fileWriter(default_filename_ext, "a", "", f"{sutp_transm}\n", 
                    "Сохранение в отчет информации о постконтроле состояния ПУ в СУТП "
                    "о результате проверки")

        else:
            txt1=f"{sutp_to_save_color}{sutp_to_save_descript}{bcolors.ENDC}"

            print(txt1)
            fileWriter(default_filename_ext, "a", "", f"{txt1}\n", 
                "Сохранение в отчет информации об обмене с СУТП.")
            
        break

    
    history_serial_number_txt=""
    history_status_txt=""
    history_repair_txt=""

    if data_exchange_sutp=="1":
        res=getMeterAllSN(meter_tech_number, workmode, "0")
        if res[0]=="1":
            history_serial_number_txt=res[3]
            if history_serial_number_txt=="":
                history_serial_number_txt="нет данных"
            history_serial_number_txt=f"\nИстория присвоения серийных номеров ПУ № " \
            f"{meter_tech_number}:\n{history_serial_number_txt}" \
                f"\n---Конец истории присвоения серийных номеров ПУ.\n"


        res=getDeviceHistory(device_number=meter_tech_number,
                             employee_print="1")
        if res[0]=="1":
            history_status_txt=res[2]
            history_status_txt=f"\nИстория движения ПУ № " \
                f"{meter_tech_number}:\n{history_status_txt}" \
                f"---Конец истории движения ПУ.\n"

            
        res=getDeviceRepayHistory(meter_tech_number, workmode=workmode)
        history_repair_txt=""
        if res[0]=="1":
            history_repair_txt=res[2]
            if history_repair_txt=="":
                history_repair_txt="нет данных"
            history_repair_txt=f"\nИстория ремонта ПУ № " \
                f"{meter_tech_number}:\n{history_repair_txt}" \
                f"\n---Конец истории ремонта ПУ.\n"
            
    
    if config_send_mail!="0" and meter_config_check[0]!="0":
        res=sendMailErrConfigMeter(history_status_txt, workmode, meter_type,
            meter_config_param_filename, meter_config_filename)
        if res[0]=="1":
            a_txt="Отправлено уведомление об отрицательном результате " \
                "проверки конфигурации ПУ"
            printGREEN(f"{a_txt}.")

            a_to_mail=res[2]

            if len(a_to_mail.split(","))>1:
                a_txt=a_txt+f" на адреса: {a_to_mail}."

            else:
                a_txt=a_txt+f" на адрес: {a_to_mail}."

            fileWriter(default_filename_ext, "a", "", f"{a_txt}\n", 
                "Сохранение в отчет информации об отправке "
                "уведомления об отрицательном результате проверки"
                "конфигурации ПУ.")
    
    txt1=f"\n---Конец протокола."
    meter_pw_low_encrypt=reestr_dic["meter_pw_low_encrypt"]
    meter_pw_high_encrypt=reestr_dic["meter_pw_high_encrypt"]
    txt1=f"{txt1}\n{meter_pw_low_encrypt}\n" \
        f"{meter_pw_high_encrypt}\n"
    fileWriter(default_filename_ext, "a", "", f"{txt1}\n", 
        "Сохранение в отчет метки конца протокола.")

    
    history_all_txt=history_serial_number_txt + history_status_txt+ \
        history_repair_txt
    
    if history_all_txt!="":
        fileWriter(default_filename_ext, "a", "", history_all_txt,
            "Сохранение в отчет историй о ПУ из СУТП.")
        
   
    if meter_config_check[0]=="3" and save_mode!="0":
        res=readGonfigValue("mass_config.json",[],{}, workmode, "1")
        if res[0]=="1":
            mass_responce=res[2]["mass_responce"]
            if mass_responce=="1":
                res=readGonfigValue("mass_log_line_multi.json",[],{}, workmode, "1")
                if res[0]=="1":
                    log_line_txt="Записи отсутствуют.\n"
                    mass_result_dic=res[2].get(meter_tech_number, {})
                    if len(mass_result_dic)==0:
                        a_err_txt="Отсутствует информация о проверке конфигурации ПУ " \
                            f"в ф.'mass_log_line_multi.json'."
                        printWARNING(a_err_txt)
                        keystrokeEnter()

                    else:
                        log_line_txt="\n".join(mass_result_dic["log_line_file_list"])

                    a_txt="\nСодержимое log-файла программы " \
                        "MassProdAutoConfig.exe: \n"
                    log_line_txt=a_txt+log_line_txt+ \
                        "\n---Конец блока записей из log-файла.\n"
                    
                    fileWriter(default_filename_ext, "a", "", log_line_txt,
                        "Сохранение в отчет содержимое log-файла.")
                    
    
    if rep_err_send_mail!="0" and external_condition!=None and \
        external_condition!="":
        res=sendMailErrRepMeter(default_filename_ext, rep_err_send_mail,
            external_condition, meter_tech_number, workmode, meter_type)
        if res[0]=="1":
            a_txt="Отправлено уведомление об отрицательном результате " \
                "проверки ПУ"
            printGREEN(f"{a_txt}.")

            a_to_mail=res[2]

            if len(a_to_mail.split(","))>1:
                a_txt=a_txt+f" на адреса: {a_to_mail}."

            else:
                a_txt=a_txt+f" на адрес: {a_to_mail}."

            fileWriter(default_filename_ext, "a", "", f"{a_txt}\n", 
                "Сохранение в отчет информации об отправке "
                "уведомления об отрицательном результате проверки ПУ.")
    

    if no_data_in_SUTP_send_mail!="0" and rep_remark_txt!=None and \
        rep_remark_txt!="":
        res=sendMailNoDataInSUTP(no_data_in_SUTP_send_mail, rep_remark_txt,
            meter_tech_number, workmode, meter_type)
        if res[0]=="1":
            a_txt="Отправлено уведомление об отсутствии данных в СУТП"
            printGREEN(f"{a_txt}.")

            a_to_mail=res[2]

            if len(a_to_mail.split(","))>1:
                a_txt=a_txt+f" на адреса: {a_to_mail}."

            else:
                a_txt=a_txt+f" на адрес: {a_to_mail}."

            fileWriter(default_filename_ext, "a", "", f"{a_txt}\n", 
                "Сохранение в отчет информации об отправке "
                "уведомления об отсутствии информации в СУТП.")
    
    
    if workmode == "эксплуатация" and rep_copy_public=="1":
        toCopyReportFile(default_dirname, work_dirname, dirname_sos, filename1)
        
    reestr_dic["employee_ip_adr"]=employee_ip_adr

    reestr_dic["sutp_transm"]=sutp_transm
    reestr_json=json.dumps(reestr_dic, ensure_ascii=False)
    reestr_json=reestr_json.replace("\\n","")

    _,_, reestrname= getUserFilePath('reestr_meter_json.txt', workmode=workmode)
    if reestrname=="":
        return ["0", "Не найден путь к ф.'reestr_meter_json.txt'"]
    fileWriter(reestrname, "a","",reestr_json+"\n", \
        "Запись в файл reestr_meter.txt в локальной папке")
    
    if workmode == "эксплуатация":
        _,_, reestrname= getUserFilePath('public_reestr_meter_json.txt',
            workmode=workmode)
        if reestrname=="":
            return ["0", "Не найден путь к ф.'public_reestr_meter_json.txt'"]
        fileWriter(reestrname, "a", "", reestr_json+"\n", \
                    "Запись в файл reestr_meter.txt в общей папке","no", dirname_sos)

    
    saveStatistic(employees_name=employees_name,meter_type=meter_typegroup, 
                  grade=int(meter_next_step))
    
    title_search = "Статистика по проверке ПУ"
    res = searchTitleWindow(title_search)
    if res[0] != "1":
        getLocalStatistic(date_filter="",meter_type_filter=meter_typegroup)
        
    return 



def getSutpToSaveDescript(sutp_to_save: str):

    a_dic={"0": ["Автоматическое сохранение результата проверки ПУ в СУТП отключено.", 
        bcolors.FAIL], 
        "01": ["Автоматическое сохранение результата проверки ПУ в СУТП отключено.", 
        bcolors.FAIL],
        "02": ["Автоматическое сохранение результата проверки ПУ в СУТП отключено.", 
        bcolors.FAIL],
        "03": ["Автоматическое сохранение результата проверки ПУ в СУТП отключено.", 
        bcolors.FAIL],
        "1": ["Результат проверки ПУ в СУТП сохраняется в ручном режиме.", 
        bcolors.WARNING], 
        "2": ["Включено автоматическое сохранение результата проверки в СУТП:годен/брак.", 
        bcolors.OKGREEN],
        "3": ["В СУТП сохраняется информация только о выявленном браке.", 
        bcolors.OKGREEN]}
    a_descript=a_dic.get(sutp_to_save, "Не указан способ сохранения "
        "результата проверки в СУТП.")[0]
    a_color=a_dic.get(sutp_to_save, bcolors.FAIL)[1]
    return ["1", "Описание сформировано.", a_descript, a_color]



def inputResultExtInspection(header="", only_defects="0", 
    defect_auto_list=[], add_defects_txt_list=[], 
    param_mandatory_filter="прочие", param_user_filter="прочие", 
    workmode="эксплуатация", mode="новый"):
    
    
    def innerChangeVarDefectsDic(all_defects_dic_list=[], defects_list_1=[]):

        res=toformatNow()
        dt_sec=res[3]
        dt_t=res[0]
        for i in range(0, len(all_defects_dic_list)):
            a_dic=all_defects_dic_list[i]
            j=0
            while j!=len(defects_list_1):
                a_cur=defects_list_1[j]
                if a_cur==a_dic["defects"]:
                    a_dic["lastdt"]=dt_sec
                    a_dic["lastdt_t"]=dt_t
                    a_dic["count"]+=1
                    all_defects_dic_list[i]=a_dic
                    del defects_list_1[j]
                    j-=1
                j+=1
        return all_defects_dic_list, defects_list_1
    
    
    defects_list = []
    defects_txt = ""
    err_txt = ""
    defects_list_ret=[]

    if len(defect_auto_list)>0:
        defects_list.extend(defect_auto_list)

    defects_mandatory_list=[]
    
    file_name_alluser_defects="defects_alluser_dic.json"
    res=readGonfigValue(file_name_alluser_defects,
        [],{}, workmode, "0")
    if res[0]=="2":
        print(f"{bcolors.WARNING}Отсутствует доступ к "
            f"файлу с обязательными описаниями дефектов в "
            f"общей папке.{bcolors.ENDC}\n"
            f"{bcolors.WARNING}Использую локальный файл.{bcolors.ENDC}")
        file_name_alluser_defects="defects_alluser_dic_local.json"
        res=readGonfigValue("defects_alluser_dic_local.json",
            [],{}, workmode,"0")
    defects_mandatory_dic=res[2].get("обязательные",{})
    defects_alluser_mandatory_dic_list=defects_mandatory_dic.get(
        param_mandatory_filter,[])
    for a_cur_dic in defects_alluser_mandatory_dic_list:
        defects_mandatory_list.append(a_cur_dic["defects"])
    
    defects_alluser_dic=res[2].get("пользовательские",{})
    defects_alluser_dic_list = defects_alluser_dic.get(
        param_mandatory_filter, [])
    defects_alluser_list=[]
    for a_cur in defects_alluser_dic_list:
        a_txt = a_cur.get("defects", "")
        if a_txt!="":
            defects_alluser_list.append(a_txt)
    

    file_name_edit_dic={"новый":"defects_user_dic",
        "редактирование all":"defects_alluser_dic",
        "редактирование user":"defects_user_dic"}
    file_name_edit_base=file_name_edit_dic[mode]

    file_name_edit=file_name_edit_base+".json"
    res=readGonfigValue(file_name_edit,[],{}, workmode, "0")
    if res[0] == "2":
        print(f"{bcolors.WARNING}Отсутствует доступ к "
            f"файлу с пользовательскими описаниями дефектов в "
            f"общей папке.{bcolors.ENDC}\n"
            f"{bcolors.WARNING}Использую локальный файл.{bcolors.ENDC}")
        file_name_edit=file_name_edit_base+"_local.json"
        res=readGonfigValue(file_name_edit,[],{}, workmode, "0")
    defect_user_dic = res[2].get(param_user_filter, {})
    defect_user_dic_list = defect_user_dic.get(param_mandatory_filter,[])
    lastdt_list=[]
    for a_cur in defect_user_dic_list:
        a_dt=a_cur.get("lastdt",None)
        if a_dt!=None and (not a_dt in lastdt_list):
            lastdt_list.append(a_dt)
    lastdt_list.sort(reverse=True)
    defects_user_list=[]
    for dt_cur in lastdt_list:
        for a_cur in defect_user_dic_list:
            if a_cur["lastdt"]==dt_cur:
                a_txt=a_cur.get("defects","")
                if a_txt!="":
                    defects_user_list.append(a_txt)
    
    a_list=[]
    if mode=="новый":
        for a_cur in defects_alluser_list:
            res=appendTxtList(a_cur, defects_user_list)
            if res[0]=="1":
                a_list=res[2].copy()
        if len(a_list)>len(defects_user_list):
            defects_user_list=a_list.copy()
        defects_user_list.sort()
        defects_user_list=defects_user_list[0:19]

        a_list = defects_mandatory_list+defects_user_list
        defects_user_list=a_list.copy()
        
        if len(add_defects_txt_list)>0:
            add_defects_txt_list.extend(defects_user_list)
            defects_user_list=add_defects_txt_list.copy()
    defects_user_file_list=defects_user_list.copy()
    
    if "редактирование" in mode:
        defects_list=defects_user_list.copy()


    cicl_edit = True
    first_pass_cicle=True
    while cicl_edit:
        if first_pass_cicle==False:
            os.system('CLS')
        first_pass_cicle=False
        txt1_1=""
        speс_keys_hidden=[]
        spec_list=["добавить новый дефект", "отменить ввод", "закончить ввод"]
        spec_keys=["+", "/", "\r"]
        if "редактирование" in mode:
            spec_list=['добавить новый дефект', 'закончить ввод']
            spec_keys=["+", "\r"]    
        list_txt=defects_user_list.copy()
        spec_id=spec_list.copy()

        if len(defects_list)>0:
            txt1_1=(f'{bcolors.OKBLUE}Указаны следующие дефекты:' 
                    f'{bcolors.ENDC}')
            if "редактирование" in mode:
                txt1_1=(f'{bcolors.OKBLUE}Список вариантов дефектов:' 
                    f'{bcolors.ENDC}')
            i_num=0
            for a in defects_list:
                txt1_1=f'{txt1_1}\n{i_num}. {a}'
                speс_keys_hidden.append(f"-{i_num}")
                speс_keys_hidden.append(f"--{i_num}")
                i_num+=1
                try:
                    index = list_txt.index(a)
                    del list_txt[index]
                except ValueError:
                    pass
            txt1_1=f'{txt1_1}\n\n'
            if err_txt!="":
                txt1_1=f'{txt1_1}{bcolors.WARNING}{err_txt}{bcolors.ENDC}\n'
            txt1_1=f'{txt1_1}{bcolors.OKBLUE}"-" и номер пункта - ' \
                f'редактирование описания дефекта{bcolors.ENDC}\n' \
                f'{bcolors.OKBLUE}"--" и номер пункта - ' \
                f'удаление описания дефекта из списка{bcolors.ENDC}'
        else:
            spec_list=["добавить новый дефект", "замечаний нет"]
            spec_keys=["+", "\r"]
            spec_id = spec_list.copy()
            if err_txt!="":
                txt1_1=f'{bcolors.WARNING}{err_txt}{bcolors.ENDC}\n'
        err_txt=""
        if len(defects_user_list)>0:
            txt1_1=f"{txt1_1}\n{bcolors.OKBLUE}Выберите замечание из списка или " \
                f"добавьте свое замечание.{bcolors.ENDC}"

        if header!="":
            txt1_1=f'{header}\n{txt1_1}'

        list_id = list_txt.copy()
        oo=questionFromList(f"{bcolors.OKBLUE}", txt1_1, list_txt,
            list_id, "",spec_list,spec_keys,spec_id, 1, 1, 0,
            speс_keys_hidden)
        defect_descript=oo
        if oo=="закончить ввод" or oo=="замечаний нет":
            if "редактирование"in mode:
                defect_user_dic_list.clear()
            else:
                if len(defects_list)==0:
                    return ["0", "", []]
                defects_txt="\n".join(defects_list)
                defects_list_ret=defects_list.copy()

                res = innerChangeVarDefectsDic(
                    defects_alluser_mandatory_dic_list, defects_list)
                defects_alluser_mandatory_dic_list = res[0]
                defects_list=res[1]

                if len(defects_alluser_mandatory_dic_list)>0:
                    defects_mandatory_dic[param_mandatory_filter] = \
                        defects_alluser_mandatory_dic_list

                
                res = innerChangeVarDefectsDic(defects_alluser_dic_list,
                    defects_list)
                defects_alluser_dic_list=res[0]
                defects_list = res[1]

                if len(defects_alluser_dic_list)>0:
                    defects_alluser_dic[param_mandatory_filter] = defects_alluser_dic_list

                out_dic = {}
                out_dic["обязательные"] = defects_mandatory_dic.copy()
                out_dic["пользовательские"] = defects_alluser_dic.copy()
                res = saveConfigValue(
                    file_name_alluser_defects, out_dic, workmode)
                if res[0] == "0":
                    res = saveConfigValue(
                        "defects_alluser_dic_local.json", out_dic, workmode)
                
                
                res=innerChangeVarDefectsDic(defect_user_dic_list, defects_list)
                defect_user_dic_list = res[0]
                defects_list=res[1]
                
                if len(defects_list) > 0:
                    j=0
                    while j<len(defects_list):
                        defect_new_cur=defects_list[j]
                        if any(ch in '0123456789/*-+<>()&^%$#!\`~?[]{}=_|'
                            for ch in defect_new_cur):
                            del defects_list[j]
                            j-=1
                        j+=1
                
            if len(defects_list) > 0:
                res=toformatNow()
                dt_sec = res[3]
                dt_t = res[0]
                for defect_new_cur in defects_list:
                    a_new_dic={"defects": defect_new_cur,"lastdt": dt_sec,
                    "lastdt_t":dt_t, "count": 1}
                    defect_user_dic_list.append(a_new_dic)
            

            if len(defect_user_dic_list)==0:
                if param_mandatory_filter in defect_user_dic:
                    del defect_user_dic[param_mandatory_filter]
                    if len(defect_user_dic)==0:
                        defect_user_dic["прочие"]=[]
            else:
                defect_user_dic[param_mandatory_filter] = defect_user_dic_list.copy()
                
            if len(defect_user_dic)>0 or "редактирование"in mode:
                out_dic = {}
                out_dic[param_user_filter] = defect_user_dic
                res = saveConfigValue(file_name_edit, out_dic, workmode)
                if res[0]=="0":
                    file_name_edit=file_name_edit_base+"_local.json"
                    res = saveConfigValue(file_name_edit, out_dic, workmode)

            return ["0",defects_txt, defects_list_ret]
        
        elif oo=="отменить ввод":
            return ["9","", []]
        
        elif oo[0]=="-" and oo[1]!="-":
            i_num=abs(int(oo))
            defect_descript_del=defects_list[i_num]
            txt1=f"\n{bcolors.OKBLUE}Отредактируйте замечание:{bcolors.ENDC}"
            oo=inputSpecifiedKey("",txt1,"",[0],["/"],0, defect_descript_del)
            if oo=="/":
                continue
            defects_list[i_num]=oo
            continue
        
        elif oo[0]=="-" and oo[1]=="-":
            i_num=int(oo[2:])
            defect_descript_del=defects_list[i_num]
            del defects_list[i_num]
            continue
        
        elif oo == "добавить новый дефект":
            txt1=f"\n{bcolors.OKBLUE}Введите свое описание дефекта и " \
                f"нажмите Enter.{bcolors.ENDC}\n" \
                f"{bcolors.OKBLUE}Чтобы отменить ввод - нажмите '/'.{bcolors.ENDC}"
            oo=inputSpecifiedKey("",txt1,"",[0],["/"],0)
            defect_descript=oo
            if oo=="/":
                continue
            defect_descript=defect_descript.strip(" ")
        
        res=appendTxtList(defect_descript, defects_list)
        if res[0]=="0":
            err_txt=f"Замечание '{defect_descript}' уже имеется в списке."



def readDefaultValue(value_dict={}):
    keys_list=list(value_dict.keys())
    for key in keys_list:
        globals()[key]=value_dict[key]


def readActualVersionValue():
    global actual_version_dict
    for i in actual_version_dict:
        globals()[i[0]]=i[1]
    return



def writeDefaultValue(dict={}):
    global default_value_dict

    if len(dict)==0:
        dict = default_value_dict
    keys_list=list(dict.keys())
    for key in keys_list:
        if key in globals():
            dict[key] = globals()[key]
    return dict



def optoRunVarRead():

    global default_value_dict  # словарь со зн-ями по умолчанию

    default_value_dict=getDefaultValue()

    res=readGonfigValue(file_name_in="opto_run.json",
        var_name_list=[],default_value_dict=default_value_dict)
    if res[0]!="1":
        txt1=f"{bcolors.WARNING}При формировании конфигурационных значений " \
            f"для работы программы возникла ошибка.\n" \
            f"{bcolors.OKBLUE}Нажмите Enter"
        print (txt1)
        oo=questionSpecifiedKey("","",["\r"],"",1)
        return ["0", "Ошибка при получении значений по умолчанию.", {}]
    default_value_dict=res[2]
    return default_value_dict



def saveStatistic(employees_name: str, meter_type: str, grade: int, workmode="эксплуатация"):
    _, _, dir_name=getUserFilePath(file_name="localStatistic",
        only_dir="1",workmode=workmode)
    if dir_name=="":
        sys.exit()
    t1=threading.Thread(target=saveStatisticThread, args=(employees_name, \
        meter_type, grade, dir_name,), daemon=False)
    t1.start()


    return



def saveStatisticThread(employees_name: str, meter_type: str, grade: int, dir_name: str):
    date_cur = toformatNow()[1]
    employee_ip_adr=getLocalIP()
    file_name = os.path.join(dir_name, f"{employee_ip_adr}_reestr_statistic.json")
    grade0=0
    grade1=0
    if grade == 1:
        grade1 = 1
    else:
        grade0=1
    statistic_list=[]
    if os.path.exists(file_name) and os.path.getsize(file_name)==0:
        os.remove(file_name)

    if os.path.exists(file_name) == False:
        statistic_cur={"Date":date_cur,"FIO":employees_name,"ipAdr":employee_ip_adr, \
            "meterType":meter_type, "grade0":grade0,"grade1":grade1}
        statistic_list.append(statistic_cur)
        try:
            with open(file_name, "w", errors="ignore", encoding='utf-8') as file:
                json.dump(statistic_list, file,  ensure_ascii=False)
            return
        except Exception as e:
            saveStatisticInSharedFolder(employees_name, meter_type, grade)
            return
    
    try:
        with open(file_name, "r", errors="ignore", encoding='utf-8') as file:
            statistic_list = json.load(file)
    except Exception as e:
        txt1_1=f"{bcolors.WARNING}При чтении файла статистики {file_name} " \
            f"возникла ошибка {e}{bcolors.ENDC}\n" \
            f"{bcolors.OKBLUE}Нажмите Enter.{bcolors.ENDC}"
        spec_keys=["\r"]
        questionSpecifiedKey("",txt=txt1_1,specified_keys_in=spec_keys, \
            file_name_mp3="",specified_keys_only=1)
        return
    for statistic_rec in statistic_list:
        if statistic_rec["Date"] == date_cur and \
            statistic_rec["ipAdr"]==employee_ip_adr and \
            statistic_rec["FIO"] == employees_name and \
            statistic_rec["meterType"] == meter_type:
            grade0=statistic_rec["grade0"]
            grade1=statistic_rec["grade1"]
            if grade == 1:
                grade1 += 1
            else:
                grade0 += 1
            statistic_rec["grade0"]=grade0
            statistic_rec["grade1"]=grade1
            try:
                with open(file_name, "w", errors="ignore", encoding='utf-8') as file:
                    json.dump(statistic_list, file, ensure_ascii=False)
                return
            except Exception as e:
                saveStatisticInSharedFolder(employees_name, meter_type, grade)
                return
    statistic_cur={"Date":date_cur,"FIO":employees_name, "ipAdr":employee_ip_adr, \
                "meterType":meter_type, "grade0":grade0,"grade1":grade1}
    statistic_list.append(statistic_cur)
    try:
        with open(file_name, "w", errors="ignore", encoding='utf-8') as file:
            json.dump(statistic_list, file, ensure_ascii=False)
        return
    except Exception as e:
        saveStatisticInSharedFolder(employees_name, meter_type, grade)
        return



def saveStatisticInSharedFolder(employees_name: str, meter_type: str, grade: int):
    _, _, dir_name = getUserFilePath(file_name="statisticShared", 
        only_dir="1",workmode=workmode)
    if dir_name == "":
        sys.exit()
    saveStatisticThread(employees_name=employees_name, meter_type=meter_type,grade=grade,
        dir_name=dir_name)
    return



def TestBattery(test_num=""):
    global default_filename_full, battery_level
    global rep_err_list, clipboard_err_list
    global test_start_time

    battery_dict = {0: "Батарея заряжена", 1: "Батарея скоро будет полностью разряжена".upper(),
                    2: "Батарея полностью разряжена или отсутствует".upper()}
    time_now=toformatNow()[2]
    dt=datetime.strptime(test_start_time,"%d.%m.%Y %H:%M:%S")
    duration_test = str(int(abs(time_now - dt).seconds))
    txt = f"{test_num}. Тест основной батареи (с момента начала " \
        f"проверки ПУ прошло {duration_test} сек.): "
    print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")
    res=readBatteryLevel()
    if res[0]=="0":
        return["0","Ошибка связи",None,""]
    battery_level=res[2]
    battery_level_txt=battery_dict[battery_level]
    txt1= f"Заряд основной батареи: {battery_level_txt}"
    if test_num!="":
        txt1 = f"{test_num}.1. {txt1}"
    if battery_level != 0:
        txt1 = txt1+". Необходимо заменить батарею."
        print(f"{bcolors.FAIL}{txt1}{bcolors.ENDC}")
        rep_err_list.append(f"{battery_level_txt}. Необходимо заменить батарею")
        clipboard_err_list.append(f"{battery_level_txt}. Необходимо заменить батарею")
    else:
        print(f"{bcolors.OKGREEN}{txt1}{bcolors.ENDC}")
    txt = txt+"\n"+txt1
    if test_num!="":
        fileWriter(default_filename_full, "a", "", f"{txt}\n", \
            "Сохранение в отчет результата теста заряда основной батареи",join="on")
    return ["1","Тест пройден",battery_level,battery_level_txt]



def readBatteryLevel():
    battery_level=0
    res=toReadDataFromMeter("0.0.96.6.1.255", 2,"Статус основной батареи ПУ")
    if res[0]==0:
        return ["0","Ошибка связи.",""]
    battery_level=res[1]
    return ["1","Статус батареи получен.", battery_level]



def readMCSoft():
    mc_soft=""
    res=toReadDataFromMeter("0.0.2.164.0.255", 2,
        "Запрос версии ПО МС")
    if res[0]==0:
        return ["0","Ошибка связи.",""]
    mc_soft=res[1]
    return ["1","Номер версии ПО ПУ получен.", mc_soft]



def checkVersMeter(meter_soft:str, device_ver_list=[],print_msg="1"):
    
    meter_ver_list=[]
    for i in device_ver_list:
        meter_ver_list.append(i[0])
    meter_vers_str=", ".join(meter_ver_list)+"."
    txt2=f"Версия ПО прибора учета: {meter_soft}"
    for meter_ver_cur in meter_ver_list:
        if meter_soft==meter_ver_cur:
            if print_msg=="1":
                print (f"{bcolors.OKGREEN}{txt2} соответствует актуальной версии.{bcolors.ENDC}")
            return ["1", "Версия ПО ПУ актуальна.", meter_vers_str]
    if print_msg=="1":
        txt1=f"{bcolors.FAIL}{txt2}. Необходимо обновление ПО.{bcolors.ENDC}\n" \
            f"{bcolors.OKGREEN}Список актуальных версий ПО для ПУ: " \
                f"{meter_vers_str}{bcolors.ENDC}"
        print (txt1)
    return ["2", "Версия ПО ПУ отсутствует в списке актуальных версий.", 
            meter_vers_str]



def checkVersMC(meter_soft:str, mc_soft:str, 
                device_ver_list=[], print_msg="1"):
    
    mc_ver_list=[]
    for i in device_ver_list:
        if i[0]==meter_soft:
            mc_ver_list=i[1]
            break
    if len(mc_ver_list)==0:
        txt1=f"Для ПУ с версией ПО {meter_soft} не удалось подобрать актуальную версию ПО модема."
        if print_msg=="1":
            print(f"{bcolors.WARNING}{txt1}{bcolors.ENDC}")
        return ["3",txt1,""]
    else:
        mc_ver_list_txt=""
        i=0
        for mc_ver_cur in mc_ver_list:
            mc_ver_list_txt=f"{mc_ver_list_txt}{mc_ver_cur}"
            i+=1
            if i<len(mc_ver_list):
                mc_ver_list_txt=f"{mc_ver_list_txt}, "
    for mc_ver_cur in mc_ver_list:        
        if mc_soft == mc_ver_cur:
            txt1=f"Для ПУ с версией ПО {meter_soft} версия ПО модуля связи {mc_soft} " \
                f"соответствует актуальной версии."
            if print_msg=="1":
                print(f"{bcolors.OKGREEN}{txt1}{bcolors.ENDC}")
            return ["1",txt1,mc_ver_list_txt]
    else:
        txt1=f"Для ПУ с версией ПО {meter_soft} версия ПО модуля связи {mc_soft} " \
            f"не соответствует актуальной версии: {mc_ver_list_txt}."
        if print_msg=="1":
            print(f"{bcolors.FAIL}{txt1}{bcolors.ENDC}")
        return ["2",txt1,mc_ver_list_txt]
            


def toReadStaticDataOpto():

    global meter_serial_number          #серийный номер ПУ
    global meter_date_of_manufacture    #дата калибровки ПУ
    global meter_type                   #тип сч-ка (i-prom.1, i-prom-3)
    global meter_type_ep                # тип ПУ, который записан в электронном паспорте (i-prom.1)
    global meter_product_type           # модель ПУ по серийному номеру ПУ
    global meter_model_ep               # модель ПУ из электронного паспорта (i-prom.1-3-1/2-M-R-Y-Y)
    global meter_soft                   #версия ПО ПУ
    global meter_presence_relay         #наличие реле (силового контактора):str да/нет/-
    global meter_phase                  #число фаз у ПУ
    global meter_voltage_str                #сводная строка значений напряжений
    global meter_amperage_str               #сводная строка значений токов
    global meter_voltage_dic    #словарь с мгновенными значениями напряжения
    global meter_amperage_dic   #словарь с мгновенными значениями тока
    global meter_energy_dic     #словарь с данными о накопленной энергии

    global last_clock_sync  #дата и время последней синхронизации часов ПУ
    global timezone         #часовая зона ПУ
 
    global battery_level    #уровень заряда батареи
    
    global magnetic_field   #фиксация наличия магнитного поля
    global vskritie_korpusa #состояние концевика корпуса
    global obzim_result     #обжатие электронной пломбы
    global actual_device_version_list   #список актуальных версий ПО ПУ и модема
    global disconnect_output_state, current_output_status, disconnect_control_status    #состояние силового реле-контактора
    global gsm_soft         #версия ПО модуля связи

    
    static_data_stand = [
        ["meter_serial_number","0.0.96.1.0.255", 2, "Серийный номер ПУ", "utf-8"],
        ["meter_date_of_manufacture","0.0.96.1.4.255", 2, "Дата выпуска ПУ", "utf-8"],
        ["meter_type","0.0.96.1.1.255", 2, "Тип ПУ ", "utf-8"],
        ["meter_model_ep","0.0.96.1.9.255", 2, "Модель ПУ ", "utf-8"],
        ["last_clock_sync", "0.0.96.2.12.255", 2,
            "Последняя корректировки времени ПУ",""],
        ["meter_soft","0.0.96.1.8.255", 2, "Версия ПО ПУ", "utf-8",""],
        ["gsm_soft","0.0.2.164.0.255", 2, "Версия ПО модуля связи",""],
        ["magnetic_field","0.0.96.51.3.255", 2, "Признак фиксации магнитного поля",""],
        ["vskritie_korpusa","0.0.96.51.0.255", 2, "Состояние концевика корпуса",""]
        ]
    
    static_data=static_data_stand
    voltage=""
    voltage1=""
    voltage2=""
    voltage3=""
    amp=""
    amp1=""
    amp2=""
    amp3=""
    energy_consumed=""
    energy_export=""
    
    len_static=len(static_data)
    txt1_1 = "Считываем данные из ПУ"
    with tqdm(total=len_static+4, desc=txt1_1) as bar:
        for i in static_data:
            txt1_1 = f"{i[3]}"
            bar.set_postfix_str(txt1_1)
            bar.update(1)
            res = toReadDataFromMeter(i[1],i[2],i[3]+": ",i[4])
            if res[0] == 0:
                return 0
            elif res[0] == 1:
                globals()[i[0]]=res[1]
            else:
                print(f"{bcolors.WARNING}.Ошибка. Считывание данных из ПУ: "+ \
                        f"неизвестный код возврата от toReadDataFromMeter.")
                toCloseConnectOpto()
                return 2

        bar.set_postfix_str("Определяем наличие силового реле")
        bar.update(1)
        meter_presence_relay=""
        sheet_name = "Product1"
        res = toGetProductInfo2(meter_serial_number, sheet_name)
        if res[0] == "0":
            txt1_2 = f"Проверка прервана, т.к. не удалось определить модель ПУ по его номеру."
            print(f"{bcolors.FAIL}{txt1_2}{bcolors.ENDC}")
            toCloseConnectOpto()
            return 0
        meter_product_type = res[2]
        meter_product_group_1=res[4]
        meter_type_1=res[29]
        meter_presence_relay = res[13]
        meter_type_ep=meter_type
        meter_phase=res[18]

        if meter_type!=meter_type_1:
            meter_type_select=meter_type_1
            cicl1=True
            while cicl1:
                txt1_1=f"\n{bcolors.FAIL}Тип ПУ, указанный в электронном паспорте '{meter_type}'" \
                        f"отличается от типа,{bcolors.ENDC}\n" \
                        f"{bcolors.FAIL}который определен по серийному номеру '{meter_type_1}'.{bcolors.ENDC}\n" \
                        f"{bcolors.OKBLUE}Укажите фактический тип ПУ из списка:{bcolors.ENDC}"
                list_txt=[meter_type,meter_type_1]
                list_id=list_txt
                cur_id=meter_type_select
                spec_list=["ok","Прекратить проверку"]
                spec_keys=["\r","/"]
                spec_id=["ok","/"]
                oo=questionFromList(colortxt=bcolors.OKBLUE,txt1=txt1_1, list_txt=list_txt, \
                    list_id=list_id, cur_id=cur_id, spec_list=spec_list, spec_keys=spec_keys, \
                    spec_id=spec_id)
                if oo=="9":
                    toCloseConnectOpto()
                    return 9
                elif oo=="ok":
                    meter_type=meter_type_select
                    break
                else:
                    meter_type_select=oo
        
        bar.set_postfix_str("Данные о часовой зоне")
        bar.update(1)
        cicl1=True
        while cicl1:
            try:
                timezone = GXDLMSClock("0.0.1.0.0.255")
                timezone = int(reader.read(timezone, 3))
                break
            except Exception as e:
                oo = communicationTimoutError("Считывание установленного часового пояса в ПУ: ", e.args[0])
                if oo=="0" or oo == "-1":
                    toCloseConnectOpto()
                    return 0
        
        bar.set_postfix_str("Значения мгновенного напряжения/тока")
        bar.update(1)
        meter_voltage_str=""
        meter_amperage_str=""
        cicl1 = True
        while cicl1:
            try:
                if meter_phase == "3":
                    voltage1 = GXDLMSRegister("1.0.32.7.0.255")
                    v = reader.read(voltage1, 2)
                    voltage1 = str(int(v / 1000))
                    voltage2 = GXDLMSRegister("1.0.52.7.0.255")
                    v = reader.read(voltage2, 2)
                    voltage2 = str(int(v / 1000))
                    voltage3 = GXDLMSRegister("1.0.72.7.0.255")
                    v = reader.read(voltage3, 2)
                    voltage3 = str(int(v / 1000))
                    meter_voltage_str=f"{voltage1}/{voltage2}/{voltage3}"

                    amp1 = GXDLMSRegister("1.0.31.7.0.255")
                    v = reader.read(amp1, 2)
                    amp1 = str(v)
                    amp2 = GXDLMSRegister("1.0.51.7.0.255")
                    v = reader.read(amp2, 2)
                    amp2 = str(v)
                    amp3 = GXDLMSRegister("1.0.71.7.0.255")
                    v = reader.read(amp3, 2)
                    amp3 = str(v)
                    meter_amperage_str=f"{amp1}/{amp2}/{amp3}"
                    
                elif meter_phase == "1":
                    voltage = GXDLMSRegister("1.0.12.7.0.255")
                    v = reader.read(voltage, 2)
                    voltage = str(int(v / 1000))
                    meter_voltage_str=f"{voltage}"

                    amp = GXDLMSRegister("1.0.11.7.0.255")
                    v = reader.read(amp, 2)
                    amp = str(v)
                    meter_amperage_str=f"{amp}"
                break
            except Exception as e:
                oo = communicationTimoutError(
                    "Считывание значения мгновенного напряжения/тока: ", e.args[0])
                if oo == "0" or oo == "-1":
                    toCloseConnectOpto()
                    return 0
        
        a_reg=GXDLMSRegister("1.0.1.8.0.255")
        energy_consumed=str(int(reader.read(a_reg, 2)/1000))
        a_reg=GXDLMSRegister("1.0.2.8.0.255")
        energy_export=str(int(reader.read(a_reg, 2)/1000))
        meter_energy_dic={"energy_consumed":energy_consumed,
                          "energy_export":energy_export}
        
        
        bar.update(1)
        if meter_presence_relay== "да":
            bar.set_postfix_str("Cостояние реле")
            cicl1 = True
            while cicl1:
                try:
                    disconnect_control = GXDLMSDisconnectControl("0.0.96.3.10.255")
                    disconnect_output_state = reader.read(disconnect_control, 2)
                    current_output_status = reader.read(disconnect_control, 3)
                    disconnect_control_status = reader.read(disconnect_control, 4)
                    break
                except Exception as e:
                    oo = communicationTimoutError(
                        "Чтение состояния и режима работы реле:", e.args[0])
                    if oo == "0" or oo == "-1":
                        toCloseConnectOpto()
                        return 0
                
            
    actual_device_vers_dict = {"i-prom.1": device_1, "i-prom.3-1": device_3, 
        "i-prom.3-3": device_3, "i-prom.3-3T": device_3T,
        "i-prom.3-3Z": device_3, "i-prom.3Z": device_3, "i-prom.3Z-3": device_3,
        "i-prom.3Z-3T": device_3T}
    actual_device_version_list=actual_device_vers_dict[meter_product_group_1]



def getLocalStatistic(date_filter: str, fio_filter="", 
    meter_type_filter="",workmode="эксплуатация"):
    if date_filter == "":
        date_filter = toformatNow()[1]
    _, _, dir_name = getUserFilePath(file_name="localStatistic", 
        only_dir="1",workmode=workmode)
    if dir_name == "":
        sys.exit()
    employee_ip_adr = getLocalIP()
    file_name = os.path.join(
        dir_name, f"{employee_ip_adr}_reestr_statistic.json")
    
    if os.path.exists(file_name):
        if os.path.getsize(file_name)==0:
            os.remove(file_name)
    
    else:
        txt1 = f"{bcolors.OKGREEN}Статистика по проверке ПУ на данном рабочем месте за {date_filter} отсутствует.{bcolors.ENDC}"
        if dir_name[0:5] == "\\\\":
            txt1 = f"{bcolors.WARNING}Нет доступа к сетевой папке {dir_name}.{bcolors.ENDC}"
        else:
            if not os.path.isdir(dir_name):
                txt1 = f"{bcolors.WARNING}Отсутствует папка {dir_name}.{bcolors.ENDC}"
        print(txt1)
        return

    try:
        with open(file_name, "r", errors="ignore", encoding='utf-8') as file:
            statistic_list = json.load(file)
    except Exception:
        printWARNING(f"Ошибка при чтении локальной статистики из ф.{file_name}.")
        return
    
    grade0_day_all = 0
    grade1_day_all = 0
    employee_list = []
    heading = False
    for statistic_rec in statistic_list:
        if statistic_rec["Date"] == date_filter:
            if heading == False:
                txt1 = f"{bcolors.OKGREEN}Статистика по проверке ПУ на данном рабочем месте за {date_filter}:{bcolors.ENDC}"
                print(txt1)
                if fio_filter != "":
                    txt1 = f"{bcolors.OKGREEN}установлен фильтр по сотруднику: {fio_filter}.{bcolors.ENDC}\n "
                    print(txt1)
                if meter_type_filter != "":
                    txt1 = f"{bcolors.OKGREEN}установлен фильтр по типу ПУ: {meter_type_filter}.{bcolors.ENDC}\n "
                    print(txt1)
                heading = True
            fio = statistic_rec["FIO"]
            if fio in employee_list or (fio_filter != "" and fio != fio_filter):
                continue
            employee_list.append(fio)
            if fio_filter == "":
                print(f"{bcolors.OKGREEN}Сотрудник: {fio}{bcolors.ENDC}")
            grade0 = 0
            grade1 = 0
            grade0_day = 0
            grade1_day = 0
            grade_sum = grade0+grade1
            for statistic_rec_1 in statistic_list:
                if statistic_rec_1["Date"] == date_filter and statistic_rec_1["FIO"] == fio:
                    meter_type = statistic_rec_1["meterType"]
                    if meter_type_filter == "" or (meter_type_filter != "" and meter_type == meter_type_filter):
                        grade0 = statistic_rec_1["grade0"]
                        grade1 = statistic_rec_1["grade1"]
                        grade_sum = grade0+grade1
                        grade0_day = grade0_day+grade0
                        grade1_day = grade1_day+grade1
                        print(f"{bcolors.OKGREEN}{meter_type}:\tна поверку, шт: {grade1}"
                              f"\tв ремонт, шт: {grade0}\t\tИтого, шт: {grade_sum}{bcolors.ENDC}")
            grade_day_sum = grade1_day+grade0_day
            if meter_type_filter != "" and grade_day_sum == 0:
                print(f"{bcolors.OKGREEN}{meter_type_filter}:\tна поверку, шт: {grade1}"
                      f"\tв ремонт, шт: {grade0}\t\tИтого, шт: {grade_sum}{bcolors.ENDC}")
            print(f"{bcolors.OKGREEN}{'-'*86}{bcolors.ENDC}")
            print(f"{bcolors.OKGREEN}ВСЕГО за день:\tна поверку, шт: {grade1_day}"
                  f"\tв ремонт, шт: {grade0_day}\t\tИтого, шт: {grade_day_sum}{bcolors.ENDC}\n")
            grade0_day_all = grade0_day_all+grade0_day
            grade1_day_all = grade1_day_all+grade1_day
    if len(employee_list) > 1:
        grade_day_all_sum = grade0_day_all+grade1_day_all
        print(f"{bcolors.OKGREEN}{'='*86}{bcolors.ENDC}")
        print(f"{bcolors.OKGREEN}ВСЕГО за день:\tна поверку, шт: {grade1_day_all}"
              f"\tв ремонт, шт: {grade0_day_all}\t\tИтого, шт: {grade_day_all_sum}{bcolors.ENDC}")
    elif len(employee_list) == 0:
        txt1 = f"{bcolors.OKGREEN}Статистика по проверке ПУ на данном рабочем месте " \
            f"за {date_filter} отсутствует.{bcolors.ENDC}"
        print(txt1)
    return



def toCompressOfSeal(cover_count):

    res=checkSealCompress()
    if res[0]=="0":
        return ["0","Ошибка связи."]
    obzim_result = res[2] 
    if obzim_result==5:
        return ["1","Обжатие выполнено."]
    
    obzim_tst_konceviki = True  # метка зажатия всех концевиков

    res=checkMagneticField()
    if res[0] == "0":
        return ["0","Ошибка связи."]
    magnetic_status=res[3]
    if magnetic_status != 0:
        res=clearSignMagneticField()
        if res[0]=="0":
            return ["0","Ошибка связи."]
        elif res[0]=="2":
            return ["2","Не удалось сбросить признак магнитного поля."]
    magnetic_status=res[3]

    res=checkBtnBody()
    if res[0]=="0":
        return ["0","Ошибка связи."]
    vskritie_korpusa = res[2]
    if vskritie_korpusa != 0:
        return ["2","Концевик корпуса разомкнут."]
    
    res=questionCloseCover(cover_count)
    if res[0]=="0":
        return ["0","Ошибка связи."]
    elif res[0]=="9":
        return ["9","Пользователь прервал выполнение проверки ПУ."]
    elif res[0]=="8":
        return ["8","Пользователь пропускает текущий тест."]
    
    elif res[0]=="2":
        return ["5","Неисправен концевиу(и)."]

    while True:
        try:
            obzim = GXDLMSClient.createObject(ObjectType.DATA)
            obzim.logicalName = "0.0.96.51.6.255"
            obzim.value = 1
            obzim.setDataType(2, DataType.UINT8)
            reader.write(obzim, 2)
            break
        except Exception as e:
            oo = communicationTimoutError(
                "Обжатие электронной пломбы:", e.args[0])
            if oo == "0" or oo == "-1":
                return ["0","Ошибка связи."]
    
    res=checkSealCompress()
    if res[0]=="0":
        return ["0","Ошибка связи."]
    obzim_result = res[2]
    if obzim_result==5:
        return["1","Пломба обжата."]
    return["4","Не удалось обжать пломбу."]



def testCompressOfSeal(test_num:str, cover_count: int):
    
    txt = f"{test_num}. Обжатие электронной пломбы: "
    res=checkSealCompress()
    if res[0]=="0":
        return ["0","Ошибка связи.",""]
    obzim_result = res[2] 
    if obzim_result >= 5:
        txt = txt+" было выполнено ранее"
        print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")
    else:
        res=toCompressOfSeal(cover_count)
        txt_err=f"{txt} не выполнено: {res[1]}."
        if res[0]=="0":
            return ["0","Ошибка связи.",""]
        elif res[0]=="1":
            txt = txt+" выполнено"
            print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")
        elif res[0]=="2" or res[0]=="3" or res[0]=="4":
            print(f"{bcolors.WARNING}{txt_err}{bcolors.ENDC}")
            txt=txt_err
            return ["2","Не удалось обжать пломбу",txt]
        elif res[0]=="8":
            txt = txt+" не выполнено: пропущено" 
            print(f"{bcolors.WARNING}\n{txt}{bcolors.ENDC}")
            return ["8","Тест пропущен.",txt]
        elif res[0]=="9":
            if test_num!="":
                return ["9","Проверка прервана пользователем.",""]
    return ["1","Обжатие выполнено.",txt]



def checkSealCompress():
    res = toReadDataFromMeter(
        "0.0.96.51.5.255", 2, "Проверка обжатия электронной пломбы: ")
    if res[0]==0:
        return ["0","Обрыв связи.",None]
    obzim_result = res[1]
    return ["1","Операция выполнена успешно.",obzim_result]
           


def checkMagneticField():
    res = toReadDataFromMeter(
        "0.0.96.51.3.255", 2, "Проверка отсутствия признака магнитного поля: ")
    if res[0] == 0:
        return ["0","Обрыв связи.","",None]
    magnetic_field = res[1]
    magnetic_field_bin = format(magnetic_field, "08b")
    magnetic_field_0 = magnetic_field_bin[-1]
    magnetic_field_2 = magnetic_field_bin[-3]
    magnetic_status_id=-1
    magnetic_status="указан неизвестный код"
    if magnetic_field_0 == '0':
        magnetic_status_id=0
        magnetic_status_txt="не было зафиксировано"
    elif magnetic_field_2 == '0':
        magnetic_status_id=1
        magnetic_status_txt="было зафиксировано, сейчас отсутствует"
    elif magnetic_field_2=='1':
        magnetic_status_id=2
        magnetic_status_txt="было зафиксировано, сейчас присутствует"    
    return ["1","Операция выполнена успешно.", magnetic_field_bin, 
            magnetic_status_id,magnetic_status_txt]
    


def clearSignMagneticField():
    while True:
        try:
            reset_MF = GXDLMSClient.createObject(ObjectType.DATA)
            reset_MF.logicalName = "0.0.96.51.7.255"
            reset_MF.value = 1
            reset_MF.setDataType(2, DataType.UINT8)
            reader.write(reset_MF, 2)
            break
        except Exception as e:
            oo = communicationTimoutError("Сброс признака магнитного поля:", e.args[0])
            if oo=="0" or oo == "-1":
                return ["0","Обрыв связи с ПУ."]
    
    res=checkMagneticField()
    if res[0]=="0":
        return ["0",res[1]]
    magnetic_status_id=res[3]
    if magnetic_status_id==0:
        return ["1","Признак фиксации магнитного поля сброшен."]
    return ["2", "Не удалось сбросить признак фиксации магнитного поля."]
    
    

def checkBtnBody():
    res = toReadDataFromMeter(
        "0.0.96.51.0.255", 2, "Проверка зажатия концевика корпуса: ")
    if res[0]==0:
        return ["0","Обрыв связи.",None]
    vskritie_korpusa = res[1]
    return ["1","Операция выполнена успешно.",vskritie_korpusa]



def offAlarmSignDisplay(visual_control="1"):

    elochka = GXDLMSClient.createObject(ObjectType.DATA)
    elochka.logicalName = "0.0.99.13.168.255"
    while True:
        elochka_status = reader.read(elochka, 2)
        if elochka_status:
            print(f"\n    Убираю символ 'елочка'...", end="")
            elochka.value = False
            elochka.setDataType(2, DataType.BOOLEAN)
            try:
                reader.write(elochka, 2)
                break
            except Exception as e:
                oo = communicationTimoutError("Убирали символ елочка:", e.args[0])
                if oo=="0" or oo=="-1":
                    return ["0","Ошибка связи."]
        else:
            break

    if visual_control=="0":
        return ["1","Признак тревоги сброшен."]
    txt1_1="\n    Проверьте: символ 'елочка' на ЖКИ исчез? 0-нет, 1-да"
    oo=questionSpecifiedKey(bcolors.OKBLUE,txt1_1,["0","1"],"LCDBellOff")
    if oo=="0":
        txt1_1="\n    Вы уверены, что символ 'елочка' на ЖКИ светится? 0-нет, 1-да"
        oo=questionSpecifiedKey(bcolors.WARNING,txt1_1,["0","1"],"LCDBellOn")
        if oo=="1":
            return ["2","Не удалось погасить символ елочка на ЖКИ."]
    return ["1","Признак тревоги сброшен."]
    


def checkBtnTerminalCover():
    res = toReadDataFromMeter("0.0.96.51.1.255", 2, 
        "Проверка зажатия концевиков крышек клеммников: ")
    if res[0]==0:
        return ["0","Обрыв связи.",None]
    vskritie_klemm = res[1]
    return ["1","Операция выполнена успешно.",vskritie_klemm]
        


def questionCloseCover(cover_count=2):
    vskritie_klemm=1
    i=0
    while True:
        res=checkBtnTerminalCover()
        i+=1
        if res[0] == "0":
            return ["0","Обрыв связи."]
        vskritie_klemm = res[2]
        if vskritie_klemm!=0:
            if i>1:
                txt1="\n    Концевик одной или обеих клеммников разжат."
                if cover_count == 1:
                    txt1="\n    Концевик клеммника разжат."
                print(f"{bcolors.WARNING}{txt1}{bcolors.ENDC}")
            txt1="\n    Зажмите концевики крышек информационного и силового клеммников и нажмите 'Enter'."
            a_err_txt="\n    Чтобы записать о неисправности клеммника(-ов) - нажмите '0'."
            if cover_count == 1:
                txt1="\n    Зажмите концевик крышки клеммника и нажмите 'Enter'."
                a_err_txt="\n    Чтобы записать о неисправности клеммника - нажмите '0'."
            txt1=f"{txt1}{a_err_txt}" \
                f"\n    Чтобы пропустить текущий тест - нажмите '8'." \
                f"\n    Для прекращения проведения проверки ПУ нажмите - '/'."
            oo=questionSpecifiedKey(bcolors.OKBLUE,txt1,["\r", "0", "8","/"],"ClampCap2")
            if oo=="/":
                return ["9","Пользователь прервал выполнение проверки."]
            elif oo=="8":
                return ["8","Пользователь пропускает текущий тест."]
            
            elif oo=="0":
                return ["2","Записать замечание о неисправности."]
            
            pause_ui(3)
        else:
            return["1","Концевики крышек клеммников ПУ зажаты."]
    


def restoreCOMPort(com_name:str):

    global default_value_dict   #словарь значений по умолчанию
    global com_opto             #COM-порт, к которому подключен оптопорт
    global com_rs485            #COM-порт, к которому подключен RS-485


    a_dic={"com_opto":["оптопорт","оптопорта"],
        "com_rs485":["преобразователь RS-485","преобразователя RS-485"]}
    print()
    comment_txt=f"{bcolors.OKGREEN}Попробуем восстановить подключение " \
        f"к COM-порту {a_dic.get(com_current,'')[1]}.{bcolors.ENDC}"
    res=getAutoCOMPort(a_dic.get(com_current,'')[0], "1", comment_txt)
    if res[0]=="1":
        com_val=res[2]
        globals()[com_name]=com_val
        default_value_dict = writeDefaultValue(default_value_dict)
        saveConfigValue('opto_run.json', default_value_dict)
        return "1"
    return res[0]



def sendMailErrConfigMeter(history_status_txt="", 
    workmode="эксплуатация", meter_type="",
    meter_config_param_filename="", meter_config_filename=""):
    

    a_meter_soft=""

    mass_prod_vers=""

    res=readGonfigValue("opto_run.json",[],{}, workmode, "1")
    if res[0]!="1":
        return ["0", "Ошибка при чтении ф.opto_run.json."]

    opto_dic=res[2]

    meter_config_check=opto_dic["meter_config_check"]
    
    err_in_log_list=opto_dic["meter_config_res_list"]

    meter_tech_number=opto_dic["meter_tech_number"]

    config_send_mail=opto_dic["config_send_mail"]

    a_meter_soft=opto_dic["meter_soft"]

    if len(err_in_log_list)==0:
        return ["2", "Замечаний к конфигурации нет."]
      
    if meter_config_check[0]=="3":
        res=readGonfigValue("mass_log_line_multi.json",[],{}, workmode, "1")
        if res[0]!="1":
            return ["0", "Ошибка при чтении ф.mass_log_line_multi.json."]

        mask_file_name=res[2].get("mask_file_name", "")
        
        mass_result_dic=res[2].get(meter_tech_number, {})

        if len(mass_result_dic)==0 or mask_file_name=="":
            a_err_txt="При формировании электронного письма выявлено, что " \
                "отсутствует информация о проверке конфигурации ПУ " \
                f"в ф.'mass_log_line_multi.json'."
            printWARNING(a_err_txt)
            keystrokeEnter()
    
        
        log_line_file_list=mass_result_dic["log_line_file_list"]

        log_line_file_txt="\n".join(log_line_file_list)

        res=readGonfigValue("mass_config.json", [], {}, workmode, "1")
        if res[0]=="1":
            mass_prod_vers=res[2]["mass_prod_vers"]

    
    err_in_log_list=insHiphenColor(err_in_log_list, "- ", "")[2]
    err_in_log_txt="<br>".join(err_in_log_list)

    
    subject=f"Замечания к конфигурации ПУ № {meter_tech_number}"
    message_txt=f"При проверке конфигурации ПУ"
    
    if meter_type!="":
        message_txt=f"{message_txt} {meter_type}"

    message_txt=f"{message_txt} № {meter_tech_number}" 
    
    if a_meter_soft!="":
        message_txt=f"{message_txt} с версией ПО ПУ {a_meter_soft}"
    
    message_txt=f"{message_txt} были выявлены следующие замечания:" \
        f"<br>{err_in_log_txt}"

    
    a_txt=""

    if meter_config_param_filename!="" and \
        meter_config_filename!="":
        a_txt="Имя файла конфигурации ПУ, указанного в заказе: " \
            f"{meter_config_param_filename}.<br>" \
            "Имя файла, использованного при конфигурировании ПУ: " \
            f"{meter_config_filename}.<br>"
        
        if meter_config_filename!=meter_config_param_filename:
            a_txt=a_txt + "Имена файлов конфигурации ПУ отличаются.<br>"
    
    
    if a_txt!="":
        a_txt=a_txt+"<br>"

    a_txt=a_txt + "Проверка осуществлялась с помощью программы " \
        "'MassProdAutoConfig.exe'"
    if meter_config_check[0]=="2":
        a_txt=a_txt+" в ручном режиме."

    elif meter_config_check[0]=="3":
        a_txt=a_txt+" в автоматическом режиме.<br>" \
            f"Использовался mask-файл '{mask_file_name}'."
        if mass_prod_vers!="":
            a_txt=a_txt+f"<br>Использовалась версия {mass_prod_vers} программы " \
                "'MassProdAutoConfig.exe'."

    message_txt=f"{message_txt}<br><br>{a_txt}"

    attach_file_list=[]

    if meter_config_check[0]=="3":
        attach_file_list=["Фрагмент log-файла.txt"]
        attach_file_dir="file_attach_mail"
        _,_, attach_file_dir = getUserFilePath(attach_file_dir, "1", workmode=workmode)
        if attach_file_dir=="":
            return ["0","Ошибка при формировании пути до папки " 
                    "'file_attach_mail' ."]

        log_file_path = os.path.join(attach_file_dir, "Фрагмент log-файла.txt")

        fileWriter(log_file_path, "w", "utf-8", f"{log_line_file_txt}\n", \
            "Сохранение в ф.'Фрагмент log-файла.txt' замечаний "
            "к конфигурации ПУ.",join="on")

        if history_status_txt!="":
            attach_file_list.append("История статусов ПУ.txt")
            history_file_path = os.path.join(attach_file_dir, 
                "История статусов ПУ.txt")

            fileWriter(history_file_path, "w", "utf-8", f"{history_status_txt}\n",
                "Сохранение в ф.'История статусов ПУ.txt' историю "
                "статусов ПУ.",join="on")

    
    rec_block_name="конфигурация"
    
    if config_send_mail=="2":
        rec_block_name="Васильеву"

    res=sendMail("send_mail.json", subject, message_txt, attach_file_list, 
        workmode, "1", rec_block_name)
    
    if res[0]=="0":
        return ["0", "Ошибка при отправке письма."]

    return ["1", "Письмо успешно отправлено.", res[2]]



def sendMailErrRepMeter(rep_path: str, rep_err_send_mail:str,
   rep_err_txt: str, meter_tech_number: str, workmode="эксплуатация", 
   meter_type=""):
    

    rep_err_list=rep_err_txt.split("\n")

    if len(rep_err_txt)==0:
        return ["2", "Замечаний к ПУ нет."]
      
    rep_err_list=insHiphenColor(rep_err_list, "- ", "")[2]
    rep_err_txt="<br>".join(rep_err_list)

    
    subject=f"Замечания к состоянию ПУ № {meter_tech_number}"
    message_txt=f"При проверке ПУ"
    
    if meter_type!="":
        message_txt=f"{message_txt} {meter_type}"

    message_txt=f"{message_txt} № {meter_tech_number}" 
    
    message_txt=f"{message_txt} были выявлены следующие замечания:" \
        f"<br>{rep_err_txt}"

    attach_file_list=[]

    rep_file_name=os.path.split(rep_path)[1]
    attach_file_list=[rep_file_name]
    attach_file_dir="file_attach_mail"
    _,_, attach_file_dir = getUserFilePath(attach_file_dir, "1", workmode=workmode)
    if attach_file_dir=="":
        return ["0","Ошибка при формировании пути до папки " 
                "'file_attach_mail' ."]

    rep_attach_file_path = os.path.join(attach_file_dir, rep_file_name)
    
    try:
        shutil.copy2(rep_path, rep_attach_file_path)

    except:
        a_err_txt=f"При копировании файла {rep_file_name} в папку " \
            f"{attach_file_dir} возникла ошибка."
        printWARNING(a_err_txt)
        printWARNING("Уведомление по электронной почте о выявлении " 
            "замечаний к состоянию ПУ не отправлено.")
        return ["0", a_err_txt]
    
    rec_block_name="дефект"
    
    if rep_err_send_mail=="2":
        rec_block_name="Васильеву"

    res=sendMail("send_mail.json", subject, message_txt, attach_file_list, 
        workmode, "1", rec_block_name)
    
    if res[0]=="0":
        return ["0", "Ошибка при отправке письма."]

    return ["1", "Письмо успешно отправлено.", res[2]]



def sendMailNoDataInSUTP(no_data_in_SUTP_send_mail:str,
   rep_remark_txt: str, meter_tech_number: str, workmode="эксплуатация", 
   meter_type=""):
    

    if len(rep_remark_txt)==0:
        return ["2", "Примечаний к процессу проверки ПУ нет."]

    a_list=rep_remark_txt.split("\n")

    remark_list=[]

    for a_remark in a_list:
        if "СУТП" in a_remark:
            remark_list.append(a_remark)

    if len(remark_list)==0:
        return ["2", "Примечаний об отсутствии данных в СУТП нет."]
    
    remark_list=insHiphenColor(remark_list, "- ", "")[2]
    remark_txt="<br>".join(remark_list)

    subject=f"В СУТП отсутствует/некорректна информация о ПУ № {meter_tech_number}"
    message_txt=f"При проверке ПУ"
    
    if meter_type!="":
        message_txt=f"{message_txt} {meter_type}"

    message_txt=f"{message_txt} № {meter_tech_number}" 
    
    message_txt=f"{message_txt} было выявлено, что в СУТП отсутствует " \
        f"или некорректна следующая информация:<br>{remark_txt}"

    attach_file_list=[]


    

    
    rec_block_name="нет данных в СУТП"
    
    if no_data_in_SUTP_send_mail=="2":
        rec_block_name="Васильеву"

    res=sendMail("send_mail.json", subject, message_txt, attach_file_list, 
        workmode, "1", rec_block_name)
    
    if res[0]=="0":
        return ["0", "Ошибка при отправке письма."]

    return ["1", "Письмо успешно отправлено.", res[2]]



def read_opto(read_mode_opto="полная проверка"):  # в протоколе 1


    global default_value_dict  # словарь со зн-ями по умолчанию:
    global employees_name       #ФИО пользователя
    global employee_id          # таб. номер пользователя
    global employee_pw_encrypt  # зашифрованный пароль пользователя для доступа в СУТП
    global rep_copy_public      #метка возможности копирования протокола в общую папку:
    global speaker              #метка вкл/откл диктора. "0"-откл, "1"-вкл
    global number_of_meters     #количество подключаемых ПУ на стенде
    global com_opto             #COM-порт, к которому подключен оптопорт
    global com_rs485             #COM-порт, к которому подключен оптопорт
    global com_current          #вид активного порта для связи с ПУ:"com_opto","com_rs485"
    global meter_color_body     #цвет корпуса ПУ по умолчанию
    global meter_color_body_man     #цвет корпуса ПУ введенный по QR-коду
    global meter_config_check   #метод проверки конфигурации ПУ: 
    global config_send_mail     #отправка сообщения по электронной почте о
    global rep_err_send_mail    #отправка сообщения по электронной почте о
    global no_data_in_SUTP_send_mail    #отправка сообщения по электронной почте о
    global meter_config_res_list    #список с замечаниями по проверке конфигурации ПУ
    global meter_adjusting_clock  # корректировка часов счетчика:
    global meter_type_def       # тип ПУ по умолчанию
    global meter_type           # тип текущего ПУ
    global meter_type_ep        # тип ПУ, которое записано в электронном паспорте (i-prom.1)
    global meter_model_ep       # модель ПУ из электронного паспорта (i-prom.1-3-1/2-M-R-Y-Y)
    global meter_product_type   # модель ПУ, определенная по серийному номеру
    global meter_soft           # версия ПО ПУ
    global meter_serial_number  #серийный номер ПУ (эталон). Порядок выбора:СУТП, электронный паспорт,
    global meter_sn_source      #источник получения серийного номера (СУТП, электронный паспорт, QR-код/надпись)
    global meter_sn_ep          #серийный номер ПУ из электронного паспорта ПУ

    global meter_tech_number    #технический номер ПУ (эталон). Порядок выбора: наклейка на крышке, СУТП
    global meter_tn_source      #источник получения технического номера (наклейка, СУТП)
    global meter_date_of_manufacture    #дата калибровки (выпуска) ПУ
    global meter_presence_relay #наличие реле (силового контактора):str да/нет
    global meter_pw_default     #пароль по умолчанию для подключения к ПУ
    global meter_password_descript  #описание пароля по умолчанию ("Стандартный", "Карелия"...) для подключения к ПУ
    global meter_pw_level       # уровень текущего доступа к ПУ: "High", "Low"
    global meter_pw_default_descript  #описание пароля по умолчанию ("Стандартный высокого уровня", "Карелия"...) для подключения к ПУ
    global meter_pw_low_encrypt     #зашифрованный пароль низкого уровня подключения к ПУ
    global meter_pw_low_descript     #описание текущего пароля низкого уровня подключения к ПУ
    global meter_pw_high_encrypt     #зашифрованный пароль высокого уровня подключения к ПУ
    global meter_pw_high_descript     #описание пароля высокого уровня подключения к ПУ

    global meter_phase          #число фаз у ПУ:"1", "3"
    global meter_voltage_str   #сводная строка значений мгновенных напряжений
    global meter_amperage_str   #сводная строка значений мгновенных токов
    global meter_voltage_dic    #словарь с мгновенными значениями напряжения
    global meter_amperage_dic   #словарь с мгновенными значениями тока
    global meter_energy_dic     #словарь с данными о накопленной энергии
                                
    global voltage,voltage1,voltage2,voltage3
    global electrical_test_circuit  #схема подключения ПУ для проверки: 
    global ctrl_current_electr_test #контрольное значение тока в схеме подключения ПУ

    global rc_serial_number     #серийный номер пульта управления ПУ
    global rc_tech_number       #технический номер пульта управления ПУ
        
    global modem_type_def       # тип модуля связи по умолчанию
    global modem_status         #статус модема по умолчанию: "0"-не будет устанавливаться, "1"-рабочий, "2"-тестовый
    global SIMcard_status       #Статус SIM-карты по умолчанию: 0-не будет устанавливаться, 1-рабочая, 2-тестовая
    global connection_initialized   #метка инициализации канала с ПУ
    global actual_gsm_version   #актуальная версия ПО для GSM Модема
    global actual_device_version    #актуальная верси ПО ПУ
    global actual_device_version_list   #список актуальных версий ПО ПУ и модема
    global device_1,device_3,device_3T  #списки с актуальными версиями ПО для ПУ и модема по типам ПУ
    global filename_rep             #имя файла-отчета без постфикса "_отчет.txt"
    global default_filename_full     #имя отчета (протокола) вместе с именем тек. директории
    global workmode             #метка режима работы программы "тест" - режим теста1, 
    global txt_break_test       #текст " Для прекращения проведения проверки ПУ нажмите '/' и Enter."
    global txt_skip_test        #текст " Чтобы пропустить текущий тест нажмите 8 и Enter."
    global test_result          #строка для хранения результатов теста
    global test_result_buf      #строка для хранения результатов теста для буфера Windows
    global gsm_serial_number    #номер GSM модема (эталон). Порядок выбора:крышка, СУТП
    global gsm_SIM_number       #номер SIM-карты ("0"-номер не требуется)
    global gsm_product_type     #тип модема по серийному номеру
    global gsm_soft             #версия ПО модуля связи
    global mc_checkVersPO_set   #необходимость проверки версии ПО МС ("да"/"нет") из ф.ProductNumber.xlsx
    global mc_on_board          #метка фактической установки МС в ПУ ("0"-отсутствует, "1"-установлен) 
    global last_clock_sync

    global battery_level

    global magnetic_field
    global vskritie_korpusa
    global obzim_result
    global disconnect_output_state, current_output_status, disconnect_control_status
    global data_exchange_sutp   #метка обмена данными с СУТП:"0"-откл.,"1"-вкл.
    global sutp_to_save             #способ записи рез-тов теста в БД СУТП ("0"-отключен, "1"-ручной,
    global order_control        # метка контроля принадлежности ПУ определенному заказу 
    global order_control_descript   # номер и описание контролируемого заказа
    global order_num            #номер заказа проверяемого ПУ
    global order_descript       #описание заказа проверяемого ПУ
    global order_ev

    global test_start_time      #дата и время начала теста
    global rep_err_list         #список с выявленными ошибками для записи в отчет
    global clipboard_err_list   #список выявленных ошибок для сохранения в буфере Windows
    global rep_remark_list      #список с примечаниями для записи в отчет

    global duration_test        #продолжительность проверки ПУ, мин

    txt_break_test="Для прекращения проведения проверки ПУ нажмите '/' и Enter."
    txt_skip_test = "Чтобы пропустить текущий тест нажмите 8 и Enter."
 


    def innerToInitConnectOpto(com_port=com_opto):
        global connection_initialized
        global meter_pw_high_encrypt
        global default_value_dict
        global com_current
        global reader
        global settings
        global meter_config_check, config_send_mail
        global rep_err_send_mail

        if connection_initialized==True:
            return "1"

        while True:
            connection_initialized=False

            res=toInitConnectOpto(com_opto=com_port)
            if res==1:
                connection_initialized=True
                return "1"
            elif res==2:
                colortxt=bcolors.WARNING
                txt1_1="Нет связи с ПУ через оптопорт. Внесете замечания в отчет и в СУТП? 0 - нет, 1 - да"
                oo=questionSpecifiedKey(colortxt, txt1_1, ["0","1"],"")
                if oo=="1":
                    err_list=["нет связи с ПУ через оптопорт"]
                    param_mandatory_filter="оптопорт"
                    param_user_filter=employee_id
                    external_condition = ""
                    txt1_1=f"\n{bcolors.OKBLUE}Запишите выявленные замечания к ПУ № "+ \
                        f"{meter_tech_number}:{bcolors.ENDC}"
                    res = inputResultExtInspection(header=txt1_1, only_defects="1", 
                        defect_auto_list=err_list, param_mandatory_filter=param_mandatory_filter,
                        param_user_filter=param_user_filter, workmode=workmode,
                        mode="новый")
                    external_condition=res[1]
                    if res[0]=="9":
                        print(f"{bcolors.WARNING}\nВвод данных прерван.{bcolors.ENDC}")
                        return "0"
                    if external_condition!="":
                        reestr_dic=innerSetValReestr()
                        res=toSaveResultExtInspection(meter_serial_number,meter_tech_number, 
                            external_condition, employees_name, employee_id, default_filename_ext,
                            default_dirname, work_dirname,dirname_sos,filename_ext,
                            sutp_to_save, data_exchange_sutp, "0", "0", workmode, reestr_dic, 
                            rep_copy_public, meter_config_check, config_send_mail, rep_err_send_mail)
                return "0"
            
            elif res==3:
                colortxt=bcolors.OKBLUE
                txt1_1="Нажмите Enter."
                oo=questionSpecifiedKey(colortxt, txt1_1, ["\r"],"", 1)
                return "0"

            elif res==5:
                res=cryptStringSec("расшифровать", meter_pw_high_encrypt)
                a_password=res[2]
                com_opto=default_value_dict['com_opto']
                com_rs485=default_value_dict['com_rs485']
                a_dic={"com_opto":com_opto, "com_rs485":com_rs485}
                comport=a_dic[com_current]
                serial_num=""
                if com_current=="com_rs485":
                    serial_num=meter_serial_number[-4:]

                res = settingOpt(password=a_password,
                    serial_num=serial_num, comport=comport, msg_print="0", 
                    authentication="High")
                if res[0]=="0":
                    return "0"

                reader=res[1]
                settings=res[2]
                continue

            elif res==4:
                res=restoreCOMPort(com_current)
                if res!="1":
                    return "0"


    def innerSelectActions(msg_err:str, err_list:list, 
        select_mode="1", param_mandatory_filter="прочие",
        menu_item_add_list=[], menu_id_add_list=[], 
        err_no_edit_list=[]):
        
        global meter_tech_number    #технический номер ПУ
        global meter_serial_number
        global employees_name
        global employee_id
        global sutp_to_save
        global data_exchange_sutp
        global config_send_mail
        global rep_err_send_mail    #отправка сообщения по электронной почте о
        global meter_config_check

        nonlocal default_filename_ext
        nonlocal default_dirname
        nonlocal work_dirname
        nonlocal dirname_sos
        nonlocal filename_ext
        
        param_user_filter=employee_id
        if type(err_list)==str:
            err_list=[err_list]
        err_txt="\n".join(err_list)
        err_all_list=err_list.copy()

        header=msg_err

        while True:
            if msg_err!="":
                print(f"{msg_err}")
            txt1="Выберите дальнейшее действие:"
            action_dic={"1": ["Отправить ПУ в ремонт", 
                "Продолжить проверку с фиксацией замечания в списке дефектов"],
                "2": ["Записать замечания в список дефектов", 
                      "Замечаний нет. Продолжить проверку."],
                "3": ["Отправить ПУ в ремонт"],
                "4": ["Продолжить проверку"],
                "5": ["Отправить ПУ в ремонт", 
                      "Продолжить проверку с фиксацией замечания в списке дефектов",
                      "Продолжить проверку с фиксацией замечания в примечании"],
                "6": ["Пропустить данный тест"],
                "7": ["Записать новые замечания и отправить ПУ в ремонт",
                      "Записать новые замечания в список дефектов, затем продолжить проверку",
                      "Замечаний нет. Продолжить проверку"],
                "8": ["Записать новые замечания в список дефектов", 
                      "Замечаний нет. Продолжить проверку."],
                "9": []}
            menu_item_list=action_dic[select_mode]
            action_id_dic={"1": ["ремонт", "проверка"],
                "2": ["замечания", "проверка"],
                "3": ["ремонт"], "4": ["проверка"],
                "5": ["ремонт", "проверка", "проверка без замечания"],
                "6": ["пропустить тест"],
                "7": ["ремонт", "замечания", "проверка"],
                "8": ["замечания", "проверка"],
                "9": []}
            menu_id_list=action_id_dic[select_mode]
            if len(menu_item_add_list)>0 and len(menu_id_add_list)>0:
                menu_item_list.extend(menu_item_add_list)
                for i in range(0,len(menu_id_add_list)):
                    menu_id_add_list[i]=f"#4-{menu_id_add_list[i]}"
                menu_id_list.extend(menu_id_add_list)
            spec_list=["Прервать проверку"]
            spec_keys=["/"]
            spec_id_list=["прервать"]
            oo="ремонт"
            if select_mode!="9":
                oo=questionFromList(bcolors.OKBLUE, txt1, menu_item_list, menu_id_list,
                    "", spec_list, spec_keys, spec_id_list, 1, start_list_num=1)
                print()
            if oo=="прервать":
                print (f"{bcolors.WARNING}Проверка прервана.{bcolors.ENDC}")
                return ["9", "Проверка прервана пользователем.", err_all_list, ""]

            elif oo=="ремонт":
                os.system("CLS")
                err_list.extend(err_no_edit_list)
                external_condition = ""
                
                header=f"{header}\n{bcolors.WARNING}Отправка ПУ № " \
                    f"{meter_tech_number} в ремонт.{bcolors.ENDC}\n"
                
                if data_exchange_sutp=="1":
                    a_mode="1"
                    res=getInfoAboutDevice(meter_tech_number, workmode, "", 
                        "", a_mode)
                    if res[0] in ["1", "2"]:
                        status_txt=res[7]
                        if not status_txt in ["Гравировка пройдена",
                            "Состыкован с МС"]:
                            header=f"{header}{bcolors.WARNING}Текущий статус ПУ: " \
                                f"{status_txt}.\n"

                res = inputResultExtInspection(header=header, only_defects="1", 
                    defect_auto_list=err_list, param_mandatory_filter=param_mandatory_filter,
                    param_user_filter=param_user_filter, workmode=workmode,
                    mode="новый")
                external_condition=res[1]
                if res[0]=="9":
                    print(f"{bcolors.WARNING}\nВвод данных прерван.{bcolors.ENDC}")
                    if select_mode=="9":
                        return ["9", "Ввод данных прерван.", err_all_list, ""]
                    
                    continue

                if external_condition!="":
                    err_all_list=res[2]
                    reestr_dic=innerSetValReestr()
                    res=toSaveResultExtInspection(meter_serial_number, meter_tech_number,
                        external_condition, employees_name, employee_id, default_filename_ext,
                        default_dirname, work_dirname,dirname_sos,filename_ext,
                        sutp_to_save, data_exchange_sutp, "0", "0", workmode, reestr_dic,
                        rep_copy_public, meter_config_check, config_send_mail, rep_err_send_mail)
                return ["2", "ПУ успешно отправили в ремонт", err_all_list, ""]

            elif oo=="замечания":
                os.system("CLS")
                if len(err_no_edit_list)>0:
                    a_err_no_edit_list=err_no_edit_list.copy()
                    res=insHiphenColor(a_err_no_edit_list, "- " ,
                        bcolors.WARNING)
                    a_err_txt="\n".join(res[2])
                    header = header+f"\n{bcolors.WARNING}Ранее в список были " \
                        f"внесены следующие дефекты:{bcolors.ENDC}\n" \
                        f"{a_err_txt}"

                external_condition = ""
                txt1_1=f"{header}\n{bcolors.OKBLUE}Запишите выявленные замечания к ПУ № "+ \
                    f"{meter_tech_number}:{bcolors.ENDC}"
                res = inputResultExtInspection(header=header, only_defects="1", 
                    defect_auto_list=err_list, param_mandatory_filter=param_mandatory_filter,
                    param_user_filter=param_user_filter, workmode=workmode,
                    mode="новый")
                external_condition=res[1]
                if res[0]=="9":
                    print(f"{bcolors.WARNING}\nВвод данных прерван.{bcolors.ENDC}")
                    continue
                
                if external_condition!="":
                    a_err_list=external_condition.split("\n")
                    err_all_list.extend(a_err_list)
                return ["3", "Замечания успешно внесли в список", err_all_list, ""]
            
            elif oo=="проверка":
                return ["1", "Продолжить проверку.", err_all_list, ""]
            
            elif oo=="проверка без замечания":
                return ["5", "Продолжить проверку без замечания.", err_all_list, ""]

            elif oo=="пропустить тест":
                return ["6", "Пропустить данный тест.", err_all_list, ""]
            
            elif "#4-" in oo:
                oo=oo.replace("#4-","")
                return ["4", "Дополнительный пункт из меню.", err_all_list, oo]
    

    
    def innerSetValReestr():

        
        sutp_transm=""
        id_rec=""
        employee_ip_adr=getLocalIP()

        reestr_clipboard_err_txt=",".join(clipboard_err_list)
    
        rep_remark_txt=", ".join(rep_remark_list)
        delta_pc_minus_device_txt=str(delta_pc_minus_device)

        reestr_key_list=["id_rec","test_start_time", "meter_phase", "meter_type",
            "meter_tech_number", "meter_serial_number","employees_name",
            "meter_color_body", "meter_date_of_manufacture", 
            "meter_soft", "delta_pc_minus_device_txt", "otnoshenie_str",
            "gsm_product_type","gsm_serial_number", "gsm_soft", 
            "rc_serial_number", "rc_soft", 
            "meter_grade", "duration_test", "meter_pw_high_encrypt", 
            "meter_pw_low_encrypt",  "meter_soft_sutp", 
            "meter_model_sutp", "gsm_docked_tn_sutp", "gsm_docked_sn_sutp",
            "gsm_docked_model_sutp", "gsm_docked_soft_sutp",
            "order_num", "order_descript", "sutp_transm",
            "electrical_test_circuit", "ctrl_current_electr_test",
            "reestr_clipboard_err_txt", "rep_remark_txt", "employee_ip_adr",
            "meter_voltage_str", "meter_amperage_str", "meter_config_check",
            "meter_config_param_filename", "energy_consumed", "energy_export",
            "meter_config_disconnect_control_mode","disconnect_control_mode",
            "meter_model_ep", "meter_config_filename", "order_ev"]

        reestr_val_list=[id_rec,test_start_time, meter_phase, meter_type,
            meter_tech_number, meter_serial_number, employees_name,
            meter_color_body, meter_date_of_manufacture, 
            meter_soft, delta_pc_minus_device_txt, otnoshenie_str,
            gsm_product_type, gsm_serial_number, gsm_soft, 
            rc_serial_number, rc_soft, 
            meter_grade, duration_test, meter_pw_high_encrypt, 
            meter_pw_low_encrypt,  meter_soft_sutp, 
            meter_model_sutp, gsm_docked_tn_sutp, gsm_docked_sn_sutp,
            gsm_docked_model_sutp, gsm_docked_soft_sutp, 
            order_num, order_descript, sutp_transm,
            electrical_test_circuit, ctrl_current_electr_test,
            reestr_clipboard_err_txt, rep_remark_txt, employee_ip_adr,
            meter_voltage_str, meter_amperage_str, meter_config_check,
            meter_config_param_filename, energy_consumed, energy_export,
            meter_config_disconnect_control_mode, disconnect_control_mode,
            meter_model_ep, meter_config_filename, order_ev]

        reestr_dic=dict.fromkeys(reestr_key_list, "")
        i=0
        for key in reestr_key_list:
            reestr_dic[key]=reestr_val_list[i]
            i+=1

        return reestr_dic
    
    
   
    def innerTestGSMModem(test_num:str):

        global gsm_serial_number, gsm_SIM_number
        global gsm_product_type
        global modem_type_def
        global default_value_dict, default_filename_full
        global employees_name
        global meter_serial_number, meter_tech_number
        global test_result,test_result_buf,gsm_soft, meter_soft
        global mc_checkVersPO_set, mc_on_board
        global modem_status
        global data_exchange_sutp
        global rep_err_list, clipboard_err_list, rep_remark_list
        global order_descript
        global order_ev

        nonlocal gsm_docked_tn_sutp
        nonlocal gsm_docked_sn_sutp
        nonlocal gsm_docked_model_sutp
        nonlocal gsm_docked_soft_sutp
        nonlocal meter_mc_model_list
        nonlocal mc_test_ok

        mc_order_ev_lbl=""
        
        meter_mc_model_list=[]
        
        txt = "Проверка модуля связи:"
        if test_num!="":
            txt =f"{test_num}. {txt}"
        print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")
        fileWriter(default_filename_full, "a", "", txt+ "\n", \
            "Сохранение в отчет заголовка теста модема",join="on")
    
        
        if modem_status=="0":   
            txt1=f"{bcolors.WARNING}    Модуль связи отсутствует.{bcolors.ENDC}"
            print(txt1)
            if test_num!="":
                fileWriter(default_filename_full, "a", "", txt1+ "\n", \
                    "Сохранение в отчет результата теста модема","on",)
            return ["5", "Отсутствует МС.", {}]
        
        if modem_status=="2":   
            txt1=f"{bcolors.WARNING}    Установлен тестовый МС.{bcolors.ENDC}"
            print(txt1)
            if test_num!="":
                fileWriter(default_filename_full, "a", "", txt1+ "\n", \
                    "Сохранение в отчет информацию о тестовом МС","on",)

        if data_exchange_sutp!="0" and \
            (mc_on_board_set in ["возможно", "обязательно"]) and \
            (gsm_docked_tn_sutp in [None, ""] or \
            gsm_docked_sn_sutp in [None, ""] or \
            gsm_docked_model_sutp in [None, ""] \
            or gsm_docked_soft_sutp in [None, ""]):
            res=getInfoAboutDockedMC(meter_tech_number, workmode)
            if res[0] == "1":
                gsm_docked_tn_sutp=res[2]
                gsm_docked_sn_sutp=res[3]
                gsm_docked_model_sutp=res[4]
                gsm_docked_soft_sutp=res[5]
        

        ret="1"
        ret_txt="Тест МС успешно пройден."
        

        sheet_name = "Product1"

        gsm_product_type=""
        a_sn=gsm_serial_number
        if modem_status=="1" and gsm_serial_number=="" and \
            (not gsm_docked_sn_sutp in [None, ""]):
            a_sn=gsm_docked_sn_sutp
        if a_sn!=None and a_sn!="":
            res=toGetProductInfo2(a_sn, sheet_name)
            if res[0]=="0":
                return ["0", "Ошибка при получении информации о МС из " \
                        "ф.ProductNumber.xlsx.", {}]
            gsm_product_type=res[2]
            mc_order_ev_lbl=res[34]


        res=toGetProductInfo2(meter_serial_number, sheet_name)
        if res[0]=="0":
            return ["0", "Ошибка при получении информации о МС из " \
                    "ф.ProductNumber.xlsx.", {}]
        mc_num_set=res[19]
        mc_checkVersPO_set=res[22]
        mc_checkMS_set=res[23]
        mc_visIndicat_set=res[24]
        a_mc=res[26]



        if a_mc!=None and a_mc!="":
            meter_mc_model_list=a_mc.split(",")

        if gsm_soft==None or gsm_soft=="":
            res = readMCSoft()
            if res[0] == "1":
                gsm_soft = res[2]
            else: 
                return ["4", "Нет связи с ПУ", {}]

        
        txt1_rep=""


        a_color=bcolors.OKGREEN
        if gsm_docked_tn_sutp in [None, ""]:
            a_color=bcolors.WARNING
            rep_remark_list.append("В СУТП отсутствует информация о "
                "состыкованном МС.")
        txt1=f"{bcolors.OKGREEN}Номера МС:{bcolors.ENDC}\n" \
            f"      {bcolors.OKGREEN}Технический номер состыкованного МС (из СУТП): " \
            f"{bcolors.ENDC}{a_color}{gsm_docked_tn_sutp}{bcolors.ENDC}"

        if test_num!="":
            txt1 =f"{bcolors.OKGREEN}{test_num}.1. {bcolors.ENDC}{txt1}"
            fileWriter(default_filename_full, "a", "", f"{txt1}\n", \
                "Сохранение в отчет результата теста модема", join="on")
        
        print(txt1)

        if mc_num_set=="нет" and gsm_serial_number=="":
            txt1_rep="      Серийный номер МС на крышке: не предусмотрен"
            txt1=f"{bcolors.OKGREEN}{txt1_rep}{bcolors.ENDC}"
        
        else:
            txt1_rep=f"      Серийный номер МС на крышке (QR-код): {gsm_serial_number}"
            txt1=f"{bcolors.OKGREEN}      Серийный номер МС на крышке (QR-код): " \
                f"{gsm_serial_number}{bcolors.ENDC}"
            
        print(txt1)
        if test_num!="":
            fileWriter(default_filename_full, "a", "", f"{txt1_rep}\n", \
                "Сохранение в отчет результата теста модема",join="on")
        
        
        a_color_sn_sutp=bcolors.OKGREEN
        txt1_1=""

        if mc_num_set=="да" and gsm_docked_sn_sutp!=None \
            and gsm_docked_sn_sutp!="" and gsm_docked_sn_sutp!=gsm_serial_number:
            a_color_sn_sutp=bcolors.WARNING
            txt1_2=f"Серийный номер МС в СУТП ({gsm_docked_sn_sutp}) " \
                "отличается от номера, указанного на крышке или в QR-коде МС " \
                f"({gsm_serial_number})."
            txt1_1="\n      "+txt1_2
            if modem_status=="1":
                rep_err_list.append(txt1_2)
                clipboard_err_list.append(txt1_2)
            
            elif modem_status=="2":
                rep_remark_list.append(txt1_2)
        
        if gsm_docked_sn_sutp in [None, ""]:
            a_err_txt="В СУТП отсутствует серийный номер " \
                "состыкованного МС."
            a_color_sn_sutp=bcolors.WARNING
            rep_remark_list.append(a_err_txt)

        txt1_rep=f"      Серийный номер состыкованного МС (из СУТП): " \
        f"{gsm_docked_sn_sutp}"+txt1_1
        txt1=f"{bcolors.OKGREEN}      Серийный номер состыкованного МС (из СУТП): " \
            f"{a_color_sn_sutp}{gsm_docked_sn_sutp}{bcolors.ENDC}" \
            f"{bcolors.WARNING}{txt1_1}{bcolors.ENDC}"
        
        print(txt1)
        if test_num!="":
            fileWriter(default_filename_full, "a", "", f"{txt1_rep}\n", \
                "Сохранение в отчет результата теста модема",join="on")
        
        txt=""
        
        txt1=""
        SIM_txt_dict={"0":[f"{bcolors.WARNING}","SIM карта отсутствует."],
            "2":[f"{bcolors.WARNING}","Установлена тестовая SIM карта."],
            "3-0":[f"{bcolors.WARNING}","Номер SIM карты не ввели."],
            "3-1":[f"{bcolors.OKGREEN}",f"Номер SIM-карты: {gsm_SIM_number}"]}
        a=SIMcard_status
        if a=="3" and gsm_SIM_number=="0":
            a="3-0"
        elif a=="3" and gsm_SIM_number!="0":
            a="3-1"
        
        txt1=SIM_txt_dict.get(a)[1]
        a_color=SIM_txt_dict.get(a)[0]
        if test_num!="":
            txt1 =f"{test_num}.2. {txt1}"
        print(f"{a_color}{txt1}{bcolors.ENDC}")
        txt = txt1
        if test_num!="":
            fileWriter(default_filename_full, "a", "", f"{txt}\n", \
                "Сохранение в отчет результата теста модема",join="on")
    
        
        txt1_rep=""
        a_color_model_sutp=bcolors.OKGREEN
        a_color_capt=bcolors.OKGREEN
        a_model_caption=" "+gsm_product_type+" (значение по умолчанию)"

        if order_ev=="1":
            a_model_caption=" "+mc_order_ev_lbl+" (значение по умолчанию)"
            
        txt1_1=""

        txt1_rep=f"Модель МС:\n"
        txt1=f"{bcolors.OKGREEN}{txt1_rep}{bcolors.ENDC}"

        if mc_num_set=="нет" and gsm_product_type==None:
            txt1_rep=txt1_rep+f"      - определенная по серийному " \
                f"номеру МС на крышке: не определялась"
            txt1=f"{bcolors.OKGREEN}{txt1_rep}{bcolors.ENDC}"
 
        elif mc_num_set=="нет" and gsm_product_type!="":
            txt1_rep=txt1_rep+f"      - определенная по серийному " \
                f"номеру МС из СУТП: {gsm_product_type}"
            txt1=f"{bcolors.OKGREEN}{txt1_rep}{bcolors.ENDC}"
        
        elif mc_num_set=="да" and gsm_product_type!="":
            if modem_type_def == "спрашивать каждый раз" or gsm_product_type != modem_type_def:
                a_txt=f"      На крышке МС указана модель '{gsm_product_type}'. " \
                    "Верно? 0- нет, 1- да"
                
                if order_ev=="1":
                    a_txt="      ПУ включен в заказ розничной продажи. " \
                        f"На крышке МС указана модель '{mc_order_ev_lbl}'. " \
                        "Верно? 0- нет, 1- да"

                oo = questionSpecifiedKey(bcolors.OKBLUE, a_txt, ["0","1"] , "", 1)
                print()
                if oo=="0":
                    a_err_txt="Модель на крышке МС отличается от модели, " \
                        f"определенной по серийному номеру МС ({gsm_product_type})."
                    a_model_caption=" отличается от модели, определенной по серийному номеру."
                    a_color_capt=bcolors.FAIL
                    rep_err_list.append(a_err_txt)
                    clipboard_err_list.append(a_err_txt)

                else:
                    a_model_caption=" "+gsm_product_type
                    if order_ev=="1":
                        a_model_caption=" "+mc_order_ev_lbl
                        
                    if modem_type_def!="спрашивать каждый раз":
                        a_txt=f"Заменить модель МС по умолчанию '{modem_type_def}' " \
                            f"на новое значение '{gsm_product_type}'? 0-нет, 1-да"
                        oo=questionSpecifiedKey(bcolors.OKBLUE, a_txt, ["0","1"], "", 1)
                        print()
                        if oo=="1":
                            modem_type_def=gsm_product_type
                            a_save_ok=True
                            default_value_dict = writeDefaultValue(default_value_dict)
                            res=saveConfigValue('opto_run.json', default_value_dict)
                            if res[0]=="0":
                                printWARNING(res[1])
                                a_save_ok=False

                            a_dic={"modem_type_def": modem_type_def}
                            res=saveConfigValue('opto_run.json_last', a_dic)
                            if res[0]=="1" and a_save_ok:
                                print(f"{bcolors.OKGREEN}Установлено новое значение по " 
                                f"умолчанию для модели ПУ: {modem_type_def}.")
                            else:
                                printWARNING(res[1])


            elif order_ev=="1" and mc_order_ev_lbl!=modem_type_def:
                a_txt="Проверьте, что на МС указана модель " \
                    f"{bcolors.ATTENTIONWARNING} '{mc_order_ev_lbl}' {bcolors.ENDC}."

            txt1_rep=txt1_rep+f"      - определенная по серийному " \
                f"номеру МС на крышке: {gsm_product_type}\n" \
                f"      - указанная на крышке МС: {a_model_caption}\n"
            txt1=txt1+f"{bcolors.OKGREEN}      - определенная по серийному " \
                f"номеру МС на крышке: {gsm_product_type}\n" \
                f"      - указанная на крышке МС:{bcolors.ENDC} " \
                f"{a_color_capt}{a_model_caption}{bcolors.ENDC}\n"

            txt1_1=""

            if gsm_docked_model_sutp!=None and gsm_docked_model_sutp!="" \
                and gsm_docked_model_sutp!=gsm_product_type:
                a_color_model_sutp=bcolors.WARNING
                txt1_2=f"Модель МС в СУТП ({gsm_docked_model_sutp}) " \
                    "отличается от модели, определенной по серийному " \
                    f"номеру МС ({gsm_product_type})."
                txt1_1="\n      "+txt1_2
                rep_err_list.append(txt1_2)
                clipboard_err_list.append(txt1_2)

            if gsm_docked_model_sutp in [None, ""]:
                a_color_model_sutp=bcolors.WARNING
                txt1_2="В СУТП отсутствует информация о модели " \
                    "состыкованного МС."
                rep_remark_list.append(txt1_2)

            txt1_rep=txt1_rep+f"      - по данным из СУТП: {gsm_docked_model_sutp}"+txt1_1
            txt1=txt1+f"{bcolors.OKGREEN}      - по данным из СУТП:{bcolors.ENDC} " \
                f"{a_color_model_sutp}{gsm_docked_model_sutp}{bcolors.ENDC}" \
                f"{bcolors.WARNING}{txt1_1}{bcolors.ENDC}"

        if gsm_product_type!=None and gsm_product_type!="":
            if len(meter_mc_model_list)==0:
                printFAIL("Список допустимых моделей МС пуст.")
                return["0", "Ошибка: список допустимых моделей МС пуст.", {}]
            
            a_txt="Модель имеется в списке допустимых моделей для " \
                f"данного ПУ: {', '.join(meter_mc_model_list)}."
            a_color=bcolors.OKGREEN
            if mc_on_board_set=="нет" and modem_status=="1" and mc_on_board=="1":
                a_txt="Для данной модели ПУ не предусмотрено применение МС."
                a_color=bcolors.FAIL
                rep_err_list.append(a_txt)
                clipboard_err_list.append(a_txt)
                
            elif not gsm_product_type in meter_mc_model_list:
                a_txt=f"Модель '{gsm_product_type}' отсутствует в списке допустимых " \
                    f"моделей для данного ПУ: {', '.join(meter_mc_model_list)}."
                a_color=bcolors.FAIL
                rep_err_list.append(a_txt)
                clipboard_err_list.append(a_txt)
            
            txt1_rep=txt1_rep+f"\n      {a_txt}"
            txt1=txt1+f"\n      {a_color}{a_txt}{bcolors.ENDC}"
        
        if test_num!="":
            txt1_rep =f"{test_num}.3. {txt1_rep}"
            txt1 =f"{bcolors.OKGREEN}{test_num}.3.{bcolors.ENDC} {txt1}"
            fileWriter(default_filename_full, "a", "", f"{txt1_rep}\n", \
                "Сохранение в отчет результата теста модема",join="on")
        
        print(txt1)

        txt=""
            
        txt1_rep = f"Версия ПО модуля связи:"
        if test_num!="":
            txt1_rep =f"{test_num}.4. {txt1_rep}"
        print(f"{bcolors.OKGREEN}{txt1_rep}{bcolors.ENDC}")

        if mc_checkVersPO_set=="нет":
            txt1="      Не требуется проверять."
            printGREEN(txt1)
            txt1="Не требуется проверять."

        if mc_checkVersPO_set=="да" and (gsm_soft==None or gsm_soft==""):
            txt1=""
            while True:
                res = readMCSoft()
                if res[0] == "1":
                    gsm_soft = res[2]
                    if gsm_soft!="":
                        break

                if res[0]!="1" or gsm_soft=="":
                    txt2=f"\n{bcolors.FAIL}ПУ не смог связаться с модулем связи.{bcolors.ENDC}\n" \
                        f"{bcolors.WARNING}Возможные причины:{bcolors.ENDC}\n" \
                        f"{bcolors.WARNING}- нет контакта между ПУ и МС{bcolors.ENDC}\n" \
                        f"{bcolors.WARNING}- запуск (перезагрузка) ПО МС{bcolors.ENDC}\n" \
                        f"{bcolors.WARNING}- ПО МС 'зависло'{bcolors.ENDC}"
                    a_err_list=["ПУ не смог связаться с модулем связи."]
                    select_mode="1"

                    if not mc_test_ok:
                        select_mode="3"

                    a_menu_add_item_list=["Повторный опрос МС", "Пропустить тест МС"]
                    a_menu_add_id_list=["повтор", "пропустить"]

                    res=innerSelectActions(txt2, a_err_list, select_mode, "МС", a_menu_add_item_list, 
                        a_menu_add_id_list, rep_err_list)
                    if res[0]=="9":
                        a_txt="Проверка ПУ прервана пользователем."
                        if test_num!="":
                            testBreak("Проверка МС", a_txt, default_filename_full, employees_name)
                        return ["9", a_txt, {}]
                    
                    elif res[0]=="2":
                        return ["2", "ПУ отправлен в ремонт."]
                    
                    elif res[0]=="4" and res[3]=="повтор":
                        print ("Повторно опрашиваем ПУ на наличие связи с МС.")
                        continue

                    elif res[0]=="1":
                        rep_err_list.extend(a_err_list)
                        clipboard_err_list.extend(a_err_list)
                        print (f"{bcolors.WARNING}Замечание '{a_err_list[0]}' " 
                               f"добавлено в список с дефектами.{bcolors.ENDC}")
                        txt1=a_err_list[0]

                        mc_test_ok=False

                        break
                
                    elif res[0]=="4" and res[3]=="пропустить":
                        ret_txt="Тест МС пропущен."
                        print(f"      {bcolors.FAIL}{ret_txt}{bcolors.ENDC}")
                        rep_remark_list.append(ret_txt)
                        if modem_status=="1":
                            rep_err_list.append(ret_txt)

                        txt1_rep=txt1_rep+"тест пропущен"
                        if test_num!="":
                            fileWriter(default_filename_full, "a", "", f"{txt1_rep}\n", \
                                "Сохранение в отчет результата теста модема",join="on")
                        return ["8", "Тест МС пропущен.", {}]
                    

        if mc_checkVersPO_set=="да" and (gsm_soft!=None and gsm_soft!=""):
            txt1 = f"      Из ПУ получено значение: {gsm_soft}."
            print(f"{bcolors.OKGREEN}{txt1}{bcolors.ENDC}")
            txt1_rep=txt1_rep+"\n"+txt1
            
            txt1=""
            res=checkVersMC(meter_soft=meter_soft,mc_soft=gsm_soft,
                            device_ver_list=actual_device_version_list,
                            print_msg="0")
            if res[0]=="1":
                txt1="Версия соответствует актуальной версии."
                print(f"      {bcolors.OKGREEN}{txt1}{bcolors.ENDC}")

            else:
                if res[0]=="3":
                    txt1="Для данного типа ПУ не удалось подобрать актуальную " \
                        "версию ПО модуля связи."
                    print(f"      {bcolors.WARNING}{txt1}{bcolors.ENDC}")

                    return ["0", txt1]

                elif res[0]=="2":
                    mc_ver_list_txt=res[2]
                    txt1=f"Версия ПО МС не соответствует актуальной версии ПО " \
                        f"({mc_ver_list_txt}). Необходимо обновление ПО."
                    print(f"      {bcolors.FAIL}{txt1}{bcolors.ENDC}")
                
                rep_err_list.append(txt1)
                clipboard_err_list.append(txt1)

                ret_txt="Тест не пройден."
                ret="3" 

        txt1_rep=txt1_rep+"\n      "+txt1

        txt1=f"      Версия ПО МС, указанная в СУТП: {gsm_docked_soft_sutp}"
        a_color=bcolors.OKGREEN
        txt1_1=""
        if mc_checkVersPO_set=="да" and (not gsm_docked_soft_sutp in [None, ""]) and \
            gsm_soft!="" and gsm_docked_soft_sutp!=gsm_soft:
            txt1_2=f"Версия ПО МС, указанная в СУТП ({gsm_docked_soft_sutp}) " \
                f"отличается от версии в МС ({gsm_soft})."
            txt1_1="\n      "+txt1_2
            
            rep_err_list.append(txt1_2)
            clipboard_err_list.append(txt1_2)
            a_color=bcolors.WARNING
        
        if gsm_docked_soft_sutp  in [None, ""]:
            a_color=bcolors.WARNING
            a_txt="В СУТП отсутствует информация о версии ПО МС."
            rep_remark_list.append(a_txt)

        print(f"{bcolors.OKGREEN}      Версия ПО МС, указанная в СУТП:{bcolors.ENDC} " 
            f"{a_color}{gsm_docked_soft_sutp}{bcolors.ENDC}"
            f"{a_color}{txt1_1}{bcolors.ENDC}")
        txt1_rep=txt1_rep+"\n"+txt1+txt1_1

        if test_num!="":
            fileWriter(default_filename_full, "a", "", f"{txt1_rep}\n", \
                "Сохранение в отчет результата теста модема",join="on")
        
        
        txt=""
            
        txt1 = "Исправность МС: "
        if test_num!="":
            txt1 =f"{test_num}.5. {txt1}"
        print(f"{bcolors.OKGREEN}{txt1}{bcolors.ENDC}")
        txt=txt1

        txt1=""
        if mc_checkMS_set!="нет" and mc_test_ok:
            while True:
                if mc_checkMS_set=="виз":
                    txt2=f"{bcolors.OKBLUE}Проверьте состояние МС согласно методики:\n" \
                        f"{bcolors.WARNING}{mc_visIndicat_set}\n" \
                        f"{bcolors.OKBLUE}По результату проверки МС исправно?" \
                        " 0-нет, 1-да, 8-пропустить тест, /-прервать проверку."

                elif "GSM" in mc_checkMS_set:
                    try:
                        signal_level = GXDLMSRegister("0.0.99.13.164.255")
                        signal_level = reader.read(signal_level, 2)
                    except Exception as e:
                        oo = communicationTimoutError("Чтение уровня GSM сигнала:", e.args[0])
                        if oo=="0" or oo == "-1":
                            return ["4", "Нет связи с ПУ.", {}]
            
                    if signal_level == 99:
                        print()

                        if SIMcard_status!="0":
                            txt2 = f"{bcolors.WARNING}Нет информации об уровне GSM-сигнала " \
                                f"от МС.{bcolors.ENDC}\n" \
                                f"{bcolors.OKBLUE}Дождитесь загорания соответствующего светодиода и " \
                                f"нажмите 1.{bcolors.ENDC}\n" \
                                f"{bcolors.OKBLUE}Чтобы записать замечание о неисправности МС " \
                                f"- нажмите 0.{bcolors.ENDC}\n" \
                                f"{bcolors.OKBLUE}Чтобы пропустить тест - нажмите 8.{bcolors.ENDC}\n" \
                                f"{bcolors.OKBLUE}Чтобы прекратить проверку - нажмите /.{bcolors.ENDC}\n"
                    
                        else:
                            txt2 = f"{bcolors.WARNING}SIM-карта отсутствует.{bcolors.ENDC}\n" \
                                f"{bcolors.OKBLUE}Дождитесь загорания светодиода 'POWER' и " \
                                f"нажмите 1.{bcolors.ENDC}\n" \
                                f"{bcolors.OKBLUE}Чтобы записать замечание о неисправности МС " \
                                f"- нажмите 0.{bcolors.ENDC}\n" \
                                f"{bcolors.OKBLUE}Чтобы пропустить тест - нажмите 8.{bcolors.ENDC}\n" \
                                f"{bcolors.OKBLUE}Чтобы прекратить проверку - нажмите /.{bcolors.ENDC}\n"
                
                            if mc_visIndicat_set!="" and mc_visIndicat_set!=None:    
                                txt2=f"{bcolors.WARNING}SIM-карта отсутствует.{bcolors.ENDC}\n" \
                                    f"{bcolors.OKBLUE}Проверьте состояние МС согласно методики:\n" \
                                    f"{mc_visIndicat_set}\nПо результату проверки МС исправно?" \
                                    " 0-нет, 1-да, 8-пропустить тест, /-прервать проверку."
                    
                    else:
                        txt1=f"Уровень сигнала составляет {str(signal_level)} ед."
                        print(f"      {bcolors.OKGREEN}{txt1}{bcolors.ENDC}")
                        break
                
                specified_keys_list=["0", "1", "8", "/", "", 1]

                oo=questionSpecifiedKey("",txt2, specified_keys_list, "", 1)
                if oo == "8":
                    txt1="Тест пропущен"
                    print(f"\n      {bcolors.WARNING}{txt1}{bcolors.ENDC}")
                    ret_txt="Тест МС пропущен."
                    rep_err_list.append(ret_txt)
                    rep_remark_list.append(ret_txt)
                    
                    return ["8", ret_txt]
                
                elif oo=="/":
                    txt1_1="Тест модема"
                    testBreak(txt1_1,"Проверка прервана пользователем",default_filename_full, 
                        employees_name)
                    return ["9", "Проверка ПУ прервана пользователем.", {}]
                
                elif oo=="0":
                    txt2=f"\n{bcolors.OKBLUE}Опишите состояние МС (его индикации), " \
                        "по которой Вы приняли решение о его неисправности.\n" \
                        "По окончании нажмите Enter.\n" \
                        "Чтобы вернуться к тесту работоспособности МС - нажмите 1.\n" \
                        "Чтобы пропустить тест - нажмите 8.\n" \
                        f"Чтобы прекратить проверку - нажмите /.{bcolors.ENDC}"
                    spec_keys=["1","8","/"]
                    oo=inputSpecifiedKey("",txt2,"",[0],spec_keys,0)
                    if oo=="8":
                        txt1="тест пропущен"
                        print(f"      {bcolors.WARNING}{txt1}{bcolors.ENDC}")
                        ret_txt="Тест МС пропущен."
                        rep_err_list.append(ret_txt)
                        rep_remark_list.append(ret_txt)
                        ret="8"
                        return ["8", ret_txt]
                    
                    elif oo=="9":
                        return ["9", "Проверка ПУ прервана пользователем.", {}]
                    elif oo=="1":
                        continue
                    
                    txt1=f"SIM-карта отсутствует."
                    if SIMcard_status!="0":
                        txt1="Нет информации об уровне GSM-сигнала от МС."

                    txt1=txt1+ " Модуль связи неисправен по следующим признакам:"
                    print(f"      {bcolors.FAIL}{txt1}{bcolors.ENDC}\n"
                        f"      {bcolors.FAIL}{oo}{bcolors.ENDC}")
                    txt1=f"{txt1}\n{oo}"
                    rep_err_list.append(txt1)
                    clipboard_err_list.append(txt1)
                    ret_txt="Тест не пройден"
                    ret="3"
                    break
                
                else:
                    if SIMcard_status!="0":
                        continue
                    txt1="SIM-карта отсутствует. Согласно индикации светодиодов - модуль связи исправен."
                    print (f"\n      {bcolors.OKGREEN}{txt1}{bcolors.ENDC}")
                    break
            
        elif mc_checkMS_set=="нет":
            txt1="Не требуется проверять."
            print(f"      {bcolors.OKGREEN}{txt1}{bcolors.ENDC}")

        else:
            txt1="Проверка не проводилась."
            print(f"      {bcolors.WARNING}{txt1}{bcolors.ENDC}")

        
        txt = txt+"\n      "+txt1
        if test_num!="":
            fileWriter(default_filename_full, "a", "", f"{txt}\n", \
                "Сохранение в отчет результата теста модема",join="on")
        txt=""
        
        ret_dic={"gsm_docked_tn_sutp": gsm_docked_tn_sutp,
                "gsm_docked_sn_sutp": gsm_docked_sn_sutp,
                "gsm_docked_model_sutp": gsm_docked_model_sutp,
                "gsm_docked_soft_sutp": gsm_docked_soft_sutp,
                "gsm_product_type": gsm_docked_model_sutp}
        
        return [ret, ret_txt, ret_dic]

   

    def innerFirstCheckMC():

        global gsm_soft
        global mc_on_board
        global modem_status
        global default_value_dict
        global SIMcard_status
        global gsm_serial_number
        global gsm_SIM_number
        global gsm_product_type
        global rep_remark_list
        global rep_err_list
        global clipboard_err_list
        global connection_initialized

        nonlocal gsm_docked_tn_sutp
        nonlocal gsm_docked_sn_sutp
        nonlocal gsm_docked_model_sutp
        nonlocal gsm_docked_soft_sutp
        nonlocal mc_test_ok



        if (gsm_soft!=None and gsm_soft!=""):
            mc_on_board="1"

            if mc_on_board_set=="нет" and modem_status=="1":
                a_err_txt="Обнаружен модуль связи. Для данного ПУ не " \
                    "предусмотрена установка МС."
                a_err_msg=f"{bcolors.WARNING}{a_err_txt}{bcolors.ENDC}"
                a_err_list=rep_err_list.copy()
                a_err_list.append(a_err_txt)
                menu_item_add_list=["Исключить МС из проверки ПУ",
                    "Изменить статус МС на 'будет тестовым'"]
                menu_id_add_list=["исключить", "статус тестовый"]
                res=innerSelectActions(a_err_msg, a_err_list, "3", "МС", 
                    menu_item_add_list, menu_id_add_list)
                if res[0]=="9":
                    txt1_2 = f"Проверка прервана пользователем."
                    testBreak("", txt1_2, default_filename_full, employees_name)
                    return ["9", txt1_2]
                
                elif res[0]=="2":
                    return ["2", "ПУ отправлен в ремонт."]

                elif res[0]=="4" and res[3]=="исключить":
                    modem_status="0"
                    default_value_dict = writeDefaultValue(default_value_dict)
                    saveConfigValue('opto_run.json',default_value_dict)
                    print(f"{bcolors.WARNING}Тест МС будет пропущен.{bcolors.ENDC}")

                elif res[0]=="4" and res[3]=="статус тестовый":
                    modem_status="2"
                    default_value_dict = writeDefaultValue(default_value_dict)
                    saveConfigValue('opto_run.json',default_value_dict)
                    print(f"{bcolors.WARNING}Статус МС изменен на 'будет тестовый'." 
                        f"{bcolors.ENDC}")

            if modem_status=="0" and (mc_on_board_set in ["возможно", "обязательно"]):
                txt2="Обнаружен модуль связи. Укажите его статус"
                list_txt=["исключить из проверки","рабочий", "тестовый"]
                list_id = ["0", "1", "2"]
                oo = questionFromList(bcolors.WARNING, txt2, list_txt, list_id)
                print()
                if oo=="2":
                    txt1_1="\nВы уверены, что МС тестовый? 1-да, МС тестовый; 2-нет, МС рабочий; " \
                        "3-нет, МС нужно исключить из проверки"
                    spec_keys=["1","2","3"]
                    oo = questionSpecifiedKey(bcolors.WARNING,txt1_1,spec_keys,"",1)
                    print()
                    modem_status_dict={"1":"2","2":"1","3":"0"}
                    oo=modem_status_dict[oo]
                if modem_status!=oo:
                    modem_status=oo
                    default_value_dict = writeDefaultValue(default_value_dict)
                    saveConfigValue('opto_run.json',default_value_dict)
                    a_list = ["устанавливаться не будет", "будет рабочим", 
                        "будет тестовым"]
                    print (f'{bcolors.WARNING}Статус МС изменен на ' \
                        f'"{a_list[int(modem_status)]}".{bcolors.ENDC}')
                
                if modem_status == "1":
                    list_txt=["отсутствует", "рабочая, запросить номер карты", 
                                "рабочая, номер карты не нужен", "тестовая"]
                    list_id=["0","1","3","2"]
                    txt1="\nВведите информацию о статусе SIM-карты:"
                    oo = questionFromList(bcolors.OKBLUE, txt1, list_txt, list_id,
                        SIMcard_status,[],[],[])
                    SIMcard_status = oo
                    default_value_dict = writeDefaultValue(default_value_dict)
                    saveConfigValue('opto_run.json',default_value_dict)
                    SIMcard_status_dict={"0":"SIM-карта отсутствует","2":"SIM-карта тестовая", \
                        "3":"номер SIM-карты не нужен"}
                    txt1_1=SIMcard_status_dict[SIMcard_status]
                    print(f"\n{bcolors.OKGREEN}{txt1_1}{bcolors.ENDC}")
                
        elif modem_status=="1" and gsm_soft=="" \
            and mc_checkVersPO_set=="да":
            while True:
                res=innerToInitConnectOpto(com_opto)
                if res=="0":
                    return ["0", "Ошибка при инициализации канала связи."]
                
                elif res=="9":
                    txt1_2 = f"Проверка прервана пользователем."
                    return ["9", txt1_2]
                
                res = readMCSoft()
                toCloseConnectOpto()
                connection_initialized = False

                if res[0] == "1":
                    gsm_soft = res[2]
                    if gsm_soft!=None and gsm_soft!="":
                        mc_on_board="1"
                        break

                a_txt=f"ПУ не смог связаться с модулем связи."
                a_rep_err_list=rep_err_list.copy()
                a_rep_err_list.append(a_txt)
                txt2=f"{bcolors.FAIL}{a_txt}{bcolors.ENDC}"
                a_menu_item_add_list=["Повторно прочитать версию ПО МС из ПУ",
                    "МС установлен. Версию ПО МС прочитаем позднее.", 
                    "Исключить МС из проверки ПУ"]
                a_menu_id_add_list=["повторить", "игнорировать", "исключить"]
                res=innerSelectActions(txt2, a_rep_err_list, "3",
                    "МС", a_menu_item_add_list, a_menu_id_add_list)
                if res[0]=="9":
                    txt1_2 = f"Проверка прервана пользователем."
                    return ["9", txt1_2]

                elif res[0]=="2":
                    return ["2", "ПУ отправлен в ремонт."]

                elif res[0]=="4" and res[3]=="исключить":
                    modem_status="0"
                    default_value_dict = writeDefaultValue(default_value_dict)
                    saveConfigValue('opto_run.json',default_value_dict)
                    print(f"{bcolors.WARNING}Проверка МС будет пропущена.{bcolors.ENDC}")
                    break

                elif res[0]=="4" and res[3]=="игнорировать":
                    mc_on_board="1"
                    break
            
        elif gsm_soft=="" and (mc_on_board_set in ["обязательно", "возможно"]) and \
            mc_checkVersPO_set=="нет":
            txt2=f"{bcolors.OKBLUE}Модуль связи установлен? 0-нет, 1-да, " \
                f"/-прервать проверку.{bcolors.ENDC}"
            spec_keys=["0","1","/"]
            oo = questionSpecifiedKey("", txt2,spec_keys,"",1)
            print()
            if oo=="0":
                if mc_on_board_set=="возможно":
                    modem_status="0"
                    default_value_dict = writeDefaultValue(default_value_dict)
                    saveConfigValue('opto_run.json',default_value_dict)
                    print(f"{bcolors.WARNING}Тест МС будет пропущена.{bcolors.ENDC}")
                
                elif mc_on_board_set=="обязательно":
                    a_err_txt="Отсутствует модуль связи. Для данного ПУ " \
                        "предусмотрена обязательная установка МС."
                    a_err_msg=f"{bcolors.WARNING}{a_err_txt}{bcolors.ENDC}"
                    a_err_list=rep_err_list.copy()
                    a_err_list.append(a_err_txt)
                    menu_item_add_list=["Исключить МС из проверки ПУ"]
                    menu_id_add_list=["исключить"]
                    res=innerSelectActions(a_err_msg, a_err_list, "3", "МС", 
                        menu_item_add_list, menu_id_add_list)
                    if res[0]=="9":
                        txt1_2 = f"Проверка прервана пользователем."
                        return ["9", txt1_2]
                    
                    elif res[0]=="2":
                        return ["2", "ПУ отправлен в ремонт."]
                    
                    elif res[0]=="4" and res[3]=="исключить":
                        modem_status="0"
                        mc_on_board="0"
                        default_value_dict = writeDefaultValue(default_value_dict)
                        saveConfigValue('opto_run.json',default_value_dict)
                        print(f"{bcolors.WARNING}Тест МС будет пропущен.{bcolors.ENDC}")
            
            elif oo=="1":
                mc_on_board="1"
            
            else:
                txt1_2 = f"Проверка прервана пользователем."
                return ["9", txt1_2]


        if mc_on_board=="1" and modem_status=="0":
            a_txt=f"{bcolors.WARNING}К ПУ подключен МС. При этом в статусе МС " \
                f"указано 'не будет устанавливаться'.{bcolors.ENDC}\n" 
            menu_item_add_list=["Изменить статус МС на 'рабочий'",
                "Изменить статус МС на 'тестовый'"]
            menu_id_add_list=["рабочий", "тестовый"]
            res=innerSelectActions(a_txt, [], "4", "МС", 
                menu_item_add_list, menu_id_add_list)
            if res[0]=="9":
                txt1_2 = f"Проверка прервана пользователем."
                return ["9", txt1_2]
            
            elif  res[0]=="4" and (res[3]=="рабочий" or res[3]=="тестовый"):
                modem_status="1"
                if res[3]=="тестовый":
                    modem_status="2"
                default_value_dict = writeDefaultValue(default_value_dict)
                saveConfigValue('opto_run.json',default_value_dict)
                print(f"{bcolors.WARNING}Статус МС изменен на '{res[3]}'."
                      f"{bcolors.ENDC}")

            else:
                print (f"{bcolors.WARNING}Тест МС будет пропущен.{bcolors.ENDC}")


        if gsm_soft!="" and gsm_soft!=None and modem_status in ["1", "2"] and \
            mc_checkVersPO_set=="да":
            res=checkVersMC(meter_soft=meter_soft,mc_soft=gsm_soft,
                device_ver_list=actual_device_version_list,print_msg="1")
            if res[0]=="2":
                err_txt=res[1]
                rep_err_list.append(err_txt)
                clipboard_err_list.append(err_txt)
                mc_test_ok=False
                header=""
                res=innerSelectActions(header, rep_err_list, "1")
                if res[0]=="9":
                    txt1_2 = f"Проверка прервана пользователем."
                    return ["9", txt1_2]
                
                elif res[0]=="2":
                    return ["2", "ПУ отправлен в ремонт."]
                

            elif res[0]!="1":
                a_txt="При проверке соответствии версий ПО МС и ПО ПУ " \
                    f"возникла ошибка: {res[1]}"
                return ["0", a_txt]
        
        
        if mc_on_board=="1" and modem_status=="2":
            print(f"\n{bcolors.WARNING}Установлен тестовый модем.{bcolors.ENDC}")
        
        
        if data_exchange_sutp!="0" and \
            (mc_on_board_set in ["возможно", "обязательно"]):
            res=getInfoAboutDockedMC(meter_tech_number, workmode)
            if res[0] == "1":
                gsm_docked_tn_sutp=res[2]
                gsm_docked_sn_sutp=res[3]
                gsm_docked_model_sutp=res[4]
                gsm_docked_soft_sutp=res[5]
                print ("Из СУТП получена информация:")
                print ("Технический номер состыкованного МС: "
                    f"{gsm_docked_tn_sutp}")
                print ("Серийный номер состыкованного МС: "
                    f"{gsm_docked_sn_sutp}")
                print (f"Модель состыкованного МС: {gsm_docked_model_sutp}")
                print (f"Версия ПО состыкованного МС: {gsm_docked_soft_sutp}")

        return ["1", "Сверка проведена успешно."]



    def innerGetVarDefAll():

        global meter_type_def
        global modem_type_def

        res=readGonfigValue(file_name_in="opto_run.json_last",
            var_name_list=[],default_value_dict=default_value_dict)
        if res[0]!="1":
            a_err_txt="Ошибка при чтении данных из ф.opto_run.json_last"
            printWARNING(a_err_txt)
            keystrokeEnter()
            return ["0", a_err_txt]
        
        a_dic=res[2]

        meter_type_def=a_dic["meter_type_def"]

        modem_type_def=a_dic["modem_type_def"]

        return ["1", "Значение переменных по умолчанию обновлены."]
    

    
    def innerReadMeterTime():

        while True:
            try:
                data = GXDLMSClock("0.0.1.0.0.255")
                device_date_time = str(reader.read(data, 2))
                device_date_time1 = f"{device_date_time[3:5]}." \
                    f"{device_date_time[0:2]}.{device_date_time[6:8]} " \
                    f"{device_date_time[9:]}"
                break
                
            except Exception as e:
                oo = communicationTimoutError("Считывание текущей даты и времени ПУ: ", e.args[0])
                if oo=="0" or oo == "-1":
                    a_dic={"0": ["Ошибка связи с ПУ при чтении текущей даты и времени ПУ.", "0"],
                        "-1": ["Проверка прервана пользователем.", "9"]}
                    ret_txt=a_dic[oo][0]
                    ret_id=a_dic[oo][1]
                    return [ret_id, ret_txt]
            

        pc_time = datetime.now()
        pc_time1 = str(pc_time)
        pc_time1 = f"{pc_time1[8:10]}.{pc_time1[5:7]}.{pc_time1[2:4]} {pc_time1[11:19]}"

        device_date_time2 = datetime.strptime(device_date_time, "%m/%d/%y %H:%M:%S")
        delta_pc_minus_device = int(abs(pc_time - device_date_time2).seconds)

        return ["1", "Успешно", device_date_time, device_date_time1, pc_time, 
                pc_time1, delta_pc_minus_device]
    
    
    
    def innerChangeMeterTimezone(tz_new: int):

        while True:
            try:
                data = GXDLMSClock("0.0.1.0.0.255")
                data.timeZone = GXTimeZone(tz_new)
                reader.write(data, 3)
                return "1"

            except Exception as e:
                oo = communicationTimoutError("Изменение часового пояса в ПУ: ", e.args[0])
                if oo in ["0","-1"]:
                    return oo


    
    test_go=True #метка для прерывания теста по цифре 9
    cur_test_skip=False #метка, чтобы пропустить выполнение тек. теста по цифре 8

    employees_name=""

    order_num=""
    order_descript=""
    order_ev="0"

    meter_soft=""
    meter_soft_sutp=""
    meter_type=""
    meter_model_sutp=None
    meter_model_ep=None
    meter_tech_number=""
    meter_serial_number=""
    meter_date_of_manufacture=""
    meter_pw_high_encrypt=""
    meter_pw_low_encrypt=""
    meter_color_body_man=""
    meter_color_body=""
    meter_confg_check="0"
    meter_config_res_list=[]
    meter_mc_model_list=[]
    energy_consumed=""
    energy_export=""
    meter_config_disconnect_control_mode=""
    disconnect_control_mode=""

    mc_test_ok=True

    meter_pw_high_encrypt=""

    meter_config_param_filename=""
    meter_config_filename=""

    delta_pc_minus_device=0
    otnoshenie_str=""
    meter_grade=""
    duration_test=0

    rc_serial_number=""     #серийный номер пульта управления ПУ
    rc_tech_number=""       #технический номер пульта управления ПУ
    rc_soft=""

    gsm_serial_number = ""  #номер модуля связи
    gsm_SIM_number=""       #номер SIM карты
    gsm_soft=None             #версия ПО модуля связи, прочитанная через ПУ
    modem_status=""         #статус модема по умолчанию: "0"-не будет устанавливаться, "1"-рабочий, "2"-тестовый
    
    gsm_product_type=None
    gsm_docked_tn_sutp=None
    gsm_docked_sn_sutp=None
    gsm_docked_model_sutp=None
    gsm_docked_soft_sutp=None
    
    rep_err_list=[]
    clipboard_err_list=[]
    rep_remark_list=[]
    test_start_time=""

    test_result="" 
    test_result_buf="" 
    
    default_filename_full=""
    
    res=readGonfigValue(file_name_in="opto_run.json",
        var_name_list=[],default_value_dict=default_value_dict)
    if res[0]!="1":
        a_err_txt="Ошибка при чтении данных из ф.opto_run.json."
        printWARNING(a_err_txt)
        keystrokeEnter()
        return ["0", a_err_txt]
    
    default_value_dict=res[2]
    readDefaultValue(default_value_dict)

    
    res=innerGetVarDefAll()
    if res[0]=="0":
        return ["0", res[1]]
    
    
    _, _, default_dirname=getUserFilePath(file_name="otk_report",
        only_dir="1",workmode=workmode)
    if default_dirname == "":
        a_err_txt="Ошибка при формировании пути к локальной папке с отчетами."
        printWARNING(a_err_txt)
        keystrokeEnter()
        return ["0", a_err_txt]
    
    
    
    _, _, work_dirname=getUserFilePath(file_name="work_dirname",
        only_dir="1",workmode=workmode)
    if work_dirname == "":
        a_err_txt="Ошибка при формировании пути к общей папке с отчетами."
        printWARNING(a_err_txt)
        keystrokeEnter()
        return ["0", a_err_txt]
    
    _, _, dirname_sos=getUserFilePath(file_name="sharedFolder",
        only_dir="1",workmode=workmode)
    if dirname_sos=="":
        a_err_txt="Ошибка при формировании пути к резервной папке с отчетами."
        printWARNING(a_err_txt)
        keystrokeEnter()
        return ["0", a_err_txt]
    
    if len(rep_err_list)>0:
        test_result="\n".join(rep_err_list)
        test_result_buf="\n".join(clipboard_err_list)

    
    new_filename=filename_rep

    filename_report=new_filename+"_отчет.txt"

    default_filename_full=os.path.join(default_dirname, filename_report)

    if read_mode_opto=="отчет о конфигурации":
        a_path_old=default_filename_full

        filename_report=new_filename+"_конф.txt"

        default_filename_full=os.path.join(default_dirname, filename_report)

        try:
            os.rename(a_path_old, default_filename_full)
            
        except Exception:
            txt1 = "Ошибка при изменении имени файла " \
                f"{a_path_old}."
            printFAIL(txt1)
            return ["0", "Ошибка при изменении имени файла"] 


    filename_ext=new_filename+"_внеш.осмотр.txt"
    default_filename_ext =os.path.join(default_dirname, filename_ext)

    connection_initialized=False

    
    res=checkVarProgrAvailable("sutp_to_save", sutp_to_save, "1")
    if res[0]!="1":
        sys.exit()


    if data_exchange_sutp!="0":
        res=getInfoAboutDevice(meter_tech_number, workmode, employee_id, 
                               employee_pw_encrypt, "1")
        if res[0]=="0":
            print(f"\n{bcolors.WARNING}Не удалось получить информацию о "
                  f"модели и версии ПО ПУ из СУТП.{bcolors.ENDC}")
        elif res[0] in ["1", "2"]:
            meter_model_sutp=res[12]
            meter_soft_sutp=str(res[17])

    
    txt=f"{bcolors.OKGREEN}3.2. Проверка конфигурации ПУ:{bcolors.ENDC}"
    file_path=""

    if  meter_config_check[0]!="0" and data_exchange_sutp=="1":
        res=getMeterConfigFilePath(meter_tech_number, "1", workmode, 
            "0", None)
        if res[0]!="0":
            file_path=res[2]
            meter_config_param_filename=res[3]
            meter_config_filename=res[4]

    if meter_config_check[0]=="0":
        txt=txt+f"{bcolors.WARNING} не проводилась.{bcolors.ENDC}"

    else:
        a_txt1=f"{bcolors.OKGREEN}Имя файла конфигурации ПУ, указанного в заказе: " \
            f"{meter_config_param_filename}.{bcolors.ENDC}"
        
        a_txt2=f"{bcolors.OKGREEN}Имя файла, использованного при конфигурировании ПУ: " \
            f"{meter_config_filename}.{bcolors.ENDC}"

        if data_exchange_sutp!="1":
            txt=txt+f"\n     {bcolors.WARNING}Обмен данными с СУТП отключен.{bcolors.ENDC}"

        if meter_config_param_filename=="":
            a_txt1=f"{bcolors.WARNING}Имя файла конфигурации ПУ, указанного в заказе: " \
                f"не удалось получить из СУТП.{bcolors.ENDC}"
        
        if meter_config_filename=="":
            a_txt2=f"{bcolors.WARNING}Имя файла, использованного при конфигурировании ПУ: " \
                f"не удалось получить из СУТП.{bcolors.ENDC}"
        
        txt=txt+f"\n     {a_txt1}\n     {a_txt2}"
        if meter_config_param_filename!="" and meter_config_filename!="" and \
            meter_config_param_filename!=meter_config_filename:
            txt=txt+f"\n     {bcolors.WARNING}Имя файла конфигурации ПУ, указанного в заказе " \
                    f"({meter_config_param_filename}) отличается от имени файла, " \
                    f"использованного при конфигурировании ПУ ({meter_config_filename})." \
                    f"{bcolors.ENDC}"
            
    if meter_config_check[0] in ["2", "3"]:
        mass_prod_vers=""
        res=readGonfigValue("mass_config.json", [], {}, workmode, "1")
        if res[0]=="1":
            mass_prod_vers=res[2]["mass_prod_vers"]

        a_txt=f"\n     {bcolors.OKGREEN}Проверка конфигурации ПУ проводилась с помощью программы " \
                f"'MassProdAutoConfig.exe' версией {mass_prod_vers}. Тест пройден.{bcolors.ENDC}"
        
        if meter_config_check[0]=="2":
            a_txt=f"\n     {bcolors.OKGREEN}Проверка конфигурации ПУ проводилась с помощью программы " \
                f"'MassProdAutoConfig.exe' в ручном режиме. Тест пройден.{bcolors.ENDC}"
            
        if len(meter_config_res_list)>0:
            a_err_list=meter_config_res_list.copy()
            for i in range(0, len(a_err_list)):
                a_err_list[i] = f"     {bcolors.FAIL}- " \
                    f"{a_err_list[i]}{bcolors.ENDC}"
            a_err_txt="\n".join(a_err_list)
            a_txt=f"\n     {bcolors.OKGREEN}Проверка конфигурации ПУ проводилась с помощью программы " \
                f"'MassProdAutoConfig.exe' версией {mass_prod_vers}.{bcolors.ENDC}" \
                f"\n     {bcolors.FAIL}Выявлены следующие замечания:{bcolors.ENDC}" \
                f"\n{a_err_txt}"
        txt=txt+a_txt

        if meter_config_check[0]=="3":
            res=readGonfigValue("mass_log_line_multi.json", [], {}, 
                workmode, "1")
            if res[0]=="1":
                a_mask_file_name=res[2]["mask_file_name"]
                a_txt=f"\n     {bcolors.OKGREEN}Имя mask-файла: {a_mask_file_name}.{bcolors.ENDC}"
                txt=txt+a_txt

    elif meter_config_check[0]=="1":
        txt1_2 = f"Проверка прервана, т.к. проверка через otk не реализована."
        testBreak(txt, txt1_2, default_filename_full, employees_name)
        return ["0", txt1_2]
    
        
    err_msg="Сохранение в отчет результата проверки конфигурационных " \
        "параметров ПУ."
    fileWriter(default_filename_full,"a", "", txt+"\n", err_msg,
        "on", "", "on", "on")
    
    if read_mode_opto=="полная проверка":
        print(f"Запрашиваю у ПУ данные...")
        res=innerToInitConnectOpto(com_opto)
        if res=="0" or res=="9":
            a_dic={"0": "Ошибка при запросе у ПУ данных",
                "9": "Проверка ПУ прервана пользователем."}
            return [res, a_dic[res]]

        
        mc_on_board="0"
        mc_checkVersPO_set=""
        res=toGetProductInfo2(meter_serial_number, "Product1", workmode,)
        if res[0]=="0":
            return ["0", "Ошибка при получении информации из ф.ProductNumber.xlsx."]
        
        mc_num_set=res[19]
        rc_num_set=res[20]
        mc_checkVersPO_set=res[22]
        meter_presence_relay=res[13]
        
        a_mc=res[26]
        if a_mc!=None and a_mc!="":
            meter_mc_model_list=a_mc.split(",")
                
        mc_on_board_set=res[27]



        if meter_presence_relay=="да" and meter_config_param_filename!="" \
            and file_path!="":
            a_filter_dic={"Parameter": ["Режим работы реле"]}
            res=getDataFromXlsx(file_path, "Config", a_filter_dic, "", 1)
            if res[0] in ["0", "2"]:
                a_txt=f"При получении данных из файла с " \
                    f"конфигурацией ПУ возникла ошибка:\n{res[1]}"
                printWARNING(a_txt)
                keystrokeEnter()
                return ["0", a_txt]
            
            meter_config_disconnect_control_mode=str(res[2][0].get("Value", None))
            
            while True:
                disconnect_control = GXDLMSDisconnectControl("0.0.96.3.10.255")
                try:
                    disconnect_control_mode = str(reader.read(disconnect_control, 4))
                    break
                except Exception as e:
                    oo = communicationTimoutError(
                        "Определяли состояние реле: ", e.args[0])
                    if oo == "0" or oo == "-1":
                        a_dic={"0": ["Ошибка связи с ПУ при чтении состояния реле.", "0"],
                            "-1": ["Проверка прервана пользователем.", "9"]}
                        ret_txt=a_dic[oo][0]
                        ret_id=a_dic[oo][1]
                        return [ret_id, ret_txt]
                    
            txt=f"     {bcolors.OKGREEN}В конфигурационном файле указан режим работы реле: " \
                f"{meter_config_disconnect_control_mode}.{bcolors.ENDC}\n" \
                f"     {bcolors.OKGREEN}Фактический режим работы реле: " \
                f"{disconnect_control_mode}.{bcolors.ENDC}"

            if disconnect_control_mode!=meter_config_disconnect_control_mode:
                a_err_txt=f"Фактический режим работы реле {disconnect_control_mode} " \
                    f"отличается от значения, указанного в конфигурационном файле " \
                    f"{meter_config_disconnect_control_mode}."
                rep_err_list.append(a_err_txt)
                clipboard_err_list.append(a_err_txt)
                txt=txt+f"\n     {bcolors.FAIL}{a_err_txt}{bcolors.ENDC}"


            err_msg="Сохранение в отчет результата проверки режима работы реле ПУ."
            fileWriter(default_filename_full,"a", "", txt+"\n", err_msg,
                "on", "", "on", "on")
            
            if disconnect_control_mode!=meter_config_disconnect_control_mode:
                print()
                a_txt=f"{bcolors.FAIL}Режим работы реле отличается от необходимого " \
                    f"режима.{bcolors.ENDC}"
                res=innerSelectActions(a_txt, rep_err_list, "1", "конфигурация ПУ")
                if res[0]=="9":
                    txt1_2 = f"Проверка прервана пользователем."
                    testBreak("", txt1_2, default_filename_full, employees_name)
                    return ["9", txt1_2]
                elif res[0]=="2":
                    return ["2", "ПУ отправлен в ремонт."]

        
        res=toReadStaticDataOpto()
        if res==0:
            return ["0", "Обрыв связи с ПУ при чтении статических данных."]
        elif res==9:
            txt1_2 = f"Проверка прервана пользователем при считывании статических данных из ПУ."
            testBreak("",txt1_2,default_filename_full, employees_name)
            return ["9", txt1_2]

        toCloseConnectOpto()
        connection_initialized = False

            
        res=checkVersMeter(meter_soft, actual_device_version_list,
                        print_msg="1")
        if res[0]=="2":
            txt1=f"Версия ПО прибора учета: {meter_soft}. Необходимо обновление ПО."
            rep_err_list.append(txt1)
            clipboard_err_list.append(txt1)
            header=f"{bcolors.WARNING}{txt1}{bcolors.ENDC}"
            res=innerSelectActions(header, rep_err_list, "1")
            if res[0] in ["2", "9"]:
                a_dic={"2": ["ПУ отправлен в ремонт.", "1"],
                    "9": ["Проверка прервана пользователем.", "9"]}
                ret_txt=a_dic[res[0]][0]
                ret_id=a_dic[res[0]][1]
                return [ret_id, ret_txt]


        res=innerFirstCheckMC()
        if res[0] in ["0", "2", "9"]:
            a_dic={"0": ["Ошибка при сверке параметров МС", "0"],
                "2": ["ПУ отправлен в ремонт.", "1"],
                "9": ["Проверка прервана пользователем.", "9"]}
            ret_txt=a_dic[res[0]][0]
            ret_id=a_dic[res[0]][1]
            return [ret_id, ret_txt]
                    

        res = innerToInitConnectOpto(com_opto)
        if res=="1":
            connection_initialized = True

        else:
            return ["9", "Проверка прервана пользователем."]
        

        while test_go:
            print()
            txt1=""
            
            txt="3.3. Дата выпуска ПУ (калибровки): "
            if meter_date_of_manufacture=="":
                meter_date_of_manufacture="не указана"
                txt=txt+meter_date_of_manufacture
                print(f"{bcolors.FAIL}{txt}{bcolors.ENDC}")
            else:
                txt=txt+ meter_date_of_manufacture
                print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")
            
            err_msg="Сохранение в отчет информации о дате выпуска ПУ."
            fileWriter(default_filename_full,"a", "", txt+"\n", err_msg,
                "on", "", "on")
            

            txt1_1=""
            a_color=bcolors.OKGREEN
            if meter_type!=meter_type_ep:
                a_color=bcolors.FAIL
                txt1_1=f"\n     Тип ПУ, указанный в электронном паспорте ПУ ({meter_type_ep}) " \
                    "отличается от типа, который определен по коду изделия в серийном номере ПУ " \
                    f"({meter_type}.)"
                rep_err_list.append(txt1_1)
                clipboard_err_list.append(txt1_1)
            txt1_rep=f"4.1. Тип прибора учета:\n" \
                f"     - определенный по серийному номеру: {meter_type}\n" \
                f"     - по данным электронного паспорта: {meter_type_ep}" \
                f"{txt1_1}"
            txt1=f"{bcolors.OKGREEN}4.1. Тип прибора учета:{bcolors.ENDC}\n" \
                f"{bcolors.OKGREEN}     - определенный по серийному номеру: {meter_type}{bcolors.ENDC}\n" \
                f"{bcolors.OKGREEN}     - по данным электронного паспорта: {bcolors.ENDC}{a_color}" \
                f"{meter_type_ep}{bcolors.ENDC}" \
                f"{txt1_1}"
            print(f"{txt1}")

            err_msg="Сохранение в отчет результата проверки типа ПУ"
            fileWriter(default_filename_full,"a", "", txt1_rep+"\n", err_msg,
                "on", "", "on")
            
            sheet_name = "Product1"
            res = toGetProductInfo2(meter_serial_number, sheet_name)
            if res[0] == "0":
                txt1_2 = f"Проверка прервана, т.к. не удалось определить модель ПУ по его номеру."
                testBreak(txt,txt1_2,default_filename_full, employees_name)
                return ["9", txt1_2]
            meter_product_type = res[2]

            
            a_color_sutp=bcolors.OKGREEN
            a_color_capt=bcolors.OKGREEN
            a_model_caption=" "+meter_product_type+" (значение по умолчанию)"
            a_diff_txt=""
            if meter_type_def == "спрашивать каждый раз" or meter_product_type != meter_type_def:
                oo = questionSpecifiedKey(bcolors.OKBLUE, "     На крышке ПУ указана модель '" 
                    +meter_product_type +"'. Верно? 0- нет, 1- да",["0","1"])
                print()
                if oo=="0":
                    a_model_caption="отличается от модели, определенной по серийному номеру"
                    a_color_capt=bcolors.FAIL
                    a_txt="Модель ПУ, указанная на крышке ПУ отличается от модели, " \
                        "которая определена по коду изделия в серийном номере ПУ " \
                        f"'{meter_product_type}'."
                    rep_err_list.append(a_txt)
                    clipboard_err_list.append(a_txt)
                    a_diff_txt=a_diff_txt+f"\n       {a_color_capt}{a_txt}{bcolors.ENDC}"
                
                else:
                    a_model_caption=" "+meter_product_type
                    if meter_type_def!="спрашивать каждый раз":
                        a_txt=f"Заменить модель ПУ по умолчанию '{meter_type_def}' \n" \
                            f"на новое значение '{meter_product_type}'? 0-нет, 1-да"
                        oo=questionSpecifiedKey(bcolors.OKBLUE, a_txt, ["0","1"], "", 1)
                        print()
                        if oo=="1":
                            meter_type_def=meter_product_type
                            a_save_ok=True
                            default_value_dict = writeDefaultValue(default_value_dict)
                            res=saveConfigValue('opto_run.json', default_value_dict)
                            if res[0]=="0":
                                printWARNING(res[1])
                                a_save_ok=False

                            a_dic={"meter_type_def": meter_type_def}
                            res=saveConfigValue('opto_run.json_last', a_dic)
                            if res[0]=="1" and a_save_ok:
                                print(f"{bcolors.OKGREEN}Установлено новое значение по " 
                                f"умолчанию для модели ПУ: {meter_type_def}.")
                            else:
                                printWARNING(res[1])



            if meter_model_sutp!=None and meter_model_sutp!=meter_product_type:
                a_color_sutp=bcolors.WARNING
                a_txt=f"Модель ПУ, указанная в СУТП '{meter_model_sutp}' " \
                    f"отличается от модели, которая определена по коду " \
                    f"изделия в серийном номере ПУ '{meter_product_type}'."
                rep_remark_list.append(a_txt)
                a_diff_txt=f"\n       {a_color_sutp}{a_txt}{bcolors.ENDC}"

            a_color_model_ep=bcolors.OKGREEN
            if meter_model_ep==None or meter_model_ep!=meter_product_type:
                a_txt=f"В электронном паспорте ПУ нет информации о его модели."
                if meter_model_ep!=meter_product_type:
                    a_txt=f"Модель ПУ, указанная в электронном паспорте '{meter_model_ep}' " \
                        f"отличается от модели, которая определена по коду " \
                        f"изделия в серийном номере ПУ '{meter_product_type}'."
                a_color_model_ep=bcolors.FAIL
                rep_err_list.append(a_txt)
                clipboard_err_list.append(a_txt)
                a_diff_txt=a_diff_txt+f"\n       {a_color_model_ep}{a_txt}{bcolors.ENDC}"


            txt1=f"{bcolors.OKGREEN}4.2. Модель ПУ:{bcolors.ENDC}\n" \
                f"{bcolors.OKGREEN}     - определенная по серийному номеру: {meter_product_type}{bcolors.ENDC}\n" \
                f"{bcolors.OKGREEN}     - указанная на крышке ПУ: {bcolors.ENDC}" \
                f"{a_color_capt}{a_model_caption}{bcolors.ENDC}\n" \
                f"{bcolors.OKGREEN}     - указанная в электронном паспорте ПУ: {bcolors.ENDC}" \
                f"{a_color_model_ep}{meter_model_ep}{bcolors.ENDC}\n" \
                f"{bcolors.OKGREEN}     - по данным из СУТП: {bcolors.ENDC}" \
                f"{a_color_sutp}{meter_model_sutp}{bcolors.ENDC}" \
                f"{a_diff_txt}"
            fileWriter(default_filename_full,"a", "", txt1+"\n", err_msg,
                "on", "", "on", "on")
            
            
            txt="4.3. Цвет корпуса прибора учета: "
            if meter_color_body_man!="":
                txt = f"{txt}{meter_color_body_man}"
                print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")
            elif meter_color_body=="спрашивать каждый раз":
                print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")
                txt1="     Укажите цвет корпуса ПУ:\n"
                meter_color_dict = ["белый", "черный", "серый"]
                l1 = len(meter_color_dict)
                m1=[]
                for i in range(l1):
                    m1.append(str(i+1))
                    if i < (l1-1):
                        txt1=txt1+"     "+ str(i+1)+" - "+ meter_color_dict[i]+"\n"
                    else:
                        txt1=txt1+"     "+ str(i+1)+" - "+ meter_color_dict[i]
                oo = questionSpecifiedKey(bcolors.OKBLUE,txt1,m1)
                print()
                meter_color_body=meter_color_dict[int(oo)-1]
                txt=txt+meter_color_body
            else:
                txt=txt+"указан по умолчанию - "+meter_color_body 
                print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")      
            with open(default_filename_full, "a", errors="ignore") as file:
                file.write(f"{txt}" + "\n")
                

            res=innerReadMeterTime()
            if res[0] in ["0", "9"]:
                return [res[0], res[1]]

            device_date_time1=res[3]
            pc_time=res[4]
            pc_time1=res[5]
            delta_pc_minus_device=res[6]
 
            txt="5. Проверка текущей даты и времени:\n" \
                f"   Дата и время в ПЭВМ: {pc_time1}\n" \
                f"   Дата и время в ПУ:   {device_date_time1}"
            
            with open(default_filename_full, "a", errors="ignore") as file:
                file.write(f"{txt}" + "\n")
            
            printGREEN(txt)

            try:
                datetime_normal = f'{int.from_bytes(last_clock_sync[3:4]):02}.{int.from_bytes(last_clock_sync[2:3]):02}.{int.from_bytes(last_clock_sync[:2])}' + " " + str(
                    int.from_bytes(last_clock_sync[5:6])) + ":" + str(
                    int.from_bytes(last_clock_sync[6:7])) + ":" + str(
                    int.from_bytes(last_clock_sync[7:8]))
                if datetime_normal == "00.00.0 0:0:0":
                    datetime_normal = "01.01.2000 0:00:00"
                txt="   Дата и время последней корректировки времени: "+datetime_normal
                with open(default_filename_full, "a", errors="ignore") as file:
                    file.write(f"{txt}" + "\n")
                print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")
            except Exception:
                with open(default_filename_full, "a", errors="ignore") as file:
                    txt="   Дата и время последней корректировки времени: НЕ ПРОВОДИЛАСЬ"
                    file.write(f"{txt}" + "\n")
                print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")
                datetime_normal = "01.01.2000 0:00:00"

            delta_MSK_minus_timezone = (-180 - timezone) / 60
            timezone_txt="МСК"
            if delta_MSK_minus_timezone>0:
                timezone_txt=f"МСК+{delta_MSK_minus_timezone}"
            
            elif delta_MSK_minus_timezone<0:
                timezone_txt=f"МСК{delta_MSK_minus_timezone}"

            txt=f"   Часовой пояс: {timezone_txt}"
            
            with open(default_filename_full, "a", errors="ignore") as file:
                file.write(f"{txt}" + "\n")
            print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")

            txt="   Расхождение по времени, сек: "+ str(delta_pc_minus_device)
            with open(default_filename_full, "a", errors="ignore") as file:
                file.write(f"{txt}" + "\n")
            
            if delta_pc_minus_device<60:
                printGREEN(txt)
            else:
                printWARNING(txt)

        
            try:
                datetime_normal1 = datetime.strptime(datetime_normal, "%d.%m.%Y %H:%M:%S")
            except Exception:
                datetime_normal1 = datetime.strptime("01.01.2000 0:00:00", "%d.%m.%Y %H:%M:%S")
            
            delta_pc_minus_corrected = int(abs(pc_time - datetime_normal1).days + 1)

            otnoshenie = delta_pc_minus_device / delta_pc_minus_corrected
            otnoshenie_str = "{:.2f}".format(otnoshenie)

            limit_otnoshenie=1
            txt="   Расчетная погрешность со времени последней установки времени, сек/сутки: " \
                f"{otnoshenie_str}\n   Допустимое значение погрешности при откл. напряжении, сек/сут: {limit_otnoshenie}"
            with open(default_filename_full, "a", errors="ignore") as file:
                file.write(f"{txt}" + "\n")

            if otnoshenie>limit_otnoshenie:
                printWARNING(txt)
            else:
                printGREEN(txt)

            if otnoshenie>limit_otnoshenie or delta_pc_minus_device>60: #delta_pc_minus_device!=0:
                if meter_adjusting_clock=="1" or (meter_adjusting_clock=="2" and 
                    str(timezone) == "-180"):
                    oo="1"
                
                else:
                    a_txt="   Скорректировать время? 0-нет, 1-да"
                    oo = questionSpecifiedKey(bcolors.OKBLUE,a_txt, ["0","1"],
                        "ToCorrectClock", 1)
                    print()

                if oo=="1":
                    tz_local=toformatNow()[8]*(-1)/60

                    while True:
                        timenow = time.time()
                        tz_delta=(tz_local-timezone)*60
                        time_new=timenow+tz_delta
                        while True:
                            try:
                                data = GXDLMSClock("0.0.1.0.0.255")
                                a_t=datetime.fromtimestamp(time_new)
                                data.time = GXDateTime(a_t)
                                reader.write(data, 2)
                                break
                            except Exception as e:
                                oo = communicationTimoutError("Корректировка даты и времени в ПУ: ", e.args[0])
                                if oo=="0" or oo == "-1":
                                    a_dic={"0": ["Ошибка связи с ПУ при корректировке даты и времени.", "0"],
                                        "-1": ["Проверка прервана пользователем.", "9"]}
                                    ret_txt=a_dic[oo][0]
                                    ret_id=a_dic[oo][1]
                                    return [ret_id, ret_txt]

                        timezone_hour=timezone*(-1)/60
                        txt=f"   Дата и время скорректированы по часовому поясу {timezone_txt}.\n"
                        
                        
                        res=innerReadMeterTime()
                        if res[0] in ["0", "9"]:
                            return [res[0], res[1]]

                        device_date_time1=res[3]
                        pc_time1=res[5]
                        delta_pc_minus_device=res[6]

                        txt=f"{txt}   Дата и время после корректировки в ПУ:\n" \
                            f"   - в ПЭВМ: {pc_time1}\n" \
                            f"   - в ПУ:   {device_date_time1}"

                        if delta_pc_minus_device>60:
                            printWARNING(txt)
                            a_txt=f"Расхождение по времени составляет {delta_pc_minus_device} сек.\n" \
                                "Повторить корректировку времени? 0-нет, 1-да"
                            oo = questionSpecifiedKey(bcolors.OKBLUE, a_txt, ["0","1"], "", 1)
                            print()
                            if oo=="1":
                                continue

                            else:
                                a_txt="   Дата и время в ПУ остались прежними."
                                printWARNING(a_txt)
                            
                        else:
                            printGREEN(txt)

                        with open(default_filename_full, "a", errors="ignore") as file:
                            file.write(f"{txt}\n")
                        
                        break
                        
                elif oo=="0":
                    txt="   Дата и время в ПУ остались прежними."
                    printWARNING(txt)
                    with open(default_filename_full, "a", errors="ignore") as file:
                        file.write(f"{txt}" + "\n")


            default_day_active = [[1, ['00:00:00', '0.0.10.0.100.255', 2, '07:00:00', '0.0.10.0.100.255', 1, '23:00:00',
                                    '0.0.10.0.100.255', 2]]]  # НЕ ТРОГАТЬ, ДЕФОЛТНОЕ ТР ДЛЯ ДНЯ
            cicl1=True
            while cicl1:
                try:
                    device_day_active = reader.read_day_profile_active()
                    break
                except Exception as e:
                    oo = communicationTimoutError("Считывание тарифного расписания из ПУ: ", e.args[0])
                    if oo=="0" or oo == "-1":
                        a_dic={"0": ["Ошибка связи с ПУ при чтении состояния реле.", "0"],
                            "-1": ["Проверка прервана пользователем.", "9"]}
                        ret_txt=a_dic[oo][0]
                        ret_id=a_dic[oo][1]
                        return [ret_id, ret_txt]

            txt="6. Тарифное расписание: по умолчанию"
            if default_day_active == device_day_active:
                with open(default_filename_full, "a", errors="ignore") as file:
                    file.write(f"{txt}" + "\n")
                    print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")
            else:
                txt="6. Тарифное расписание: НЕСТАНДАРТНОЕ"
                with open(default_filename_full, "a", errors="ignore") as file:
                    file.write(f"{txt}" + "\n")
                    print(f"{bcolors.WARNING}{txt}{bcolors.ENDC}")

            txt="7. Проверка версии ПО прибора учета: "
            res=checkVersMeter(meter_soft=meter_soft,device_ver_list=actual_device_version_list,
                            print_msg="0")
            txt1=f"   Список актуальных версий ПО: {res[2]}"
            print(f"{bcolors.OKGREEN}{txt}\n{txt1}{bcolors.ENDC}")
            a_meter_soft_sutp=meter_soft_sutp
            if meter_soft_sutp==None or meter_soft_sutp=="":
                a_meter_soft_sutp="None"
            txt1_rep=txt+"\n"+txt1
            txt1="\n   Версия ПО прибора учета:\n" \
                f"   - полученная из ПУ: {meter_soft}\n" \
                f"   - полученная из СУТП: {a_meter_soft_sutp}"
            a_color=bcolors.OKGREEN
            txt1_1=""

            if meter_soft_sutp!=meter_soft:
                a_color=bcolors.WARNING
                txt1_2=f"Версия ПО ПУ в СУТП ({meter_soft_sutp}) отличается " \
                    f"от фактической версии ПО ПУ ({meter_soft})."
                txt1_1="\n    "+txt1_2
                if meter_soft_sutp==None or meter_soft_sutp=="":
                    txt1_1=""
                    txt1_2=f"В СУТП отсутствует информация о версии ПО ПУ ({meter_soft})."
                    a_color=bcolors.WARNING
                rep_remark_list.append(txt1_2)
            txt1_rep=txt1_rep+txt1+txt1_1
            txt1_2=f"{bcolors.OKGREEN}   Версия ПО прибора учета:\n" \
                f"   - полученная из ПУ: {meter_soft}\n" \
                f"   - полученная из СУТП: {bcolors.ENDC}" \
                f"{a_color}{a_meter_soft_sutp}{bcolors.ENDC}" \
                f"{bcolors.WARNING}{txt1_1}{bcolors.ENDC}"
            print (txt1_2)

            if res[0]=="2":
                txt1="Необходимо обновление ПО."
                print(f"{bcolors.FAIL}{txt1}{bcolors.ENDC}")
                txt1_1="Необходимо обновление ПО ПУ (записана версия "+meter_soft+")"
                rep_err_list.append(txt1_1)
                clipboard_err_list.append(txt1_1)
                txt1_rep=txt1_rep+"\n"+txt1
            
            fileWriter(default_filename_full, "a", "", txt1_rep+ "\n", \
                "Сохранение в отчет результата проверки версии ПО ПУ",join="on")
            

            res=TestBattery(test_num="8")
            if res[0]=="0":
                return ["0", "Ошибка связи с ПУ при проведении теста заряда батареи."]

            
            if meter_phase =='3':
                txt=f"9.1. Мгновенные значения напряжения, В: {meter_voltage_str}\n" \
                    f"9.2. Мгновенные значения тока, мА: {meter_amperage_str}\n"
            elif meter_phase =="1":
                txt=f"9.1. Мгновенное значение напряжения, В: {meter_voltage_str}\n" \
                    f"9.2. Мгновенное значение тока, мА: {meter_amperage_str}\n"
                
            res=readGonfigValue("var_all_value.json", [], {},
                workmode, "1")
            if res[0]=="1":
                a_dic=res[2].get("electrical_test_circuit", {})
                if len(a_dic)>0:
                    a_dic=a_dic.get("all_value",{})
                    a_descript=list(a_dic.keys())
                    a_val=list(a_dic.values())
                    a_dic=dict(zip(a_val, a_descript))
            txt=txt+f"     Схема подключения ПУ на испытательном стенде: {a_dic[electrical_test_circuit]}"
            if electrical_test_circuit in ["1-1", "2-1", "2-2", "3-1",
                "3-2", "3-2", "3-3"]:
                txt=txt+f"\n     Контрольное значение тока {ctrl_current_electr_test*1000} мА."
            printGREEN(txt)
            with open(default_filename_full, "a", errors="ignore") as file:
                file.write(f"{txt}" + "\n")


            a_color_consumed=bcolors.OKGREEN
            a_color_export=bcolors.OKGREEN
            a_err_consumed=""
            a_err_export=""
            a_limit_dic={"1": 6, "3": 20}
            energy_consumed=meter_energy_dic['energy_consumed']
            energy_export=meter_energy_dic['energy_export']
            if int(energy_consumed)>a_limit_dic[meter_phase]:
                a_color_consumed=bcolors.FAIL
                a_err_consumed=f"\n     {a_color_consumed}Превышает допустимое значение " \
                    f"{a_limit_dic[meter_phase]} кВтч.{bcolors.ENDC}"
                a_txt_consumed=f"Значение общей потребленной энергии {energy_consumed} " \
                    f"кВтч превышает допустимое значение {a_limit_dic[meter_phase]} кВтч."

            if int(energy_export)>a_limit_dic[meter_phase]:
                a_color_export=bcolors.FAIL
                a_err_export=f"\n     {a_color_export}Превышает допустимое значение " \
                    f"{a_limit_dic[meter_phase]} кВтч.{bcolors.ENDC}"
                a_txt_export=f"Значение общей обратной энергии {energy_export} " \
                    f"кВтч превышает допустимое значение {a_limit_dic[meter_phase]} кВтч."

            txt=f"{a_color_consumed}9.3. Общая потребленная активная энергия, кВтч: " \
                    f"{energy_consumed}{bcolors.ENDC}" \
                    f"{a_err_consumed}" \
                f"\n{a_color_export}9.4. Общая обратная активная энергия, кВтч: " \
                    f"{energy_export}{bcolors.ENDC}" \
                    f"{a_err_export}"

            err_msg="Сохранение в отчет информации о накопленной энергии."
            fileWriter(default_filename_full,"a", "", txt+"\n", err_msg,
                "on", "", "on", "on")


            if a_err_consumed!="" or a_err_export!="":
                a_err_list=rep_err_list.copy()

                if a_err_consumed!="":
                    a_err_list.append(a_txt_consumed)

                if a_err_export!="":
                    a_err_list.append(a_txt_export)

                print()
                a_err_msg=f"{bcolors.FAIL}Значение накопленной энергии " \
                    f"превышает допустимое значение.{bcolors.ENDC}"
                res=innerSelectActions(a_err_msg, a_err_list, "5", "энергия")
                if res[0]=="9":
                    txt1_2 = f"Проверка прервана пользователем."
                    testBreak("", txt1_2, default_filename_full, employees_name)
                    return ["9", txt1_2]
                
                elif res[0]=="2":
                    return ["2", "ПУ отправлен в ремонт."]
                
                elif res[0]=="1":
                    if a_err_consumed!="":
                        rep_err_list.append(a_txt_consumed)
                        clipboard_err_list.append(a_txt_consumed)
                    
                    if a_err_export!="":
                        rep_err_list.append(a_txt_export)
                        clipboard_err_list.append(a_txt_export)
    
                elif res[0]=="5":
                    if a_err_consumed!="":
                        rep_remark_list.append(a_txt_consumed)

                    if a_err_export!="":
                        rep_remark_list.append(a_txt_consumed)

                    

            magnetic_status="указан неизвестный код"
            magnetic_field_bin= format(magnetic_field, "08b")
            magnetic_field_0=magnetic_field_bin[-1]
            magnetic_field_2=magnetic_field_bin[-3]
            if magnetic_field_0=='0':
                magnetic_status="не было зафиксировано"
            elif magnetic_field_2=='0':
                magnetic_status="было зафиксировано, сейчас отсутствует"
            elif magnetic_field_2=='1':
                magnetic_status="было зафиксировано, сейчас присутствует"

            txt="10. Магнитное поле: "+ magnetic_status
            if magnetic_field_0 == '0':
                print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")
            else:
                cicl=True
                while cicl:
                    cicl=False
                    res=clearSignMagneticField()
                    if res[0]=="0":
                        return ["0", "Ошибка связи с ПУ при проверке фиксации магнитного поля"]
                    elif res[0]=="2":
                        print(f"{bcolors.OKBLUE}   Не удалось сбросить признак фиксации магнитного поля. " \
                            f"Повторить? (1-да и нажмите Enter)\n   {txt_break_test}:{bcolors.ENDC}", end="")
                        oo=input()
                        if oo=="/":
                            testBreak(txt,"Проверка прервана пользователем",default_filename_full, employees_name)
                            return ["9", "Проверка прервана пользователем."]
                        if oo=="1":
                            cicl=True
                        else:
                            a_txt="НЕ УДАЛОСЬ сбросить признак фиксации магнитного поля."
                            print(f"{bcolors.FAIL}{a_txt}{bcolors.ENDC}")
                            txt=txt+a_txt
                            rep_err_list.append(a_txt)
                            clipboard_err_list.append(a_txt)
                    else:
                        txt=txt+"\n    Произведен сброс признака."
                        print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")
            with open(default_filename_full, "a", errors="ignore") as file:
                file.write(f"{txt}" + "\n")
                
            korpus_dict = {1: "разжат", 0: "зажат"}
            txt="11. Концевик корпуса: "+ korpus_dict[vskritie_korpusa]
            txt1_1="Концевик корпуса разжат."
            if vskritie_korpusa != 0:
                txt=txt+ ". ТРЕБУЕТСЯ РЕМОНТ"
                print(f"{bcolors.FAIL}{txt}{bcolors.ENDC}")
                rep_err_list.append(txt1_1)
                clipboard_err_list.append(txt1_1)

            else:
                print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")
            with open(default_filename_full, "a", errors="ignore") as file:
                file.write(f"{txt}" + "\n")

            
            res=toGetProductInfo2(meter_serial_number, "Product1", workmode)
            if res[0]=="0":
                return ["0", "Ошибка при получении данных из ф.ProductNumber.xlsx."]
            sil_kl=res[16]
            inf_kl=res[17]

            cover_count = 1
            if sil_kl=="да" and inf_kl=="да":
                cover_count=2


            to_compress_seal=False
            seal_soft="4.12.15"
            res = toGetProductInfo2(meter_serial_number, sheet_name)
            if res[0]=="1":
                if res[30]!=None and res[30]!="":
                    seal_soft=res[30]
            if cmpVers(meter_soft, seal_soft)=="<":
                to_compress_seal=True
            

            if to_compress_seal:
                res=testCompressOfSeal(test_num="12", cover_count=cover_count)
                txt=res[2]
                if res[0]=="0":
                    return ["0", "Ошибка связи с ПУ при обжатии электронной пломбы."]
                elif res[0]=="9":
                    testBreak(txt, "Проверка прервана пользователем", 
                                default_filename_full, employees_name)
                    return ["9", "Проверка прервана пользователем." ]
                
            else:
                txt = "12. Обжатие электронной пломбы: не требуется."
                print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")
            fileWriter(default_filename_full, "a", "", f"{txt}\n", \
                "Сохранение в отчет результата теста обжатия электронной пломбы",
                join="on")


            
            txt="13. Тест концевика крышки клеммников:"
            if cover_count==2:
                txt="13. Тест концевиков крышек клеммников:"
            print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}",end="")
            cicl2_step=cover_count

            test_cover_ok="1"

            while cicl2_step>0:
                if meter_type == 'i-prom.1':
                    cover_cur="крышки информационного клеммника"
                if cicl2_step==1:
                    cover_cur="крышки силового клеммника"
                print(f"{bcolors.OKGREEN}\n    Проверка концевика {cover_cur}{bcolors.ENDC}")
                res=questionCloseCover(cover_count)
                if res[0] == "0":
                    return ["0", "Ошибка связи."]
                elif res[0]=="9":
                    testBreak(txt,"Проверка прервана пользователем",default_filename_full, employees_name)
                    return ["9", "Проверка прервана пользователем."]
                elif res[0]=="8":
                    txt=txt+" пропущен."
                    test_cover_ok="2"
                    break

                elif res[0]=="2":
                    a_err_txt="Концевик крышки силового клеммника неисправен."
                    if meter_type == 'i-prom.1':
                        a_err_txt="Один или оба концевика крышек клеммников неисправны."
                    rep_err_list.append(a_err_txt)
                    clipboard_err_list.append(a_err_txt)
                    txt=txt+" не пройден. "+ a_err_txt
                    test_cover_ok="0"
                    break
                                
                obzim_result=5
                if to_compress_seal: 
                    res=checkSealCompress()
                    if res[0] == "0":
                        return ["0", "Ошибка связи."]
                    obzim_result=res[2]
                    if obzim_result!=5:
                        res=toCompressOfSeal(cover_count)
                        txt_err=f"Ошибка при обжатии электронной пломбы: {res[1]}."
                        if res[0]=="0":
                            return ["0", "Ошибка связи."]
                        elif res[0]=="2" or res[0]=="3" or res[0]=="4":
                            print(f"{bcolors.WARNING}{txt} {txt_err}{bcolors.ENDC}")
                            txt=txt+f" не пройден. {txt_err} "
                            test_cover_ok="0"
                            break
                        elif res[0]=="8":
                            txt = txt+" Тест пропущен" 
                            print(f"{bcolors.WARNING}\n   Тест пропущен{bcolors.ENDC}")
                            test_cover_ok="2"
                            break
                        elif res[0]=="9":
                            testBreak(txt, "Проверка прервана пользователем", 
                                    default_filename_full, employees_name)
                            return ["9", "Проверка прервана пользователем."]
                    print("Электронная пломба обжата.")
                    

                visual_control="1"
                if meter_type == 'i-prom.1' and cicl2_step==1:
                    visual_control="0"
                res=offAlarmSignDisplay(visual_control=visual_control)
                if res[0]=="0":
                    return ["0", "Ошибка связи с ПУ."]
                elif res[0]=="2":
                    txt = txt + \
                        f" не пройден. Не удалось погасить символ елочка на ЖКИ " \
                        f"при проверке концевика {cover_cur}."
                    print(f"{bcolors.WARNING}\n   Не удалось погасить символ елочка на ЖКИ.{bcolors.ENDC}")
                    txt1_2="При проверке концевика крышки клеммника не исчезает символ 'елочка'."
                    rep_err_list.append(txt1_2)
                    clipboard_err_list.append(txt1_2)
                    test_cover_ok="0"
                    break 
                if meter_type == 'i-prom.1' and cicl2_step==1:
                    print(f"\n    Символ елочка должен был исчезнуть...")

                print(f"{bcolors.OKBLUE}\n    Разожмите концевик {cover_cur}" \
                        f"{bcolors.ENDC}\n    Пауза до 3 сек.")
                for i in range(2):
                    time.sleep(1)
                    res=checkBtnTerminalCover()
                    if res[0]==0:
                        return ["0", "Ошибка связи с ПУ."]
                    if res[2]==1:
                        break
                txt1_1="    Посмотрите на ЖКИ: символ 'елочка' появился? 0-нет, 1-да"
                oo=questionSpecifiedKey(bcolors.OKBLUE,txt1_1,["0","1"],"LCDBellOn")
                res=checkBtnTerminalCover()
                if res[0]==0:
                    return ["0", "Ошибка связи с ПУ."]
                vskritie_klemm=res[2]
                if oo=="0":
                    if vskritie_klemm==1:
                        txt1_1="\n    Вы уверены, что символ 'елочка' на ЖКИ погашен? 0-нет, 1-да"
                        oo=questionSpecifiedKey(bcolors.WARNING,txt1_1,["0","1"],"LCDBellOff")
                        if oo=="1":
                            txt = txt+f" не пройден. Концевик {cover_cur} разжат, а символ елочка погашен." 
                            print(f"{bcolors.FAIL}\n   Концевик {cover_cur} разжат, а символ елочка погашен.{bcolors.ENDC}")
                            txt1_2=f"При проверке концевика крышки клеммника {cover_cur} не зажигается символ 'елочка'."
                            rep_err_list.append(txt1_2)
                            clipboard_err_list.append(txt1_2)
                            test_cover_ok="0"
                            break
                    else:
                        txt1=f"Тест не пройден. Концевик {cover_cur} остался зажат."
                        print(f"{bcolors.FAIL}\n    {txt1}{bcolors.ENDC}")
                        oo=questionRetry()
                        if oo==False:
                            txt=txt+txt1
                            txt1_2=f"Концевик {cover_cur} остался зажат при снятой крышке."
                            rep_err_list.append(txt1_2)
                            clipboard_err_list.append(txt1_2)
                            break
                else:
                    if vskritie_klemm==1:
                        if meter_type == 'i-prom.1':
                            print(f"\n{bcolors.OKGREEN}Тест концевика {cover_cur} успешно пройден.{bcolors.ENDC}")
                        cicl2_step-=1
                    else:
                        txt1_1="\n    Вы уверены, что символ 'елочка' на ЖКИ светится? 0-нет, 1-да"
                        oo=questionSpecifiedKey(bcolors.WARNING,txt1_1,["0","1"],"LCDBellOn")
                        if oo=="1":
                            txt = txt+f" не пройден. Концевик {cover_cur} зажат, а символ елочка светится." 
                            print(f"{bcolors.FAIL}\n   Концевик {cover_cur} зажат, а символ елочка светится.{bcolors.ENDC}")
                            txt1_2=f"При зажатом концевике крышки клеммника {cover_cur} светится символ 'елочка'."
                            rep_err_list.append(txt1_2)
                            clipboard_err_list.append(txt1_2)
                            test_cover_ok="0"
                            break

            txt1="Тест концевиков крышек клеммников"
            if meter_type == 'i-prom.3' or meter_type == 'i-prom.3T':
                txt1="Тест концевика крышки клеммника"
            if test_cover_ok=="1":
                txt=txt+" успешно пройден"
                print(f"\n    {bcolors.OKGREEN}{txt1} успешно пройден.{bcolors.ENDC}")
            elif test_cover_ok=="0":
                print(f"\n    {bcolors.FAIL}{txt1} не пройден.{bcolors.ENDC}")
            elif test_cover_ok=="2":
                print(f"\n    {bcolors.FAIL}{txt1} пропущен.{bcolors.ENDC}")
                a_txt="Тест концевиков крышек клеммников пропущен."
                rep_err_list.append(a_txt)
                rep_remark_list.append(a_txt)
            with open(default_filename_full, "a", errors="ignore") as file:
                file.write(f"{txt}" + "\n")
        
        
            txt="14. Тест переключателя блокировки реле и работы самого реле:"
            if meter_presence_relay!= "да":
                txt=txt+" для данной модели ПУ не проводится."
                print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")
            else:
                txt=txt+"\n"
                cur_test_skip=False

                current_output_1_dict = {True: "замкнуто", False: "разомкнуто"}

                rele_fact = current_output_1_dict[disconnect_output_state]
                txt1 = "    До проверки исправности переключателя блокировки реле:\n" \
                    "     - фактическое состояние реле: " + rele_fact
                txt = txt+txt1
                current_output_2_dict = {1: "замкнуто", 0: "разомкнуто", 2: "готов к включению от кнопки МЕНЮ"}
                rele_progr = current_output_2_dict[current_output_status]
                print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}", end="")
                color1=bcolors.OKGREEN
                if (rele_fact == "замкнуто" and rele_progr == "разомкнуто") or \
                    (rele_fact == "разомкнуто" and rele_progr == "замкнуто") :
                    color1 = bcolors.FAIL
                    txt1_2 = "Не совпадают статусы положения реле: физически реле "+rele_fact+", " \
                        "а программный статус - " +rele_progr
                txt1="\n     - программное состояние реле: "+ rele_progr
                txt=txt+txt1
                print(f"{color1}{txt1}{bcolors.ENDC}",end="")

                txt1="\n    Режим работы реле: "+str(disconnect_control_status)
                if disconnect_control_status==0:
                    txt1=txt1+" (Запрет управления)"
                    print(f"{bcolors.WARNING}{txt1}{bcolors.ENDC}",end="")
                    txt=txt+txt1
                    cicl1=True
                    while cicl1:
                        try:
                            relay = GXDLMSDisconnectControl("0.0.96.3.10.255")
                            relay.controlMode = 4
                            relay.DataType = 16
                            reader.write(relay, 4)
                            break
                        except Exception as e:
                            oo = communicationTimoutError("Чтение состояния и режима работы реле:", e.args[0])
                            if oo=="0" or oo == "-1":
                                a_dic={"0": ["Ошибка связи с ПУ при чтении состояния реле.", "0"],
                                    "-1": ["Проверка прервана пользователем.", "9"]}
                                ret_txt=a_dic[oo][0]
                                ret_id=a_dic[oo][1]
                                return [ret_id, ret_txt]
                            
                    txt1="    Режим работы ПУ изменен на 4"
                    print(f"{bcolors.WARNING}{txt1}{bcolors.ENDC}")
                    txt=txt+txt1
                else:
                    print(f"{bcolors.OKGREEN}{txt1}{bcolors.ENDC}")
                    txt=txt+txt1
                cicl2=True
                while cicl2:
                    cicl=True
                    assumption=0   #предположим, что переключатель блокировки реле в крайнем левом положении
                    while cicl:
                        if assumption==1:
                            txt1_1 = f"    Переведите переключатель в крайнее {bcolors.ATTENTIONBLUE} " \
                                f"ЛЕВОЕ {bcolors.ENDC} положение и нажмите Enter.\n    {txt_skip_test}\n" \
                                f"    {txt_break_test}"
                            oo=questionSpecifiedKey(bcolors.OKBLUE,txt1_1,["\r","8","/"],"ChangeSwitchOnLeft")
                            if oo=="8":
                                cicl=False
                                cicl2=False
                                cur_test_skip=True
                                break
                            elif oo=="/":
                                testBreak(txt,"Проверка прервана пользователем",default_filename_full, employees_name)
                                return ["9", "Проверка прервана пользователем."]
                        else:
                            print(f"    Считаем, что переключатель блокировки реле установлен в крайнем левом положении")
                        print(f"\n    ",end="")
                        pause_ui(1)
                        tr1=testSwitchDisable()
                        if tr1==1:
                            tr1=testDisplayRele()
                            if tr1!="":
                                txt=txt+"\n    "+tr1
                                rep_err_list.append(tr1)
                                clipboard_err_list.append(tr1)
                                cicl2=False
                            break
                        elif tr1==3:
                            txt1_1="Неисправен переключатель блокировки реле: постоянно разомкнут."
                            txt1="    Тест не пройден. "+txt1_1
                            print(f"{bcolors.FAIL}{txt1}{bcolors.ENDC}")
                            rep_err_list.append(txt1_1)
                            clipboard_err_list.append(txt1_1)
                            cicl=False
                            cicl2=False
                            break
                        elif tr1==4:
                            return ["0", "Ошибка связи с ПУ."]
                        assumption=1    #оказалось, что переключатель в крайнем правом положении или не сработал
                    if cicl2==False:
                        break
                    print()
                    cicl1=True
                    while cicl1:
                        txt1_1 = f"    Переведите переключатель в крайнее {bcolors.ATTENTIONBLUE} " \
                            f"ПРАВОЕ {bcolors.ENDC} положение и нажмите Enter.\n    " + \
                            txt_skip_test+"\n    "+txt_break_test
                        oo=questionSpecifiedKey(bcolors.OKBLUE,txt1_1,["\r","8","/"],"ChangeSwitchOnRight")
                        if oo=="8":
                            cicl2=False
                            cur_test_skip=True
                            break
                        elif oo=="/":
                            testBreak(txt,"Проверка прервана пользователем",default_filename_full, employees_name)
                            return ["9", "Проверка прервана пользователем."]
                        print(f"\n    ",end="")
                        pause_ui(3)
                        tr1=testSwitchEnable()
                        if tr1==1:
                            tr1=testDisplayRele()
                            if tr1!="":
                                txt=txt+"\n    "+tr1
                                rep_err_list.append(tr1)
                                clipboard_err_list.append(tr1)
                                cicl2=False
                                break
                            cicl1=True
                            while cicl1:
                                try:
                                    disconnect_control = GXDLMSDisconnectControl("0.0.96.3.10.255")
                                    disconnect_control_state = reader.read(disconnect_control, 3) #программное состояние реле
                                    break
                                except Exception as e:
                                    oo = communicationTimoutError("Чтение положения реле:", e.args[0])
                                    if oo=="0" or oo == "-1":
                                        a_dic={"0": ["Ошибка связи с ПУ при чтении положения реле.", "0"],
                                            "-1": ["Проверка прервана пользователем.", "9"]}
                                        ret_txt=a_dic[oo][0]
                                        ret_id=a_dic[oo][1]
                                        return [ret_id, ret_txt]
                            if disconnect_control_state==0:
                                print(f"\n    Переведем реле в замкнутое состояние.")
                                tr1=testSwitchEnable()
                                if tr1==1:
                                    txt1="    Тест пройден."
                                    print(f"{bcolors.OKGREEN}{txt1}{bcolors.ENDC}")
                                    txt=txt+"\n"+txt1
                                    cicl2=False
                                    break
                                elif tr1==3:
                                    txt1_1="Не удалось перевести реле в замкнутое состояние."
                                    txt1="    Тест не пройден. "+txt1_1
                                    print(f"{bcolors.FAIL}{txt1}{bcolors.ENDC}")
                                    txt=txt+txt1
                                    rep_err_list.append(txt1_1)
                                    clipboard_err_list.append(txt1_1)
                                    cicl2=False
                                    break
                                elif tr1==4:
                                    return ["0", "Ошибка связи с ПУ."]
                            else:        
                                txt1="\n    Тест пройден."
                                print(f"{bcolors.OKGREEN}{txt1}{bcolors.ENDC}", end="")
                                txt=txt+txt1
                                cicl2=False
                                break
                        elif tr1==3:
                            txt1_1="Неисправен переключатель блокировки реле: постоянно замкнут."
                            txt1="    Тест не пройден. "+txt1_1
                            print(f"{bcolors.FAIL}{txt1}{bcolors.ENDC}")
                            txt=txt+"\n"+txt1
                            rep_err_list.append(txt1_1)
                            clipboard_err_list.append(txt1_1)
                            cicl2=False
                            break
                        elif tr1==4:
                            return ["0", "Ошибка связи с ПУ."]
                    if cicl2==False:
                        break  
                if cur_test_skip==True:
                    txt1="\n    Тест пропущен."
                    txt=txt+txt1
                    txt1_1="Тест переключателя блокировки реле и работы самого реле пропущен."
                    rep_err_list.append(txt1_1)
                    rep_remark_list.append(txt1_1)
                    print(f"\n    {bcolors.FAIL}{txt1_1}{bcolors.ENDC}")
                if test_go==False:
                    break
                
                while True:
                    try:
                        disconnect_control = GXDLMSDisconnectControl("0.0.96.3.10.255")
                        disconnect_output_state = reader.read(disconnect_control, 2)
                        current_output_status = reader.read(disconnect_control, 3)
                        break
                    except Exception as e:
                        oo = communicationTimoutError(
                            "Чтение состояния и режима работы реле:", e.args[0])
                        if oo == "0":
                            toCloseConnectOpto()
                            testBreak(txt,"Проверка прервана пользователем",default_filename_full, employees_name)
                            return ["9", "Проверка прервана пользователем."]

                        elif oo == "-1":
                            toCloseConnectOpto()
                            return ["0", "Ошибка связи с ПУ при получении информации "
                                "о положении реле после теста переключателя реле."]
                
                if test_go==False:
                    break
                
                rele_fact = current_output_1_dict[disconnect_output_state]
                txt1="\n    После проверки исправности переключателя блокировки реле:\n" \
                    "     - фактическое состояние реле: " + rele_fact
                txt = txt+txt1
                rele_progr = current_output_2_dict[current_output_status]
                print(f"{bcolors.OKGREEN}{txt1}{bcolors.ENDC}", end="")
                color1=bcolors.OKGREEN
                if (rele_fact == "замкнуто" and rele_progr == "разомкнуто") or \
                    (rele_fact == "разомкнуто" and rele_progr == "замкнуто") :
                    color1 = bcolors.FAIL
                    txt1_2 = "Не совпадают статусы положения реле: физически реле "+rele_fact+", " \
                        "а программный статус - " +rele_progr
                    rep_err_list.append(txt1_2)
                    clipboard_err_list.append(txt1_2)
                txt1=f"\n     - программное состояние реле: {rele_progr}"
                txt=txt+txt1
                print(f"{color1}{txt1}{bcolors.ENDC}\n",end="")

            fileWriter(default_filename_full, "a", "", f"{txt}\n", \
                "Сохранение в отчет результата теста блокировки реле",join="on")


            res=innerTestGSMModem("15")
            if res[0] in ["0","2", "4", "9"]:
                a_dic={"0": ["Ошибка.", "0"],
                    "2": ["ПУ отправлен в ремонт.", "2"],
                    "4": ["Ошибка связи с ПУ при проведении теста МС.", "0"],
                    "9": ["Проверка прервана пользователем.", "9"]}
                ret_txt=a_dic[res[0]][0]
                ret_id=a_dic[res[0]][1]
                return [ret_id, ret_txt]
            
            elif res[0] in ["1", "3"]:
                a_dic=res[2]
                gsm_docked_tn_sutp=a_dic["gsm_docked_tn_sutp"]
                gsm_docked_sn_sutp=a_dic["gsm_docked_sn_sutp"]
                gsm_docked_model_sutp=a_dic["gsm_docked_model_sutp"]
                gsm_docked_soft_sutp=a_dic["gsm_docked_soft_sutp"]
            

            if connection_initialized:
                toCloseConnectOpto()
                connection_initialized = False


            txt=f"{bcolors.OKGREEN}16. Тест пульта дистанционного управления счетчиком:" \
                f"{bcolors.ENDC}\n"
            if rc_serial_number=="":
                txt=txt+f"    {bcolors.OKGREEN}ПДУ не предусмотрен.{bcolors.ENDC}"
            else:
                txt=txt+f'    {bcolors.OKGREEN}Серийный номер пульта: {rc_serial_number}{bcolors.OKGREEN}'
            print (f"{txt}{bcolors.ENDC}")
            fileWriter(default_filename_full, "a", "", f"{txt}\n", \
                "Сохранение в отчет информации о начале теста " \
                "ПДУ.",join="on")
            
            
            if rc_serial_number!="":
                rc_soft="3.0"

            if rc_serial_number!="":
                while True:
                    txt1=f'{bcolors.OKBLUE}Проверьте работоспособность ПДУ.' \
                        f'{bcolors.ENDC}\n{bcolors.OKBLUE}' \
                        f'Версия ПО данного ПДУ 3.0. Верно? (0-нет, 1-да):{bcolors.ENDC}'
                    oo=questionSpecifiedKey(bcolors.OKBLUE, txt1, ["0", "1"], "", 1)
                    if oo=="1":
                        txt="    Версия ПО ПДУ является актуальной: 3.0."
                        print(f'{bcolors.OKGREEN}{txt}{bcolors.ENDC}')
                        break
                    
                    txt1=f"\n{bcolors.OKGREEN}Введите номер версии данного ПДУ.{bcolors.ENDC}\n" \
                        f'{bcolors.OKGREEN}Чтобы отменить ввод - нажмите "/"{bcolors.ENDC}'
                    oo=inputSpecifiedKey("", txt1, "", [0], ["/"], 0)
                    if oo!="/":
                        if oo=="3.0":
                            txt1=f'{bcolors.OKBLUE}Версия ПО данного ПДУ 3.0. " \
                                f"Верно? (0-нет, 1-да):{bcolors.ENDC}'
                            oo=questionSpecifiedKey("", txt1, ["0", "1"], "", 1)
                            if oo=="0":
                                continue
                            else:
                                txt="Версия ПО ПДУ является актуальной: 3.0."
                                print(f'\n{bcolors.OKGREEN}{txt}{bcolors.ENDC}')
                                break
                        rc_soft=oo
                        txt=f"Текущая версия ПО ПДУ {oo} не соответствует актуальной" \
                            "версии 3.0."
                        print(f'{bcolors.FAIL}{txt}{bcolors.ENDC}')
                        rep_err_list.append(txt)
                        clipboard_err_list.append(txt)
                        break
                fileWriter(default_filename_full, "a", "", f"{txt}\n", \
                    "Сохранение в отчет результата теста ПДУ",join="on")



            if len(rep_err_list)>0:
                a_err_list=rep_err_list.copy()
                res = insHiphenColor(a_err_list, "- ",
                    bcolors.WARNING)
                a_err_txt="\n".join(res[2])
                a_txt=f"{bcolors.WARNING}В ходе проверки ПУ были выявлены следующие замечания:{bcolors.ENDC}" \
                    f"\n{a_err_txt}\n"
                print (a_txt)
            a_txt=""
            header=f"{bcolors.OKGREEN}Внесение дополнительных замечаний, выявленных в ходе " \
                f"проведения проверки ПУ.{bcolors.ENDC}"
            res=innerSelectActions(header, [], "8", err_no_edit_list=rep_err_list)
            if res[0]=="9":
                return ["9", "Проверка прервана пользователем."]
            
            elif res[0]=="3" and len(res[2])>0 :
                a_txt="\n".join(res[2])
                rep_err_list.extend(res[2])
                clipboard_err_list.extend(res[2])
            txt1="17. Прочие замечания:"
            if a_txt=="":
                txt1=txt1+" нет"
                print(f"{bcolors.OKGREEN}{txt1}{bcolors.ENDC}")
            else:
                txt1=f"{txt1}\n{a_txt}"
                print(f"{bcolors.FAIL}{txt1}{bcolors.ENDC}")
            
            fileWriter(default_filename_full, "a", "", f"{txt1}\n", \
                "Сохранение в отчет прочих замечаний",join="on")
            
            break
                

    rep_err_list=delItemList(rep_err_list)
    rep_remark_list=delItemList(rep_remark_list)
    clipboard_err_list=delItemList(clipboard_err_list)
    meter_grade="4"
    rep_remark_txt="\n".join(rep_remark_list)
    clipboard_err_txt="\n".join(clipboard_err_list)

    ret_id="1"
    ret_txt="Проверка ПУ пройдена успешно."
    
    if read_mode_opto=="отчет о конфигурации":
        txt1=f"3.3. Версия ПО прибора учета, полученная из ПУ: {meter_soft}"
        printGREEN(txt1)

        fileWriter(default_filename_full, "a", "", f"{txt1}\n", 
            "Сохранение в отчет информации о версии ПО ПУ ", join="on")


    txt="\n\nЗаключение: "
    print(f"{bcolors.OKGREEN}{txt}{bcolors.ENDC}")
    if len(rep_err_list)>0:
        ret_id="3"
        ret_txt="Имеются замечания к ПУ."

        a_err_list=rep_err_list.copy()
        for i in range(0, len(a_err_list)):
            a_err=a_err_list[i]
            a_err=f"{bcolors.FAIL}{a_err}{bcolors.ENDC}"
            a_err_list[i]=a_err
        a_err_txt="\n".join(a_err_list)
        meter_grade="0"
        txt1=f"ПУ № {meter_tech_number} " \
            f"не соответствует требованиям:\n{a_err_txt}"
        

        printFAIL(f"{txt1}")

    else:
        txt1 = "ПУ № "+meter_tech_number+" соответствует требованиям."

        if read_mode_opto=="отчет о конфигурации":
            txt1="Конфигурация "+txt1
            if meter_config_check=="0":
                txt1="Проверка конфигурации ПУ отключена."
                printWARNING(txt1)
                txt1=""

        printGREEN(txt1)
        
    txt=txt+txt1
    
    if rep_remark_txt!="":
        txt1=f"Примечания:\n{rep_remark_txt}"
        print(f"{bcolors.WARNING}{txt1}{bcolors.ENDC}")
        txt=txt+"\n\n"+txt1

    txt1="\n\nПроверка проведена: "+employees_name
    pc_time = toformatNow()[2]
    dt=datetime.strptime(test_start_time,"%d.%m.%Y %H:%M:%S")
    duration_test = str(int(abs(pc_time - dt).seconds)/60)
    txt1=txt1+"\nПродолжительность проверки, мин: "+duration_test+"\n"
    txt=txt+txt1
    
    fileWriter(default_filename_full, "a", "", f"{txt}\n", 
        "Сохранение в отчет заключения, примечания, инф. "
        f"о сотруднике и продолжительности проверки",join="on")

    next_stage="1"   # ПУ отправить на поверку
    if len(rep_err_list)>0:
        next_stage="0"   #ПУ отправить в ремонт


    reestr_dic=innerSetValReestr()


    if sutp_to_save[0]=="0" and meter_tech_number!=None and \
        meter_tech_number!="" and meter_tn_source=="наклейка" \
        and read_mode_opto!="отчет о конфигурации":
        a_dic={"sutp_to_save": "автоматически: годен/брак," \
            "автоматически: брак"}
        res=checkAvailableAct(employee_id, a_dic)
        if res[0]=="1":
            res=getSutpToSaveDescript(sutp_to_save)
            printColor (f"{res[3]}{res[2]}")
            txt1_1="Включить автоматическое сохранение результата " \
                "проверки в СУТП:годен/брак? 0-нет, 1-да"
            spec_keys=["0","1"]
            oo = questionSpecifiedKey(bcolors.WARNING,txt1_1,spec_keys,"",1)
            print()
            if oo=="1":
                sutp_to_save="2"
                res=getSutpToSaveDescript(sutp_to_save)
                printColor (f"{res[3]}{res[2]}")


    a_save_mode="1"

    if read_mode_opto=="отчет о конфигурации":
        a_save_mode="2"

    res= toSaveResultExtInspection(meter_serial_number, meter_tech_number,
        clipboard_err_txt, employees_name, employee_id, default_filename_full,
        default_dirname, work_dirname, dirname_sos, filename_report, 
        sutp_to_save, data_exchange_sutp, a_save_mode, next_stage, workmode, 
        reestr_dic, rep_copy_public, meter_config_check, config_send_mail, 
        rep_err_send_mail, no_data_in_SUTP_send_mail, rep_remark_txt)
    
    rep_err_list.clear()
    clipboard_err_list.clear()
    rep_remark_list.clear()
    default_value_dict = writeDefaultValue(default_value_dict)
    saveConfigValue('opto_run.json',default_value_dict)

    return [ret_id, ret_txt]

  

def questionBreakTest(meter_position_cur: int):

    a_txt="Проверка прервана."
    ret_id="9"
    ret_txt="Проверка всех ПУ прервана."
    
    txt1=f"Выберите дальнейшее действие для ПУ на поз.{meter_position_cur+1}:"
    menu_item_list=["Повторно запустить текущий этап проверки для ПУ",
        "Перейти к следующей позиции на стенде",
        "Прервать проверку всех оставшихся ПУ на стенде"]
    menu_id_list=["повторно", "исключить ПУ", "прервать"]
    oo=questionFromList(bcolors.OKBLUE, txt1, menu_item_list, menu_id_list,
        "", [], [], [], 1, 1, 1)
    print()
    
    if oo=="исключить ПУ":
        a_txt=f"Проверка текущего ПУ на позиции {meter_position_cur+1} прервана.\n" \
            "Он будет исключен из дальнейшей проверки."
        ret_id="91"
        ret_txt="Исключить позицию ПУ на стенде из проверки."

    elif oo=="повторно":
        a_txt=f"Повторно запускаю текущий этап проверки ПУ на позиции " \
            f"{meter_position_cur+1}"
        ret_id="1"
        ret_txt="Повторная проверка ПУ."

    printWARNING (a_txt)

    return [ret_id, ret_txt]



if  __name__ ==  '__main__' :   
    
    global default_value_dict
    global actual_version_dict
    global rep_err_list

    default_value_dict={}

    a_mode="0"

    n = len(sys.argv)

    if n>1:
        a_mode=sys.argv[1]

    a_dic={"0": "полная проверка", "1": "отчет о конфигурации"}

    read_mode_opto=a_dic.get(a_mode, "полная проверка")

    res = toCheckModuleVersions()
    if res!=1:
        txt1="Нажмите любую клавишу"
        questionOneKey(bcolors.WARNING,txt1)
        sys.exit()
        

    actual_version_dict=[
        ["device_1", [["4.10.276", ["1.9.5.1","1.13.1"]]]],
        ["device_3", [["4.10.276", ["1.9.5.1","1.13.1"]]]], 
        ["device_3T",[["4.10.6.6", ["1.8.6.2"]],
                      ["4.10.276", ["1.9.5.1","1.13.1"]],
                      ["4.10.278", ["1.9.5.1"]]
                    ]
        ]
        ]


    _,_, file_name = getUserFilePath('device_version.json')
    if file_name=="":
        sys.exit()

    with open(file_name, "r", errors="ignore",encoding='utf-8') as file:
        content=json.load(file)
        actual_version_dict=content

    readActualVersionValue()

    default_value_dict=optoRunVarRead()
    readDefaultValue(default_value_dict)

    number_of_meters=default_value_dict["number_of_meters"]
    meter_tech_number_list_last=default_value_dict["meter_tech_number_list"]
    meter_status_test_list_last=default_value_dict["meter_status_test_list"]
    meter_serial_number_list_last=default_value_dict["meter_serial_number_list"]

    workmode=default_value_dict["workmode"]

    sutp_to_save_all=default_value_dict["sutp_to_save"]
    meter_config_check_all=default_value_dict["meter_config_check"]

    res=getPathOptoRun(workmode)
    if res[0]=="":
        sys.exit()
    
    opto_run_path =res[2]
    multi_config_dir=res[3]   
    
    opto_run_path_last=f"{opto_run_path}_last"
    res=copyFile(opto_run_path, opto_run_path_last, "0", "1")
    if res[0]=="0":
        a_err_txt="Ошибка при дублировании ф.'opto_run.json'."
        printFAIL(a_err_txt)
        keystrokeEnter()
        sys.exit()

    comport_active_dic={}
    
    i=0
    while i<len(meter_tech_number_list_last):
        a_tech=meter_tech_number_list_last[i]
        if a_tech==None or a_tech=="":
            i+=1
            continue
        
        if number_of_meters>1:
            a_sep="="*44+"="*len(a_tech)+ \
                "="*len(str(i+1))
            a_txt=f"{a_sep}\nПроверка ПУ №{a_tech}, " \
                f"установленного на позиции {bcolors.ATTENTIONGREEN} " \
                f"{i+1} {bcolors.ENDC}\n{a_sep}"
            
            printGREEN(a_txt)

            title_new=f"Аппаратная проверка ПУ (текущая поз.: {i+1})"
            replaceTitleWindow("", title_new)

        opto_run_multi=f"opto_run_{i}.json"
        opto_run_multi_path=os.path.join(multi_config_dir, opto_run_multi)
        res=copyFile(opto_run_multi_path, opto_run_path, "0", "1")
        if res[0]=="0":
            a_err_txt=f"Ошибка при копировании ф.'{opto_run_multi}' из папки " \
                "'multi_config' в рабочую папку."
            printFAIL(a_err_txt)
            keystrokeEnter()
        
        default_value_dict=optoRunVarRead()
        readDefaultValue(default_value_dict)

        sutp_to_save=sutp_to_save_all
        meter_config_check=meter_config_check_all
        
        meter_tech_number_list=meter_tech_number_list_last.copy()
        meter_status_test_list=meter_status_test_list_last.copy()
        meter_serial_number_list=meter_serial_number_list_last.copy()

        default_value_dict = writeDefaultValue(default_value_dict)

        res=saveConfigValue("opto_run.json", default_value_dict)
        if res[0]=="0":
            a_err_txt="Ошибка при сохранении данных в ф.opto_run.json."
            printFAIL(a_err_txt)
            keystrokeEnter()
            sys.exit()
        
        meter_pw_high_encrypt=default_value_dict['meter_pw_high_encrypt']
        res=cryptStringSec("расшифровать", meter_pw_high_encrypt)
        meter_hight_password=res[2]
        meter_serial_number=default_value_dict['meter_serial_number']
        workmode=default_value_dict['workmode']
        com_current=default_value_dict['com_current']
        com_opto=default_value_dict['com_opto']
        com_rs485=default_value_dict['com_rs485']
        a_dic={"com_opto":com_opto, "com_rs485":com_rs485}
        comport=a_dic[com_current]
        speaker=default_value_dict["speaker"]

        serial_num=""
        if com_current=="com_rs485":
            serial_num=meter_serial_number[-4:]

        if not comport in comport_active_dic:
            res = settingOpt(password=meter_hight_password,
                serial_num=serial_num, comport=comport, msg_print="0", 
                authentication="High")
            if res[0]=="0":
                sys.exit()

            comport_active_dic[comport]=[res[1], res[2]]

        reader=comport_active_dic[comport][0]
        settings=comport_active_dic[comport][0]
            
        if speaker=="1":
            playsound("speech\hello.mp3", block=False)

        res=read_opto(read_mode_opto)

        if res[0] in ["0", "9"]:
            if number_of_meters>1 and (i+1)!=len(meter_tech_number_list_last):

                res=questionBreakTest(i)
                if res[0]=="9":
                    for j in range(i, len(meter_tech_number_list_last)):
                        meter_tech_number_list_last[j]=""
                        meter_serial_number_list_last[j]=""
                        meter_status_test_list_last[j]="пропущен"

                    break

                elif res[0]=="1":
                    continue

            meter_tech_number_list_last[i]=""
            meter_serial_number_list_last[i]=""
            meter_status_test_list_last[i]="пропущен"

        elif res[0] in ["2", "3"]:
            meter_status_test_list_last[i]="ремонт"

        elif res[0]=="1":
            meter_status_test_list_last[i]="годен"

        default_value_dict=writeDefaultValue(default_value_dict)
        res=saveConfigValue("opto_run.json", default_value_dict)
        if res[0]=="0":
            a_err_txt="Ошибка при сохранении данных в ф.opto_run.json."
            printFAIL(a_err_txt)
            keystrokeEnter()
            sys.exit()

        res=copyFile(opto_run_path, opto_run_multi_path, "0", "1")
        if res[0]=="0":
            a_err_txt=f"Ошибка при копировании ф.'{opto_run_multi}' из " \
                "рабочей папки в папку 'multi_config'."
            printFAIL(a_err_txt)
            keystrokeEnter()

        i+=1


    opto_run_path_last=f"{opto_run_path}_last"
    res=copyFile(opto_run_path_last, opto_run_path, "0", "1")
    if res[0]=="0":
        a_err_txt="Ошибка при копировании ф.'opto_run.json_last'."
        printFAIL(a_err_txt)
        keystrokeEnter()
        sys.exit()

    default_value_dict=optoRunVarRead()
    readDefaultValue(default_value_dict)
    
    meter_tech_number_list=meter_tech_number_list_last.copy()
    meter_serial_number_list=meter_serial_number_list_last.copy()
    meter_status_test_list=meter_status_test_list_last.copy()
    
    default_value_dict = writeDefaultValue(default_value_dict)
    saveConfigValue('opto_run.json',default_value_dict)

    res=copyFile(opto_run_path, opto_run_path_last, "0", "1")
    if res[0]=="0":
        a_err_txt="Ошибка при дублировании ф.'opto_run.json'."
        printFAIL(a_err_txt)
        keystrokeEnter()
        sys.exit()

    
    for a_comport_dic in comport_active_dic:
        reader=comport_active_dic[a_comport_dic][0]
        settings=comport_active_dic[a_comport_dic][1]
        reader.close()

    sys.exit()
