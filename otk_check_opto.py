
import os
import sys
import time
import msvcrt
import shutil
import json  # для сохранения значений переменных по умолчанию в файле в формате json
from alive_progress import alive_bar
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
from otk_opto import getAboutOtkOpto, ExchangeBetweenPrograms, optoRunVarRead

from colorama import init

init()

def getAboutOtkCheckOpto():
    version = "04.04.2024 22:01"
    descript = "Программа для проверки связи с ПУ через оптопорт"
    return [version, descript]



def readDefaultValue(value_dict={}):
    keys_list=list(value_dict.keys())
    for key in keys_list:
        globals()[key]=value_dict[key]



def writeDefaultValue(dict={}):
    global default_value_dict

    if len(dict)==0:
        dict = default_value_dict
    keys_list=list(dict.keys())
    for key in keys_list:
        if key in globals():
            dict[key] = globals()[key]
    return dict



def restoreCOMPort(com_name:str):

    global default_value_dict   #словарь значений по умолчанию
    global com_opto             #COM-порт, к которому подключен оптопорт
    global com_rs485            #COM-порт, к которому подключен RS-485


    a_dic={"com_opto":["оптопорт","оптопорта"],
        "com_rs485":["преобразователь RS-485","преобразователя RS-485"]}
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



def toCheckModuleVersions():
    ret = 1
    module_vers_ok_dict = {"otkLib": "04.04.2024 13:56",
                           "otk_opto": "03.04.2024 22:50"}
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


def saveResultCheckPW(result:str):
    global msg_print    #метка вывода сообщений
    dt = str(toformatNow()[3])
    content = {"dateTime": dt, "source": "otk_check_opto",
                "operation": "checkPassword", "result": result}

    res = ExchangeBetweenPrograms(operation="add",
        recipient="otk_menu_result", content=content)
    if res[0] == "0" and msg_print in ["1", "2"]:
        print(
            f"{bcolors.WARNING}При сохранении результатов в файле" \
            f"возникла ошибка: {res[1]}{bcolors.ENDC}\n")
    return



def readMeterVoltage(meter_sn: str):

    voltage1="0"
    voltage2="0"
    voltage3="0"
    amp1="0"
    amp2="0"
    amp3="0"
    
    sheet_name = "Product1"
    res = toGetProductInfo2(meter_sn, sheet_name)
    if res[0] == "0":
        txt1_2 = f"Проверка прервана, т.к. не удалось определить модель ПУ по его номеру."
        print(f"{bcolors.FAIL}{txt1_2}{bcolors.ENDC}")
        toCloseConnectOpto()
        return ["0", "Не удалось определить модель ПУ по его номеру.", "", {}, {}]
    meter_phase=res[18]

    while True:
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

                amp1 = GXDLMSRegister("1.0.31.7.0.255")
                v = reader.read(amp1, 2)
                amp1 = str(v)
                amp2 = GXDLMSRegister("1.0.51.7.0.255")
                v = reader.read(amp2, 2)
                amp2 = str(v)
                amp3 = GXDLMSRegister("1.0.71.7.0.255")
                v = reader.read(amp3, 2)
                amp3 = str(v)
            elif meter_phase == "1":
                voltage1 = GXDLMSRegister("1.0.12.7.0.255")
                v = reader.read(voltage1, 2)
                voltage1 = str(int(v / 1000))

                amp1 = GXDLMSRegister("1.0.11.7.0.255")
                v = reader.read(amp1, 2)
                amp1 = str(v)
            break
        except Exception as e:
            oo = communicationTimoutError(
                "Считывание значения мгновенного напряжения: ", e.args[0])
            if oo == "0" or oo == "-1":
                toCloseConnectOpto()
                return ["0", "Ошибка при чтении значений мгновенного " 
                        "напряжения/тока на ПУ", "", {}, {}]
            
    v_dic={"voltage1":voltage1, "voltage2": voltage2, "voltage3":voltage3}
    amp_dic={"amp1": amp1, "amp2": amp2, "amp3": amp3}
    
    return ["1", "Чтения значений напряжения/тока выполнены успешно.", 
            meter_phase, v_dic, amp_dic]
        


if __name__ == '__main__':

    def innerInitConnectOpto(msg_print="1", check_mode="0", com_current="com_opto"):
        

        a_dic={"0":f"Проверяем пароль: {meter_pw_descript}.",
           "1":f"Считываем серийный номер ПУ, уровни напряжения и версию ПО ПУ."}
        
        while True:
            if msg_print=="1":
                print(f"{bcolors.OKGREEN}{a_dic[check_mode]}{bcolors.ENDC}\n")
            result="0"
            try:
                reader.initializeConnection(msg_print=False)
                result = "1"
                break
            
            except Exception as e:
                if e.args[0] == "Connection is permanently rejected\r\nAuthentication failure.":
                    if msg_print in ["1", "2"]:
                        print(f"{bcolors.FAIL}ПУ отказал в доступе.{bcolors.ENDC}")
                    result="2"
                    break

                elif "Serial port is not open" in e.args[0]:
                    print(f"{bcolors.FAIL}COM-порт устройства интерфейса не найден.{bcolors.ENDC}\n")
                    res=restoreCOMPort(com_current)
                    if res=="1":
                        continue
                    result="4"
                    break

                else:
                    if msg_print in ["1", "2"]:
                        print(f"{bcolors.FAIL}При инициализации канала связи возникла ошибка:"
                            f"{e.args[0]}{bcolors.ENDC}")
                        
            toCloseConnectOpto()
            txt1 = f"Повторить подключение (0- нет, 1-да)?"
            oo = questionSpecifiedKey(bcolors.WARNING, txt1, ["0", "1"],"", 1)
            print()
            if oo == "0":
                break

        return result
    
    
    msg_print="1"
    check_mode="0"
    n = len(sys.argv)
    if n>1:
        msg_print=sys.argv[1]
        check_mode=sys.argv[2]
        

    res = toCheckModuleVersions()
    if res != 1:
        txt1 = "Нажмите любую клавишу"
        questionOneKey(bcolors.WARNING, txt1)
        sys.exit()

    
    default_value_dict = optoRunVarRead()

    meter_pw_encrypt=default_value_dict['meter_pw_encrypt']
    _,_, meter_hight_password = cryptStringSec("расшифровать",meter_pw_encrypt)
    meter_pw_descript = default_value_dict['meter_pw_descript']
    meter_pw_level= default_value_dict['meter_pw_level']
    com_current=default_value_dict['com_current']
    com_opto = default_value_dict[com_current]
    serial_num=""
    meter_soft=""

    if com_current=="com_rs485":
        meter_serial_number=default_value_dict['meter_serial_number']
        serial_num=meter_serial_number[-4:]



    if msg_print=="1" and check_mode=="0":
        print(f"\n{bcolors.OKGREEN}Проверка пароля для подключения к ПУ.{bcolors.ENDC}")

    res = checkComPort(com_opto=com_opto, print_msg="")
    if res == "0":
        print(f"{bcolors.FAIL}COM-порт устройства интерфейса не найден.{bcolors.ENDC}")
        res=restoreCOMPort(com_current)
        if res!="1":
            saveResultCheckPW("4")
            sys.exit()

    
    res = settingOpt(password=meter_hight_password,
        serial_num=serial_num, comport=com_opto, msg_print="0", 
        authentication=meter_pw_level)
    if res[0] == "0":
        reader.close()
        saveResultCheckPW("0")
        sys.exit()
    reader = res[1]
    settings = res[2]

    
    result=innerInitConnectOpto(msg_print, check_mode, com_current)
    
        
    while True:
        if check_mode in ["1", "2"] and result=="1":
            result="0"
            res = toReadDataFromMeter(
                "0.0.96.1.0.255", 2, "Считывание серийного номера ПУ: ", "utf-8")
            if res[0]==0:
                txt1 = f"Повторить запрос к ПУ (0- нет, 1-да)?"
                oo = questionSpecifiedKey(bcolors.WARNING, txt1, ["0", "1"],"", 1)
                print()
                if oo == "0":
                    break
                toCloseConnectOpto()
                result=innerInitConnectOpto(msg_print, check_mode, com_current)
                continue

            meter_sn_ep=res[1]
            default_value_dict['meter_sn_ep']=meter_sn_ep
            saveConfigValue('opto_run.json',default_value_dict)
            result="3"

            if check_mode=="2" and result=="3":
                res=readMeterVoltage(meter_sn_ep)
                if res[0]=="0":
                    result="0"
                else:
                    a_ph=res[2]
                    a_v_dic=res[3]
                    a_amp_dic=res[4]
                    default_value_dict['meter_phase']=a_ph
                    default_value_dict['meter_voltage_dic']=a_v_dic
                    default_value_dict['meter_amperage_dic']=a_amp_dic
                    

                a_gxdlms= GXDLMSData("0.0.96.1.8.255")
                meter_soft = reader.read(a_gxdlms, 2).decode("utf-8")
                default_value_dict['meter_soft']=meter_soft
                saveConfigValue('opto_run.json',default_value_dict)
                result="3"

        
        break
            
    
    saveResultCheckPW(result)
    
    reader.close()
    sys.exit()


    
