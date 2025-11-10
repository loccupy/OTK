# стартовый модуль программы проверки ПУ для получения пароля подключения к ПУ из СУТП


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

from gtts import gTTS  # pip install gTTS для синтеза речи с пом.Google через интернет
from playsound import playsound  # pip install playsound==1.2.2 для воспроизведения звукового файла
from tqdm import tqdm   #pip install tqdm   для отображения progress bar
from prettytable import PrettyTable

from gurux_dlms.objects import GXDLMSClock, GXDLMSData, GXDLMSRegister, GXDLMSDisconnectControl
from gurux_dlms.objects import GXDLMSDisconnectControl, GXDLMSRegister, GXDLMSProfileGeneric
from datetime import datetime, timedelta
from gurux_dlms.objects import GXDLMSClock
from gurux_dlms.enums import DataType, ObjectType
from gurux_dlms import GXDLMSClient, GXTime, GXDateTime
from gurux_dlms.GXDLMSException import *
from docxtpl import DocxTemplate, InlineImage, RichText
from pathlib import Path

from libs.sutpLib import getAboutSutpLib, savetToSUTP2, getNameEmployee, findNameMeterStatus, \
    getInfoAboutDevice,request_sutp,getDeviceHistory,getDeviceRepayHistory, \
    preChecksToGhangeStatusMeter, getMeterAllSN, getDevicePw
from libs.otkLib import *

from otk_opto import getLocalStatistic, getSutpToSaveDescript, toSaveResultExtInspection, \
    saveStatistic, saveStatisticThread, saveStatisticInSharedFolder

init()


def getMeterPassDefault():
    meter_pass_default_dict = {
        "Стандартный высокого уровня": "1234567898765432",
        "Карелия 12.23 высокого уровня": "EkkGsmAdmin2021i",
        "БЭСК 01.24 высокого уровня": "besk000000000000"
    }
    return meter_pass_default_dict



def editUserDefectsFile(employee_id:str, workmode='экплуатация'):
    

    def innerPrintTbl(param_print_txt:str, defects_user_txt:str,
        header_txt=""): 

        rtbl=Table(Column(header="Наименование группы", justify="left", 
                min_width=12, style="green", header_style="green"),
            Column(header="Описания дефектов в группе", justify="left", 
                style="green", header_style="green"),
            style="green", show_lines=True)
        rtbl.add_row(param_print_txt, defects_user_txt)

        console=Console()
        console.clear()
        os.system("CLS")
        if header_txt!="":
            rtbl.title=header_txt
            rtbl.title_style="green"
            rtbl.title_justify="left"
        console.print(rtbl)
        return
        
    
    edit_mode="user"
    group_filter = employee_id

    param_num_edit=-1
    cicl1=True
    while cicl1:
        file_name_dic={"alluser": "defects_alluser_dic.json",
            "user":"defects_user_dic.json"}
        file_name=file_name_dic[edit_mode]
        res = readGonfigValue(file_name, [], {}, workmode)
        if res[0] == "0":
            return
        defect_user_dic = res[2].get(group_filter, {})
        if len(defect_user_dic)==0:
            defect_user_dic = {"прочие":[]}
            a_dic = {group_filter: defect_user_dic}
            res = saveConfigValue(file_name, a_dic, workmode)
            if res[0]=="0":
                return
        param_list=list(defect_user_dic.keys())

        param_descript_dic={"сн ПУ": "серийный номер ПУ", 
            "тн ПУ": "технический номер ПУ", 
            "сн МС": "серийный номер МС", 
            "тн МС": "технический номер МС"}
        param_print_list=param_list.copy()
        param_question_list=param_list.copy()

        for i in range(0, len(param_print_list)):
            param_name=param_print_list[i]
            if param_name in param_descript_dic:
                param_name=param_descript_dic[param_name]
            param_question_list[i]=param_name
            param_name=f"{i+1}. {param_name} "
            param_print_list[i] = param_name

        param_print_list_main=param_print_list.copy()

       
        if param_num_edit==-1:
            param_num=0
        else:
            param_num=param_num_edit
            param_num_edit=-1

        while True:
            param_print_list=param_print_list_main.copy()
            param_name=param_list[param_num]
            param_name_descript=param_name
            if param_name in param_descript_dic:
                param_name_descript=param_descript_dic[param_name]
            param_print_list[param_num]=f'{param_num+1}.[{param_name_descript}]'
            defect_user_dic_list = defect_user_dic.get(param_name,[])
            defects_user_list=[]

            a_count=0
            for a_dic in defect_user_dic_list:
                a_descript=a_dic["defects"]
                a_descript=f"{a_count+1}) {a_descript}"
                defects_user_list.append(a_descript)
                a_count+=1
            
            param_print_txt = "\n".join(param_print_list)

            defects_user_txt="\n".join(defects_user_list)

            header_txt=f'Для пользователя с табельным номером {group_filter} ' \
                f'найдены следующие описания дефектов:'
            if edit_mode!="user":
                header_txt = f'Редактируем ф."{file_name}".\n' \
                    f'Для группы дефектов "{group_filter}" ' \
                    f'найдены следующие описания:'
            innerPrintTbl(param_print_txt, defects_user_txt, header_txt)
        
            txt1=f'{bcolors.OKBLUE}Выберите наименование группы дефектов:' \
                f'{bcolors.ENDC}'
            list_txt = param_question_list
            list_id=param_list
            cur_id=param_name
            speс_keys_hidden = ["#mode"]
            spec_list=["редактировать", "выйти"]
            spec_keys=["\r", "/"]
            spec_id = spec_list
            oo=questionFromList(bcolors.OKBLUE, txt1, list_txt,
                list_id, cur_id, spec_list,spec_keys,spec_id, 1, 0, 1,
                speс_keys_hidden)
            
            if oo=="выйти":
                return
            
            elif oo=="редактировать":
                param_num_edit=param_num
                mandatory_filter = param_name
                user_filter = group_filter
                a_dic={"alluser":"редактирование all",
                       "user": "редактирование user"}
                mode_1 = a_dic[edit_mode]
                txt1=f'{bcolors.OKGREEN}Редактирование описаний дефектов для группы ' \
                    f'"{bcolors.ENDC}{bcolors.OKRESULT}{param_name_descript}' \
                    f'{bcolors.ENDC}"{bcolors.OKGREEN}.{bcolors.ENDC}\n'
                res=inputResultExtInspection(txt1, "1", [],[], mandatory_filter, 
                    user_filter, workmode, mode_1)
                if res[0]=="0":
                    break
 
            elif oo=="#mode":
                txt1=f'\n\n{bcolors.OKBLUE}Выберите режим редактирования:' \
                f'{bcolors.ENDC}'
                list_txt = list(file_name_dic.keys())
                list_id = list_txt
                cur_id=edit_mode
                spec_list=["выйти"]
                spec_keys=["/"]
                spec_id = spec_list
                oo=questionFromList(bcolors.OKBLUE, txt1, list_txt,
                    list_id, cur_id, spec_list,spec_keys,spec_id, 1, 0, 1,
                    speс_keys_hidden)
                if oo=="выйти":
                    continue
                else:
                    edit_mode_1=oo
                    file_name = file_name_dic[edit_mode_1]
                    res = readGonfigValue(file_name, [], {}, workmode)
                    if res[0] == "0":
                        return
                    a_dic = res[2]
                    txt1=f'\n\n{bcolors.OKBLUE}Выберите группу дефектов для ' \
                        f'редактирования:{bcolors.ENDC}'
                    list_txt = list(a_dic.keys())
                    list_id = list_txt
                    cur_id=""
                    spec_list=["выйти"]
                    spec_keys=["/"]
                    spec_id = spec_list
                    oo=questionFromList(bcolors.OKBLUE, txt1, list_txt,
                        list_id, cur_id, spec_list,spec_keys,spec_id, 1, 0, 1,
                        speс_keys_hidden)
                    if oo=="выйти":
                        continue
                    else:
                        edit_mode=edit_mode_1
                        group_filter=oo
                        break

            else:
                param_num=param_list.index(oo)



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
            txt1_1=(f'{bcolors.OKGREEN}Указаны следующие дефекты:' 
                    f'{bcolors.ENDC}')
            if "редактирование" in mode:
                txt1_1=(f'{bcolors.OKGREEN}Список вариантов дефектов:' 
                    f'{bcolors.ENDC}')
            i_num=1
            for a in defects_list:
                txt1_1=f'{txt1_1}\n{bcolors.OKGREEN}{i_num}. {a}{bcolors.ENDC}'
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
            list_id, "",spec_list,spec_keys,spec_id, 1, 1, 1,
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
            i_num=abs(int(oo))-1
            defect_descript_del=defects_list[i_num]
            txt1=f"\n{bcolors.OKBLUE}Отредактируйте замечание:{bcolors.ENDC}"
            oo=inputSpecifiedKey("",txt1,"",[0],["/"],0, defect_descript_del)
            if oo=="/":
                continue
            defects_list[i_num]=oo
            continue
        
        elif oo[0]=="-" and oo[1]=="-":
            i_num=int(oo[2:])-1
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

    global default_value_dict

    default_value_dict=getDefaultValue()

    res=readGonfigValue(file_name_in="opto_run.json",
        var_name_list=[],default_value_dict=default_value_dict)
    if res[0]!="1":
        txt1=f"{bcolors.WARNING}При формировании конфигурационных значений " \
            f"для работы программы возникла ошибка.{bcolors.ENDC}\n" \
            f"{bcolors.OKBLUE}Нажмите Enter{bcolors.ENDC}"
        print (txt1)
        oo=questionSpecifiedKey("","",["\r"],"",1)
        return ["0", "Ошибка при получении значений по умолчанию.", {}]
    default_value_dict=res[2]
    return default_value_dict



def changeDefaultValue():
    global var_all_value_dic    #словарь со всеми предлагаемыми вариантами значений для
    
    global default_value_dict
    
    global employees_name       #ФИО пользователя
    global employee_id          # таб. номер пользователя
    global employee_pw_encrypt  # зашифрованный пароль пользователя
    global rep_copy_public      #метка возможности копирования протокола в общую папку:
    global speaker              #метка вкл/откл диктора. "0"-откл, "1"-вкл
    global res_ext_at_begin_test #метка записи результата внешнего осмотра ПУ в начале 
    global modem_status         #статус модема по умолчанию: 0-не будет устанавливаться, 1-рабочий, 2-тестовый
    global actions_no_mc        # действия при отсутствии модуля связи:
    global SIMcard_status       #Статус SIM-карты по умолчанию: 0-не будет устанавливаться, 1-рабочая, 2-тестовая
    global meter_color_body     #цвет корпуса ПУ по умолчанию
    global meter_adjusting_clock  #корректировка часов счетчика: 
    global meter_config_check   #метод проверки конфигурации ПУ: "0"-откл.,
    global config_send_mail     #отправка сообщения по электронной почте о
    global rep_err_send_mail    #отправка сообщения по электронной почте о
    global no_data_in_SUTP_send_mail    #отправка сообщения по электронной почте о
    global com_config_current_select    #способ выбора COM-порта для проверки
    global com_config_user      # интерфейс, зафиксированный пользователем для проверки
    global electrical_test_circuit  #схема подключения ПУ для проверки: 
    global ctrl_current_electr_test #контрольное значение тока в схеме подключения ПУ
    
    global meter_type_def           #тип (модель) ПУ
    global modem_type_def           #тип модуля связи
    global sutp_to_save             #способ записи рез-тов теста в БД СУТП ("0"-отключен, "1"-ручной,
    global order_control        # метка контроля принадлежности ПУ определенному заказу: ("0"-отключен, "1"-включен)
    global order_control_descript   # номер и описание контролируемого заказа
    global data_exchange_sutp       # метка обмена данными с СУТП:"0"-откл.,"1"-вкл.
    global speaker                  #метка использования диктора:"0"-откл, "1"-вкл
    global print_number_big_font    # метка печати на экране вводимых номеров изделий

    readDefaultValue()

    meter_typegroup_def="спрашивать каждый раз"
    if meter_type_def=="" or meter_type_def==None:
        meter_type_def="спрашивать по умолчанию"

    modem_typegroup_def="спрашивать каждый раз"
    if modem_type_def=="" or modem_type_def==None:
        modem_type_def="спрашивать по умолчанию"


    cicl=True
    while cicl:
        os.system("CLS")
        employee_id_old = employee_id
        rep_copy_public_old= rep_copy_public
        meter_typegroup_def_old = meter_typegroup_def
        meter_type_def_old = meter_type_def
        meter_color_body_old = meter_color_body
        meter_adjusting_clock_old = meter_adjusting_clock
        meter_config_check_old=meter_config_check
        modem_status_old = modem_status
        modem_type_def_old = modem_type_def
        SIMcard_status_old = SIMcard_status
        data_exchange_sutp_old = data_exchange_sutp
        sutp_to_save_old = sutp_to_save
        speaker_old = speaker
        order_control_old=order_control
        electrical_test_circuit_old=electrical_test_circuit
        ctrl_current_electr_test_old=ctrl_current_electr_test
        print_number_big_font_old=print_number_big_font

        toPrintDefaultValue()
        txt1="Выберите группу значений по умолчанию, которые желаете изменить"
        list_txt=["Информация о пользователе", \
            "Информация о ПУ", "Информация о модуле связи", 
            "Настройка взаимодействия с СУТП", "Контроль принадлежности ПУ заказу",
            "Настройка проверки конфигурации ПУ", "Настройка взаимодействия "
            "с программой 'MassProdAutoConfig.exe'", "Прочие параметры"]
        list_id=["о пользователе","о ПУ","о МС","взаимодействие с СУТП",
                 "контроль заказа", "конфигурация ПУ", 
                 "взаимодействие с MassConfig", "прочие параметры"]
        cur_id=""
        spec_list=["Выход из меню"]
        spec_keys=["/"]
        spec_id=["выход"]           
        oo = questionFromList(bcolors.OKBLUE, txt1, list_txt, list_id, cur_id, \
                    spec_list=spec_list, spec_keys=spec_keys, spec_id=spec_id,
                    start_list_num=1)
        os.system("CLS")

        if oo=="выход":
            return

        elif oo=="о пользователе":
            printGREEN(f"Текущий пользователь: {employees_name}")
            printBLUE("Производим замену пользователя.")
            res=changeUser(workmode)
            if res[0]=="1":
                res = readGonfigValue("opto_run.json", [], {}, workmode, "1")
                if res[0] != "1":
                    continue

                default_value_dict = res[2]
                readDefaultValue(default_value_dict)
                

        elif oo=="о ПУ":
            cicl1=True
            while cicl1:
                sheet_name = "Product1"

                cicl2=True
                while cicl2:
                    os.system("CLS")
                    txt1 = "\nУкажите вид ПУ:"
                    list_txt = ["спрашивать каждый раз", "i-prom.1", "i-prom.3-1", 
                        "i-prom.3-3", "i-prom.3-3T", "i-prom.3-3Z",
                        "i-prom.3Z-3", "i-prom.3Z-3T", "i-prom.DC"]
                    list_id = list_txt
                    spec_list=["определить по серийному номеру ПУ",
                        "следующий параметр", "отмена"]
                    spec_keys=["*","\r","/"]
                    if meter_typegroup_def=="спрашивать каждый раз":
                        spec_list=["определить по серийному номеру ПУ","ok","отмена"]
                        spec_keys=["*","\r","/"]
                    spec_id=spec_list 
                    oo = questionFromList(bcolors.OKBLUE, txt1, list_txt,
                        list_id, meter_typegroup_def,spec_list,spec_keys,spec_id)
                    if oo == "следующий параметр":
                        break
                    elif oo=="отмена":
                        meter_typegroup_def=meter_typegroup_def_old
                        cicl1=False
                        break
                    elif oo=="определить по серийному номеру ПУ":
                        txt1 = "\nВведите или отсканируйте серийный номер ПУ, указанный " \
                                "на корпусе ПУ:" \
                                "\nЧтобы прервать ввод - введите '/' и нажмите Enter."
                        oo=inputNumberDevice([15, 13],"счетчик э/э", sheet_name, txt1,["/"])
                        if oo[0]=="1":
                            num=oo[1]
                        else:
                            continue
                        sheet_name1 = sheet_name
                        res = toGetProductInfo2(num, sheet_name1)
                        if res[0] == "0":
                            continue
                        meter_typegroup_def=res[4]
                        meter_type_def = res[2]


                    elif oo=="ok":
                        if meter_typegroup_def=="спрашивать каждый раз":
                            meter_type_def="спрашивать каждый раз"
                            modem_type_def = "спрашивать каждый раз"

                        break
                    
                    else:
                        meter_typegroup_def=oo
                
                if cicl1==False:
                    break

                if meter_typegroup_def != "спрашивать каждый раз":
                    sheet_name = "Product1"
                    if meter_typegroup_def=="i-prom.3-1":
                        sheet_name = "Product0"
                    product_filter_dict={"i-prom.1":["i-prom.1-1-","i-prom.1-2-","i-prom.1-3-"], \
                        "i-prom.3-1": ["i-prom.3-1-"], "i-prom.3-3": ["i-prom.3-3-"], \
                        "i-prom.3-3T":["i-prom.3-3T-"], "i-prom.3-3Z":["i-prom.3-3Z-"],
                        "i-prom.3Z-3":["i-prom.3Z-3-"], "i-prom.3Z-3T":["i-prom.3Z-3T-"],
                        "i-prom.DC":["i-prom.DC-"]}
                    product_filter_list=product_filter_dict.get(meter_typegroup_def, [])
                    meter_list = toFillListProductModel(product_filter_list, sheet_name)
                    
                    if not meter_type_def in meter_list:
                        meter_type_def="спрашивать каждый раз"
                        
                    cicl2 = True
                    while cicl2:
                        os.system("CLS")
                        if meter_list!="":
                            txt1 = "\nВыберите модель ПУ, которая указана на корпусе ПУ:"
                            list_txt=["спрашивать каждый раз"]+meter_list
                            list_id = list_txt
                            spec_list=["определить сейчас по серийному номеру ПУ","следующий параметр", \
                                "отмена"]
                            spec_id=spec_list
                            spec_keys=["*","\r","/"]
                            oo = questionFromList(bcolors.OKBLUE, txt1, list_txt,
                                list_id, meter_type_def, spec_list,spec_keys,spec_id)
                            if oo == "следующий параметр":
                                break
                            elif oo=="отмена":
                                meter_typegroup_def = meter_typegroup_def_old
                                meter_type_def= meter_type_def_old
                                cicl1=False
                                break
                            elif oo=="определить сейчас по серийному номеру ПУ":
                                txt1 = "\nВведите или отсканируйте номер ПУ, указанный на корпусе ПУ:" \
                                    "\nЧтобы прервать ввод - введите '/' и нажмите Enter."
                                oo=inputNumberDevice([15,13],"счетчик э/э", sheet_name, txt1,["/"])
                                if oo[0]=="1":
                                    num=oo[1]
                                else:
                                    continue
                                sheet_name1 = sheet_name
                                res = toGetProductInfo2(num, sheet_name1)
                                if res[0] == "0":
                                    continue
                                meter_product_type = res[2]
                                if not meter_product_type in meter_list:
                                    txt1 = "По введенному номеру определена модель ПУ '"+meter_product_type+ \
                                        "', но ее нет в списке.\nУстановить эту модель как значение " \
                                        "по умолчанию? (0-нет, 1-да)"
                                    oo = questionSpecifiedKey(bcolors.WARNING,txt1,["0", "1"])
                                    if oo=="0":
                                        continue

                                    meter_list.append(meter_product_type)
                                
                                meter_type_def=meter_product_type
                                
                            else:
                                meter_type_def=oo
                        
                        else:
                            printWARNING("Не удалось сформировать список моделей ПУ")
                            keystrokeEnter()
                            cicl1=False
                            break

                if cicl1==False:
                    break

                cicl2=True
                while cicl2:
                    os.system("CLS")
                    txt1 = "\nУкажите схему подключения ПУ для проведения проверки:"
                    res=readGonfigValue("var_all_value.json", [], {},
                        workmode, "1")
                    if res[0]=="1":
                        a_dic=res[2].get("electrical_test_circuit", {})
                        if len(a_dic)>0:
                            list_txt=list(a_dic["all_value"].keys())
                            list_id=list(a_dic["all_value"].values())
                    
                        if "i-prom.1" in meter_typegroup_def:
                            list_txt = ["напряжение - 1 фаза, ток - нет",
                                        "напряжение - 1 фаза, ток - 1 фаза"]
                            list_id = ["1-0", "1-1"]
                            
                        spec_list=["следующий параметр", "отмена"]
                        spec_id=spec_list
                        spec_keys=["\r","/"]
                        oo = questionFromList(bcolors.OKBLUE, txt1, list_txt,
                            list_id, electrical_test_circuit, spec_list, spec_keys, spec_id)
                        if oo == "следующий параметр":
                            break
                        elif oo=="отмена":
                            meter_typegroup_def=meter_typegroup_def_old
                            meter_type_def=meter_type_def_old
                            electrical_test_circuit=electrical_test_circuit_old
                            cicl1=False
                            break
                        else:
                            electrical_test_circuit = oo
                if cicl1==False:
                    break
                
                
                cicl2=True
                while cicl2:
                    os.system("CLS")
                    txt1 = "\nУкажите контрольное значение тока в схеме " \
                        "подключения ПУ для проведения проверки, А:"
                    res=readGonfigValue("var_all_value.json", [], {},
                        workmode, "1")
                    if res[0]=="1":
                        a_dic=res[2].get("ctrl_current_electr_test", {})
                        if len(a_dic)>0:
                            list_txt=list(a_dic["all_value"].keys())
                            list_id=list(a_dic["all_value"].values())
                            spec_list=["следующий параметр", "отмена"]
                            spec_id=spec_list
                            spec_keys=["\r","/"]
                            oo = questionFromList(bcolors.OKBLUE, txt1, list_txt,
                                list_id, ctrl_current_electr_test, spec_list, 
                                spec_keys, spec_id)
                            if oo == "следующий параметр":
                                break
                            elif oo=="отмена":
                                meter_typegroup_def=meter_typegroup_def_old
                                meter_type_def=meter_type_def_old
                                electrical_test_circuit=electrical_test_circuit_old
                                ctrl_current_electr_test=ctrl_current_electr_test_old
                                cicl1=False
                                break
                            else:
                                ctrl_current_electr_test = oo
                if cicl1==False:
                    break
                
                
                cicl2=True
                while cicl2:
                    os.system("CLS")
                    txt1 = "\nУкажите цвет корпуса ПУ:"
                    list_txt = ["спрашивать каждый раз", "белый", "черный", "серый"]
                    list_id = list_txt
                    spec_list=["следующий параметр", "отмена"]
                    spec_id=spec_list
                    spec_keys=["\r","/"]
                    oo = questionFromList(bcolors.OKBLUE, txt1, list_txt,
                                        list_id, meter_color_body, spec_list, spec_keys, spec_id)
                    if oo == "следующий параметр":
                        break
                    elif oo=="отмена":
                        meter_typegroup_def=meter_typegroup_def_old
                        meter_type_def=meter_type_def_old
                        electrical_test_circuit=electrical_test_circuit_old
                        ctrl_current_electr_test=ctrl_current_electr_test_old
                        meter_color_body=meter_color_body_old
                        cicl1=False
                        break
                    else:
                        meter_color_body = oo
                if cicl1==False:
                    break

                cicl2=True
                while cicl2:
                    os.system("CLS")
                    txt1 = "\nУкажите порядок действия при отклонении внутренних часов ПУ " \
                            "от эталона на 1 минуту или\n расчетной погрешности " \
                            "более 1 сек/сут для мск час. пояса:"
                    list_txt = ["спрашивать каждый раз", "корректировать автоматически для всех часовых поясов",
                                "корректировать автоматически только для московского часового пояса"]
                    list_id = ["0", "1", "2"]
                    spec_list = ["ok","отмена"]
                    spec_id=spec_list
                    spec_keys=["\r","/"]
                    oo = questionFromList(bcolors.OKBLUE, txt1, list_txt,
                                        list_id, meter_adjusting_clock, spec_list, spec_keys, spec_id)
                    if oo=="отмена":
                        meter_typegroup_def = meter_typegroup_def_old
                        meter_type_def = meter_type_def_old
                        electrical_test_circuit=electrical_test_circuit_old
                        ctrl_current_electr_test=ctrl_current_electr_test_old
                        meter_color_body = meter_color_body_old
                        meter_adjusting_clock = meter_adjusting_clock_old
                        cicl1=False
                        break
                    elif oo=="ok":
                        default_value_dict = writeDefaultValue(
                            default_value_dict)
                        saveConfigValue('opto_run.json', default_value_dict)
                        cicl1=False
                        break
                    else:
                        meter_adjusting_clock = oo
                
                if cicl1==False:
                    break

        elif oo=="о МС":
            cicl1=True
            while cicl1:
                cicl2=True
                while cicl2:
                    os.system("CLS")
                    txt1="\nВведите информацию о модуле связи:"
                    list_txt=["устанавливаться не будет","будет рабочим", "будет тестовым"]
                    list_id = ["0", "1", "2"]
                    spec_list = ["следующий параметр", "отмена"]
                    spec_keys=["\r","/"]
                    if modem_status=="0" or modem_status=="2":
                        spec_list = ["ok","отмена"]
                        spec_keys=["\r","/"]
                    spec_id=spec_list
                    oo = questionFromList(bcolors.OKBLUE, txt1, list_txt, list_id,
                        modem_status, spec_list, spec_keys, spec_id)
                    if oo == "следующий параметр":
                        break
                    elif oo == "отмена":
                        modem_status= modem_status_old
                        cicl1 = False
                        break
                    elif oo=="ok":
                        modem_type_def = "спрашивать каждый раз"
                        SIMcard_status="0"
                        default_value_dict = writeDefaultValue(default_value_dict)
                        saveConfigValue('opto_run.json',default_value_dict)
                        cicl1 = False
                        break
                    else:
                        modem_status=oo
                if cicl1==False:
                    break
                if modem_status_old != modem_status:
                    modem_type_def = "спрашивать каждый раз"
                    SIMcard_status = "0"

                if modem_status=="1":
                    sheet_name = "Product1"
                    if "i-prom.3-1" in meter_type_def:
                        sheet_name = "Product0"

                    while True:
                        os.system("CLS")
                        txt1 = f"\nДля ПУ '{meter_type_def}' укажите вид МС:"
                        list_txt = ["спрашивать каждый раз", "MC.1", "MC.3"] 
                        list_id = list_txt
                        spec_list=["определить по серийному номеру МС",
                            "следующий параметр", "отмена"]
                        spec_keys=["*","\r","/"]
                        if modem_typegroup_def=="спрашивать каждый раз":
                            spec_list=["определить по серийному номеру МС", "ok","отмена"]
                            spec_keys=["*", "\r", "/"]
                        spec_id=spec_list 
                        oo = questionFromList(bcolors.OKBLUE, txt1, list_txt,
                            list_id, modem_typegroup_def,spec_list,spec_keys,spec_id)
                        if oo == "следующий параметр":
                            break
                            
                        elif oo=="отмена":
                            modem_status = modem_status_old
                            cicl1=False
                            break

                        elif oo=="определить по серийному номеру МС":
                            txt1 = "\nВведите или отсканируйте серийный номер МС, указанный " \
                                    "на корпусе МС:" \
                                    "\nЧтобы прервать ввод - введите '/' и нажмите Enter."
                            oo=inputNumberDevice([15, 13],"модуль связи", sheet_name, txt1,["/"])
                            if oo[0]=="1":
                                num=oo[1]

                            else:
                                continue

                            res = toGetProductInfo2(num, sheet_name)
                            if res[0] == "0":
                                continue
                            modem_typegroup_def=res[4]
                            modem_type_def = res[2]

                        elif oo=="ok":
                            if modem_typegroup_def=="спрашивать каждый раз":
                                modem_type_def = "спрашивать каждый раз"
    
                            break

                        else:
                            modem_typegroup_def=oo
                                        
                    if modem_typegroup_def!="спрашивать каждый раз":
                        modem_filter_dict1 = {"MC.1":["MC.1"], "MC.3":["MC.3"]}
                        modem_filter=modem_filter_dict1[modem_typegroup_def]

                        modem_list = toFillListProductModel(modem_filter, sheet_name)

                        if not modem_type_def in modem_list:
                            modem_type_def="спрашивать каждый раз"
                        
                        cicl2=True
                        while cicl2:
                            os.system("CLS")
                            txt1 = f"\nДля ПУ '{meter_type_def}' укажите модель МС:"

                            if modem_list!="":
                                list_txt=["спрашивать каждый раз"]+modem_list
                                list_id = list_txt
                                spec_list=["определить сейчас по серийному номеру МС", 
                                    "следующий параметр", "отмена"]
                                spec_keys=["*","\r","/"]
                                spec_id=spec_list
                                oo = questionFromList(bcolors.OKBLUE, txt1, list_txt,
                                    list_id, modem_type_def, spec_list,spec_keys,spec_id)
                                
                                if oo=="следующий параметр":
                                    break

                                elif oo=="отмена":
                                    modem_status = modem_status_old
                                    modem_type_def=modem_type_def_old
                                    cicl1=False
                                    break

                                elif oo=="определить сейчас по серийному номеру МС":
                                    txt1 = "\nВведите или отсканируйте номер МС, указанный на его корпусе:" \
                                        "\nЧтобы прервать ввод - введите '/' и нажмите Enter."
                                    oo = inputNumberDevice([15,13], "модуль связи", sheet_name, txt1, ["/"])
                                    if oo[0] == "1":
                                        num = oo[1]
                                    
                                    else:
                                        continue

                                    res=toGetProductInfo2(num, sheet_name)
                                    if res[0]=="0":
                                        continue

                                    modem_product_type = res[2]

                                    if not modem_product_type in modem_list:
                                        txt1 = "По введенному номеру определена модель МС '"+modem_product_type+ \
                                            "', но ее нет в списке.\nУстановить эту модель как значение " \
                                            "по умолчанию? (0-нет, 1-да)"
                                        oo = questionSpecifiedKey(bcolors.WARNING,txt1,["0", "1"])
                                        if oo=="0":
                                            continue

                                        modem_list.append(modem_product_type)

                                    modem_type_def=modem_product_type
                                    
                                else:
                                    modem_type_def = oo
                            
                            else:
                                printWARNING("Не удалось сформировать список моделей МС")
                                keystrokeEnter()
                                cicl1=False
                                break
                            
                    if cicl1==False:
                        break

                if modem_status == "1":
                    cicl2=True
                    while cicl2:
                        os.system("CLS")
                        list_txt=["устанавливаться не будет", "будет рабочей, запрашивать номер карты", 
                                "будет рабочей, номер карты не нужен", "будет тестовой"]
                        list_id=["0","1","3","2"]
                        if modem_status=="2":
                            list_txt=["устанавливаться не будет", "будет тестовой"]
                            list_id=["0","2"]
                        spec_list=["ok","отмена"]
                        spec_id=spec_list
                        spec_keys=["\r","/"]
                        txt1="\nВведите информацию о статусе SIM-карты:"
                        oo = questionFromList(bcolors.OKBLUE, txt1, list_txt, list_id,
                            SIMcard_status,spec_list, spec_keys, spec_id) 
                        if oo == "ok":
                            default_value_dict = writeDefaultValue(default_value_dict)
                            saveConfigValue('opto_run.json',
                                            default_value_dict)
                            cicl1=False
                            break
                        elif oo=="отмена":
                            modem_status = modem_status_old
                            modem_type_def= modem_type_def_old
                            SIMcard_status = SIMcard_status_old
                            cicl1=False
                            break
                        else:
                            SIMcard_status=oo
                    if cicl1==False:
                        break

        elif oo=="взаимодействие с СУТП":
            var_name_list = ["sutp_to_save"]
            
            header = "Настройка взаимодействия с СУТП."

            res = menuChangeValue(var_name_list, employee_id, 
                "var_all_value.json", header, bcolors.OKBLUE, 
                workmode)
            if res[0] == "1":
                default_value_dict.update(res[2])
                readDefaultValue(default_value_dict)


        elif oo=="контроль заказа":
            header = "Принадлежность ПУ заказу."
            if order_control=="1" and data_exchange_sutp=="0":
                header=header+f"\n{bcolors.WARNING}Для выполнения " \
                    "контроля необходимо включить обмен данными " \
                    f"с СУТП.{bcolors.ENDC}"

            var_name_list = ["order_control"]
            
            res = menuChangeValue(var_name_list, employee_id, 
                "var_all_value.json", header, bcolors.OKBLUE, 
                workmode)
            if res[0] == "1":
                default_value_dict.update(res[2])
                readDefaultValue(default_value_dict)
            

        elif oo=="конфигурация ПУ":
            header = "Настройка параметров для проверки конфигурации ПУ."

            var_name_list = ["meter_config_check",
                "com_config_current_select", "config_send_mail",
                "com_config_user"]

            res = menuChangeValue(var_name_list, employee_id, 
                "var_all_value.json", header, bcolors.OKBLUE, 
                workmode)
            if res[0] == "1":
                default_value_dict.update(res[2])
                readDefaultValue(default_value_dict)


        elif oo == "прочие параметры":
            header="Настройка прочих параметров."

            var_name_list=["rep_copy_public", "actions_no_mc", 
                "res_ext_at_begin_test", "rep_err_send_mail", 
                "no_data_in_SUTP_send_mail"]
            
            res = menuChangeValue(var_name_list, employee_id, 
                "var_all_value.json", header, bcolors.OKBLUE, 
                workmode)
            if res[0]=="1":
                default_value_dict.update(res[2])
                readDefaultValue(default_value_dict)
                
        
        elif oo=="взаимодействие с MassConfig":
            header="Настройка взаимодействия с программой " \
                "'MassProdAutoConfig.exe' для проверки конфигурации ПУ."

            var_name_list=["mass_control_com_port", "time_wait_exec",
                "time_wait_open_window", "time_interval",
                "time_interval_limit", "mass_log_split_print",
                "mass_log_print_analysis"]
            
            res = menuChangeValue(var_name_list, employee_id, 
                "var_all_value.json", header, bcolors.OKBLUE, 
                workmode)
            

        elif oo=="7":
            cicl2=True
            while cicl2:
                os.system("CLS")
                txt1="\nУкажите необходимое состояние диктора " \
                    "(звуковое сопровождение проверки ПУ):"
                list_txt = ["отключен", "включен"]
                list_id = ["0", "1"]
                spec_list = ["ok","отмена"]
                spec_keys=["\r","/"]
                spec_id = spec_list
                oo= questionFromList(bcolors.WARNING, txt1, list_txt, list_id, speaker,
                    spec_list, spec_keys, spec_id)
                if oo == "ok":
                    default_value_dict = writeDefaultValue(default_value_dict)
                    saveConfigValue('opto_run.json',default_value_dict)
                    break
                elif oo == "отмена":
                    speaker = speaker_old
                    break
                else:
                    speaker = oo


def toPrintDefaultValue():
    global var_all_value_dic    #словарь со всеми предлагаемыми вариантами значений для

    global number_of_meters     #количество подключаемых ПУ
    global order_control
    global order_control_descript
    global data_exchange_sutp
    global com_config_current_select
    global com_config_user
    global ctrl_current_electr_test #контрольное значение тока в схеме подключения
                                
    
    txt1 = "Параметры программ проверки ПУ.\n" \
        f"количество одновременно подключаемых ПУ: "\
        f"{bcolors.OKGREEN}{str(number_of_meters)}{bcolors.ENDC}\n" \
        f"пользователь: {bcolors.OKGREEN}{employees_name}{bcolors.ENDC}\n" \
        "описание пароля по умолчанию для подключения: " \
        f"{bcolors.OKGREEN}{meter_pw_default_descript}{bcolors.ENDC}\n" \
        f"модель ПУ: {bcolors.OKGREEN}{meter_type_def}{bcolors.ENDC}"
    printColor(txt1)
    
    print_var_list = ["electrical_test_circuit", 
        "meter_color_body", "meter_adjusting_clock", "modem_status", "SIMcard_status",
        "data_exchange_sutp", "sutp_to_save", "meter_config_check",
        "order_control", "com_config_current_select", "config_send_mail",
        "rep_err_send_mail","no_data_in_SUTP_send_mail","res_ext_at_begin_test"]

    for var_name in print_var_list:
        a_descript = var_all_value_dic[var_name]["descript"]
        
        a_all_value_dic = var_all_value_dic[var_name]["all_value"]

        a_value_standart_list=var_all_value_dic[var_name]["value_standart"]

        txt = f"{a_descript}"
        a_keys_list=list(a_all_value_dic.keys())
        for a_key in a_keys_list:
            if a_all_value_dic[a_key]==globals()[var_name]:
                if len(a_value_standart_list)>0 and \
                    not globals()[var_name] in a_value_standart_list:
                    txt = f"{txt}: {bcolors.ATTENTIONWARNING} {a_key} "

                else:
                    txt = f"{txt}: {bcolors.OKGREEN}{a_key}"
            
        if var_name=="order_control":
            if order_control=="1":
                if order_control_descript!="":
                    txt=f"{a_descript}: {bcolors.OKGREEN}включен для заказа " \
                        f"'{order_control_descript[0:31]}...'"
                if data_exchange_sutp=="0":
                    txt=f"{txt}, {bcolors.WARNING}но требуется включить " \
                        "обмен данными с СУТП"

        elif var_name=="com_config_current_select" and \
            com_config_current_select=="0":
            a_dic={"com_config_opto": f"оптопорт, подключенный к {com_config_opto}",
                "com_config_rs485": f"RS-485, подключенный к {com_config_rs485}"}
            a_val_dic = {"0": a_dic[com_config_user], "1": "автоматически"}
            txt=f"{a_descript}: {bcolors.ATTENTIONWARNING} " \
                f"{a_val_dic[com_config_current_select]}"
            
        elif var_name=="electrical_test_circuit":
            a_dic={"1-0": f"напряжение - 1 фаза, ток - нет",
                "1-1": f"напряжение - 1 фаза, ток - 1 фаза ({ctrl_current_electr_test} А)",
                "2-0": f"напряжение - 2 фазы, ток - нет",
                "2-1": f"напряжение - 2 фазы, ток - 1 фаза ({ctrl_current_electr_test} А)",
                "2-2": f"напряжение - 2 фазы, ток - 2 фазы ({ctrl_current_electr_test} А)",
                "3-0": f"напряжение - 3 фазы, ток - нет",
                "3-1": f"напряжение - 3 фазы, ток - 1 фаза ({ctrl_current_electr_test} А)",
                "3-2": f"напряжение - 3 фазы, ток - 2 фазы ({ctrl_current_electr_test} А)",
                "3-3": f"напряжение - 3 фазы, ток - 3 фазы ({ctrl_current_electr_test} А)"
                }
            txt=f"{a_descript}: {bcolors.OKGREEN}{a_dic[electrical_test_circuit]}"
        
        printColor(txt)

        if var_name=="modem_status" and modem_status=="1":
            printColor(f"модель модуля связи: {bcolors.OKGREEN}{modem_type_def}")

    print()   



def checkPW(pass_filename="meter_pass.json", order_pw_type="заводской пароль"):

    global meter_pw_encrypt     #текущий зашифрованный пароль подключения к ПУ
    global meter_pw_descript    #описание текущего пароля
    global default_value_dict
    global meter_pw_level       # уровень доступа к ПУ: "High", "Low"
    global com_opto             # номер COM-порта для оптопорта
    global com_rs485             # номер COM-порта для RS-485

    
    meter_pw_encrypt_old=meter_pw_encrypt
    meter_pw_descript_old = meter_pw_descript
    res = readGonfigValue(pass_filename)
    if res[0] != "1":
            return ["4", "Не удалось прочитать данные "
                    "о пароле из файла", "", ""]
    
    a_stand_pw_dic={"Low": "654321", "High": "1234567898765432"}
    a_stand_pw=a_stand_pw_dic[meter_pw_level]

    meter_pw_default_dict = res[2]
    meter_pw_dic={}
    if meter_pw_encrypt!="":
        if meter_pw_descript in meter_pw_default_dict:
            meter_pw_default_dict.pop(meter_pw_descript,0)
        _,_, meter_pw=cryptStringSec("расшифровать",meter_pw_encrypt)
        meter_pw_dic[meter_pw_descript]=meter_pw

        if meter_pw_level=="High" and ("СГЕНЕРИРОВАННЫЙ" in order_pw_type.upper()):
            if meter_pw==a_stand_pw:
                a_err="Сгенерированный пароль верхнего уровня совпадает " \
                    "со стандартным значением"
                return ["7", a_err,"",""]
        
    meter_pw_dic.update(meter_pw_default_dict)
    keys_list = list(meter_pw_dic.keys())
    _, ans2, file_name_check = getUserFilePath('otk_check_opto.py',
        workmode=workmode)
    if file_name_check == "":
        return ["4", f"Ошибка в ПП getUserFilePath(): {ans2}","",""]
    try_num=0
    key_quantity=len(keys_list)
    for key in keys_list:
        meter_pw = meter_pw_dic[key]
        meter_pw_descript = key
        print (f"Проверяем пароль '{meter_pw_descript}'.")
        res=cryptStringSec("зашифровать",meter_pw)
        meter_pw_encrypt=res[2]
        default_value_dict = writeDefaultValue(
            default_value_dict)
        saveConfigValue('opto_run.json', default_value_dict)
        code_exit = os.system(f"python {file_name_check} 2 0")
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
                    operation == "checkPassword": 
                if result!="4":
                    res=readGonfigValue(file_name_in="opto_run.json",
                        var_name_list=[],default_value_dict=default_value_dict)
                    if res[0]!="1":
                        txt1=f"{bcolors.WARNING}При проверке пароля доступа к ПУ " \
                            f"возникла ошибка.{bcolors.ENDC}\n" \
                            f"{bcolors.OKBLUE}Нажмите Enter{bcolors.ENDC}"
                        oo=questionSpecifiedKey("",txt1,["\r"],"",1)
                        return ["0", "При проверке пароля доступа к ПУ возникла ошибка.",
                                "", ""]
                    a_dic=res[2]
                    com_opto=a_dic.get("com_opto","")
                    com_rs485=a_dic.get("com_rs485","")
                    
                if result=="1":
                    if meter_pw_descript!=meter_pw_descript_old:
                        txt1 = f"{bcolors.OKGREEN}Найден пароль для подключения к ПУ - " \
                            f"'{meter_pw_descript}'.{bcolors.ENDC}\n" \
                            f"{bcolors.OKBLUE}Использовать его? 0-нет, 1-да.{bcolors.ENDC}"
                        key1 = ["0","1"]
                        oo = questionSpecifiedKey("", txt1, key1)
                        print()
                        if oo=="0":
                            meter_pw_encrypt_new=meter_pw_encrypt
                            meter_pw_descript_new=meter_pw_descript
                            meter_pw_encrypt=meter_pw_encrypt_old
                            meter_pw_descript = meter_pw_descript_old
                            default_value_dict = writeDefaultValue(
                                default_value_dict)
                            saveConfigValue('opto_run.json',
                                            default_value_dict)
                            return ["3","Пользователь отказался изменить пароль",
                                    meter_pw_encrypt_new, meter_pw_descript_new]
                        else:
                            print(f"\n{bcolors.OKGREEN}Установлен пароль '{meter_pw_descript}' "
                            f"для доступа к ПУ.{bcolors.ENDC}")
                            return ["5","Подобран новый пароль",meter_pw_encrypt, \
                                    meter_pw_descript]
                    else:
                        print(f"{bcolors.OKGREEN}Доступ к ПУ по паролю '{meter_pw_descript}' "
                            f"предоставлен.{bcolors.ENDC}")
                        return ["1","Текущей пароль верен.", meter_pw_encrypt, \
                                meter_pw_descript]
                elif result=="2":
                    if try_num==0 and key_quantity>1:
                        print(f"{bcolors.WARNING}Пароль {meter_pw_descript} не подошел." 
                            f"{bcolors.ENDC}\n{bcolors.WARNING}Попробуем следующий пароль "\
                            f"из списка по умолчанию.{bcolors.ENDC}")
                    try_num+=1
                
                elif result=="4":
                    return ["6","COM-порт не найден.", "", ""]
                
                else:
                    return ["0","Прочие ошибки.", "", ""]

    txt1 = f"{bcolors.FAIL}Подходящий пароль не найден.{bcolors.ENDC}\n" \
        f"{bcolors.OKBLUE}Нажмите Enter.{bcolors.ENDC}"
    key1 = ["\r"]
    oo = questionSpecifiedKey("", txt1, key1,"",1)
    meter_pw_encrypt=meter_pw_encrypt_old
    meter_pw_descript = meter_pw_descript_old
    default_value_dict = writeDefaultValue(default_value_dict)
    saveConfigValue('opto_run.json', default_value_dict)
    return ["2","Подходящего пароля не нашли в списке паролей по умолчанию.","",""]



def restoreCOMPort(com_name:str):

    global default_value_dict   #словарь значений по умолчанию
    global com_opto             #COM-порт, к которому подключен оптопорт
    global com_rs485            #COM-порт, к которому подключен RS-485


    a_dic={"com_opto":["оптопорт","оптопорта"],
        "com_rs485":["преобразователь RS-485","преобразователя RS-485"],
        "com_config_opto":["оптопорт","оптопорта"],
        "com_config_rs485":["преобразователь RS-485",
            "преобразователя RS-485"]}
    print()
    comment_txt=f"{bcolors.OKGREEN}Попробуем восстановить подключение " \
        f"к COM-порту {a_dic.get(com_name,'')[1]}.{bcolors.ENDC}"
    res=getAutoCOMPort(a_dic.get(com_name,'')[0], "1", comment_txt)
    if res[0]=="1":
        com_val=res[2]
        globals()[com_name]=com_val
        default_value_dict = writeDefaultValue(default_value_dict)
        saveConfigValue('opto_run.json', default_value_dict)
        return "1"
    return res[0]



def connectionMeterSetup():
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


    def innerGetComPort(interface_name, window_title_list, position):

        printBLUE(f"Определение COM-порта для позиции №{position+1}.")
        printBLUE(f"Подключите {interface_name} к компьютеру.")

        res=getAutoCOMPort(interface_name,"2","", window_title_list)

        return res

    
    def innerMultiGetComPort(mode):

        global default_value_dict
        global com_opto             #COM-порт, к которому подключен оптопорт

        a_com_dic = {"мульти com-порт": ["оптопорт", "multi_com_opto_dic",
            "Настройка оптопорта для аппаратной проверки ПУ.", "com_opto"],
            "мульти com-порт RS-485": ["преобразователь RS-485", 
                                "multi_com_rs485_dic",
            "Настройка RS-485 для аппаратной проверки ПУ.", "com_rs485"],
            "мульти com-config-opto": ["оптопорт", 
                                        "multi_com_config_opto_dic",
            "Настройка оптопорта для проверки конфигурации ПУ.", 
            "com_config_opto"],
            "мульти com-config-RS485": ["преобразователь RS-485", 
                                    "multi_com_config_rs485_dic",
            "Настройка RS-485 для проверки конфигурации ПУ.",
            "com_config_rs485"]}
        
        var_interface_name=a_com_dic[mode][1]
        var_multi_dic = globals()[var_interface_name]
        header_txt=a_com_dic[mode][2]
        interface_name=a_com_dic[mode][0]
        com_var_name=a_com_dic[mode][3]
    
        header=f"{bcolors.WARNING}{header_txt}\n" \
            "Проверьте, чтобы интерфейс, для которого определяем COM-порт " \
            f"был отключен от компьютера.{bcolors.ENDC}"
        header=header+"\nВыберите позицию на стенде, для " \
            "которой необходимо определить COM-порт."

        com_list=var_multi_dic["com_name"].copy()
        captipon_list=var_multi_dic["caption"].copy()
        
        if len(com_list)!=number_of_meters:
            a_txt="Для корректного определения COM-портов отключите " \
                "все подключаемые к ПУ "
            if interface_name=="оптопорт":
                a_txt=a_txt+"оптопорты"
            
            else:
                a_txt=a_txt+"интерфейсы RS-485"
            
            a_txt=a_txt+" от компьютера.\nПри готовности - " \
                "нажмите Enter.\nДля возврата в меню - " \
                "нажмите - '/'."
            spec_keys=["\r", "/"]
            oo=inputSpecifiedKey(bcolors.OKBLUE, a_txt, "", [], 
                spec_keys, 1, "")
            if oo=="/":
                return
            
            com_list=[""]*number_of_meters
            captipon_list=[""]*number_of_meters
            
        
        alternately_on=False

        while True:
            os.system("CLS")

            menu_item_list=[]
            menu_id_list=[]

            position_comport_no_list=[]
            
            for i in range(0, number_of_meters):
                com_cur=com_list[i]
                a_caption=f"позиция № {i+1}"
                if com_cur!=None and com_cur!="":
                    res=checkComPortList([com_cur], "no", "")
                    if res[0]=="1":
                        if captipon_list[i]!=None and captipon_list[i]!="":
                            a_caption=captipon_list[i]
                        menu_item_list.append(f"{a_caption}: {com_cur}")
                        menu_id_list.append(i)

                    else:
                        position_comport_no_list.append(i)
                        menu_item_list.append(f"{a_caption}:")
                        menu_id_list.append(i)
                        com_list[i]=""

                else:
                    position_comport_no_list.append(i)
                    menu_item_list.append(f"{a_caption}:")
                    menu_id_list.append(i)

            menu_spec_item_list=["ok", "отмена"]
            menu_spec_id_list=menu_spec_item_list.copy()
            menu_spec_keys_list=["\r", "/"]
            if len(position_comport_no_list)>0:
                a_txt="поочередное определение для позиций с отсутствующими портами"
                menu_spec_item_list.append(a_txt)
                menu_spec_id_list.append("поочередное")
                menu_spec_keys_list.append("*")
                
                if "RS-485" in interface_name:
                    a_txt="определение одного общего порта для всех позиций с " \
                        "отсутствующими портами"
                    menu_spec_item_list.append(a_txt)
                    menu_spec_id_list.append("один общий")
                    menu_spec_keys_list.append("+")

            else:
                alternately_on=False

            if alternately_on:
                oo="поочередное"
                printWARNING(header_txt)
                for a_item in menu_item_list:
                    printBLUE(a_item)
                    
            else:
                oo=questionFromList(bcolors.OKBLUE, header, menu_item_list, menu_id_list,
                    "", menu_spec_item_list, menu_spec_keys_list, menu_spec_id_list, 1, 1, 1,
                    [], "")
                print()
            
            if oo=="отмена":
                return

            elif oo=="ok":
                var_multi_dic["caption"]=captipon_list.copy()
                var_multi_dic["com_name"]=com_list.copy()

                if com_list[0]!=None and com_list[0]!="":
                    globals()[com_var_name]=com_list[0]

                globals()[var_interface_name]= var_multi_dic.copy()
                default_value_dict = writeDefaultValue(default_value_dict)
                saveConfigValue('opto_run.json', default_value_dict)
                return

            elif oo=="поочередное":
                alternately_on=True
                position=position_comport_no_list[0]

                res=innerGetComPort(interface_name, [window_title], position)
                if res[0]=="1":
                    com_list[position]=res[2]
                    if captipon_list[position]==None or \
                        captipon_list[position]=="":
                        captipon_list[position]="позиция № " \
                            f"{position+1}"

                elif res[0]=="3":
                    alternately_on=False

            elif oo=="один общий":
                printBLUE("Определение одного общего COM-порта RS-485 "
                    "для свободных позиций.")
                printBLUE(f"Подключите {interface_name} к компьютеру.")

                res=getAutoCOMPort(interface_name,"2","", [window_title])
                if res[0]=="1":
                    for position in position_comport_no_list:
                        com_list[position]=res[2]
                        if captipon_list[position]==None or \
                            captipon_list[position]=="":
                            captipon_list[position]="позиция № " \
                                f"{position+1}"
                
                else:
                    continue

            
            else:
                position=oo
    
                res=innerGetComPort(interface_name, [window_title], position)
                if res[0]=="1":
                    com_list[position]=res[2]
                    if captipon_list[position]==None or \
                        captipon_list[position]=="":
                        captipon_list[position]="позиция № " \
                            f"{position+1}"
                
                else:
                    continue
   
    
    default_value_dict = optoRunVarRead()
    readDefaultValue(default_value_dict)

    
    window_title = GetWindowText(GetForegroundWindow())

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

        comports_list=[multi_com_opto_dic["com_name"], 
            multi_com_rs485_dic["com_name"],
            multi_com_config_opto_dic["com_name"],
            multi_com_config_rs485_dic["com_name"]]

        comports_txt_dic={}
        for i in range(0, len(comports_list)):
            com_cur_list=comports_list[i]
            comports_txt_dic[i]=f"{bcolors.FAIL}список используемых COM-портов пуст{bcolors.ENDC}"
            if number_of_meters==1:
                comports_txt_dic[i]=f"{bcolors.FAIL}не указан COM-порт{bcolors.ENDC}"

            if len(com_cur_list)>0:
                if len(com_cur_list)<number_of_meters:
                    comports_txt_dic[i]=f"{bcolors.FAIL}список используемых COM-портов не полон{bcolors.ENDC}"
                    if number_of_meters==1:
                        comports_txt_dic[i]=f"{bcolors.FAIL}не указан COM-порт{bcolors.ENDC}"

                else:
                    com_cur_list=com_cur_list[0:number_of_meters]
                    res=checkComPortList(com_cur_list, "no", "")
                    if res[0]=="1":
                        a_com_str=", ".join(com_cur_list)
                        comports_txt_dic[i]=f"{bcolors.OKGREEN}{a_com_str}"
                    
                    elif res[0]=="2":
                        a_str=", ".join(res[2])
                        comports_txt_dic[i]=f"{bcolors.FAIL}отсутствуют {a_str}"
                        if len(res[2])==1:
                            comports_txt_dic[i]=f"{bcolors.FAIL}отсутствует {a_str}"

                    elif res[0]=="3":
                        comports_txt_dic[i]=f"{bcolors.FAIL}список используемых " \
                            f"COM-портов не полон{bcolors.ENDC}"
                        if number_of_meters==1:
                            comports_txt_dic[i]=f"{bcolors.FAIL}не указан COM-порт{bcolors.ENDC}"
                    
                    elif res[0]=="4" and number_of_meters>1:
                        a_str=", ".join(com_cur_list)
                        comports_txt_dic[i]=f"{bcolors.WARNING}в списке используемых " \
                            f"COM-портов имеются повторения ({a_str}){bcolors.ENDC}"



        os.system("CLS")
        a_eqv_dic={"0": "отключено", "1": "включено"}
        txt1 = "Выберите пункт меню"
        list_txt = [f"Выбор пароля высокого уровня по умолчанию для подключения к ПУ: {txt_pass}",
            f"Совместное использование COM-портов для проверки ПУ: " \
            f"{bcolors.OKGREEN}{a_eqv_dic[com_config_eqv_com]}",
            f"Настройка оптопорта для аппаратной проверки ПУ: {comports_txt_dic[0]}",
            f"Настройка RS-485 для аппаратной проверки ПУ: {comports_txt_dic[1]}"]
        list_id = ["мульти пароль", "мульти использование", "мульти com-порт", 
            "мульти com-порт RS-485"]
        if com_config_eqv_com=="0":
            a_list=[f"Настройка оптопорта для проверки конфигурации ПУ: {comports_txt_dic[2]}",
                f"Настройка RS-485 для проверки конфигурации ПУ: {comports_txt_dic[3]}"]
            list_txt.extend(a_list)
            a_id_list = ["мульти com-config-opto", "мульти com-config-RS485"]
            list_id.extend(a_id_list)

        cur_id = ""
        spec_list=["Выход"]
        spec_keys=["/"]
        spec_id=["выход"]           
        oo = questionFromList(bcolors.OKBLUE, txt1, list_txt, list_id, cur_id, \
                    spec_list=spec_list, spec_keys=spec_keys, spec_id=spec_id)
        print()

        if oo == "выход":
            break

        elif oo == "мульти пароль":
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

        elif oo=="мульти использование":
            res = readGonfigValue("var_all_value.json", [], {}, workmode, "1")
            if res[0] == "1":
                all_value_dic=res[2]["com_config_eqv_com"]["all_value"]
                menu_item_list=list(all_value_dic.keys())
                menu_id_list=list(all_value_dic.values())

                header="Совместное использование COM-портов для проверки ПУ:"
                res=menuSimple(bcolors.OKBLUE, header, menu_item_list, menu_id_list,
                    com_config_eqv_com)
                if res[0]=="1":
                    com_config_eqv_com=res[2]
                    default_value_dict = writeDefaultValue(default_value_dict)
                    saveConfigValue('opto_run.json', default_value_dict)


        elif oo in ["мульти com-порт", "мульти com-порт RS-485", 
                    "мульти com-config-opto", "мульти com-config-RS485"]:
            
            innerMultiGetComPort(oo)



def changeNumOfMeters():
    global default_value_dict       # список со зн-ями по умолчанию
    global workmode                 #метка режима работы программы "тест" - режим теста1, 

    
    default_value_dict = optoRunVarRead()
    readDefaultValue(default_value_dict)

    workmode=default_value_dict["workmode"]

    header = "Изменение количества одновременно подключаемых ПУ."

    var_name_list = ["number_of_meters"]

    res = menuChangeValue(var_name_list, "", 
        "var_all_value.json", header, bcolors.OKBLUE, 
        workmode)
    if res[0] == "1":
        default_value_dict.update(res[2])
        readDefaultValue(default_value_dict)
        saveConfigValue('opto_run.json', default_value_dict)



def menuMain():
    global default_value_dict  # словарь со зн-ями по умолчанию
    global workmode             #метка режима работы программы "тест" - режим теста1, 
    global number_of_meters     #количество подключаемых ПУ на стенде
    global meter_position_cur   #номер текущей позиции ПУ на стенде
    global meter_tech_number_list #список технических номеров ПУ, проверяемых на стенде
    global meter_tech_number_start_list #стартовый список технических номеров ПУ,
    global meter_serial_number_start_list #стартовый список серийных номеров ПУ, 
    global meter_status_test_list   #список статусов ПУ при прохождении проверки:
    global employee_id          # таб. номер пользователя
    global employee_pw_encrypt  # зашифрованный пароль пользователя
    global sutp_to_save         #способ записи рез-тов теста в БД СУТП ("0"-отключен, "1"-ручной,
    global com_opto             #COM-порт, к которому подключен оптопорт
    global com_rs485            #COM-порт, к которому подключен RS-485
    global com_current          #вид активного порта для связи с ПУ:"com_opto","com_rs485"
    global com_config_opto      #COM-порт, к которому подключен оптопорт для проверки конфигурации ПУ
    global com_config_rs485     #COM-порт, к которому подключен RS-485 для проверки конфигурации ПУ
    global com_config_current   #вид активного порта для связи с ПУ при проверке конфигурации ПУ:
    global multi_com_opto_dic   #словарь со списками используемых COM-портов
    global multi_com_rs485_dic   #словарь со списками используемых COM-портов
    global multi_com_config_opto_dic   #словарь со списками используемых COM-портов
    global multi_com_config_rs485_dic   #словарь со списками используемых COM-портов
    global com_config_eqv_com   #метка использования для проверки конфигурации ПУ
    global pw_decrypt_visible   #вывод на экран паролей доступа к ПУ: "0"-откл, "1"-вкл.
    global meter_pw_visible     #вывод на экран паролей доступа к ПУ: "0"-откл, "1"-вкл.
    global meter_config_check   #метод проверки конфигурации ПУ: "0"-откл.,
    global print_number_big_font    # метка печать на экране вводимых номеров изделий
    global mass_prod_vers       # версия программы MassProd, которая должна использоваться 
    
    
    
    def innerCheckComPortInterface():

        interface_list=[["оптопортов", multi_com_opto_dic["com_name"], 
            multi_com_config_opto_dic["com_name"], bcolors.FAIL],
            ["RS-485", multi_com_rs485_dic["com_name"], 
            multi_com_config_rs485_dic["com_name"], bcolors.WARNING]]
        
        type_of_check_txt_list=["аппаратной проверки", 
            "проверки конфигурации ПУ"]

        a_com_opto_no=False
        a_com_opto_dubl=False

        a_com_rs485_no=False
        a_com_rs485_dubl=False

        for i_type_of_check in range(0,2):
            com_no_txt=f"Для {type_of_check_txt_list[i_type_of_check]} " \
                "не настроено соединение"
            com_double_txt=f"Для {type_of_check_txt_list[i_type_of_check]} " \
                "имеются повторяющиеся COM-порты для"

            if com_config_eqv_com=="1":
                com_no_txt=f"Не настроено соединение"
                com_double_txt=f"Имеются повторяющиеся COM-порты для"
                

            for ind in range(0,2):
                comport_txt=interface_list[ind][0]
                
                comports_list=[interface_list[ind][1]]
                
                if i_type_of_check==1:
                    comports_list=[interface_list[ind][2]]
                
                com_no_color=interface_list[ind][3]

                a_com_no=False
                a_com_dubl=False
                
                for i in range(0, len(comports_list)):
                    com_cur_list=comports_list[i]
                    if len(com_cur_list)>0 and len(com_cur_list)>=number_of_meters:
                        com_cur_list=com_cur_list[0:number_of_meters]
                        res=checkComPortList(com_cur_list, "no", "")
                        if res[0] in ["2", "3"]:
                            a_com_no=True

                        elif res[0]=="4" and number_of_meters>1:
                            a_com_dubl=True 

                    else:
                        a_com_no=True
                
                if a_com_no:
                    txt1=f"{com_no_color}{com_no_txt} {comport_txt} с ПУ.{bcolors.ENDC}"
                    printColor(txt1)

                    if ind==0:
                        a_com_opto_no=True

                    else:
                        a_com_rs485_no=True

                if a_com_dubl:
                    txt1=f"{com_double_txt} {comport_txt}."
                    printWARNING(txt1)

                    if ind==1:
                        a_com_opto_dubl=True

                    else:
                        a_com_rs485_dubl=True
            
            if com_config_eqv_com=="1":
                break

        return [a_com_opto_no, a_com_rs485_no, a_com_opto_dubl, 
            a_com_rs485_dubl]
    


    def innerQuestionBreakTest():

        global number_of_meters
        global meter_position_cur

        a_txt="Проверка прервана."
        ret_id="9"
        ret_txt="Проверка всех ПУ прервана."
        if number_of_meters>1 and meter_position_cur!=-1:
            txt1=f"Выберите дальнейшее действие для позиции {meter_position_cur+1}:"
            menu_item_list=["Будет произведена замена ПУ на позиции",
                "Перейти к следующей позиции",
                "Прервать проверку всех ПУ на стенде"]
            menu_id_list=["замена ПУ", "исключить ПУ", "прервать"]
            oo=questionFromList(bcolors.OKBLUE, txt1, menu_item_list, menu_id_list,
                "", [], [], [], 1, 1, 1)
            print()
            if oo=="замена ПУ":
                a_txt=f"Проверка текущего ПУ на позиции {meter_position_cur+1} прервана.\n" \
                    "Произведите замену ПУ на новый прибор."
                ret_id="91"
                ret_txt="Проверка ПУ прервана. Требуется замена ПУ."

            elif oo=="исключить ПУ":
                a_txt=f"Проверка текущего ПУ на позиции {meter_position_cur+1} прервана.\n" \
                    "Он будет исключен из дальнейшей проверки."
                ret_id="92"
                ret_txt="Исключить позицию ПУ на стенде из проверки."

        printWARNING (a_txt)

        return [ret_id, ret_txt]
    
    
    def innerMeterTest(meter_test_mode="полная проверка"):

        global default_value_dict
        global meter_config_check
        global sutp_to_save         #способ записи рез-тов теста в БД СУТП ("0"-отключен, "1"-ручной,
        global meter_tech_number
        global meter_tech_number_list #список технических номеров ПУ, проверяемых на стенде
        global meter_tech_number_start_list #стартовый список технических номеров ПУ, 
        global meter_serial_number_start_list #стартовый список серийных номеров ПУ, 
        global meter_status_test_list   #список статусов ПУ при прохождении проверки:

        global meter_serial_number
        global meter_serial_number_list #список серийных номеров ПУ, проверяемых на стенде
        global meter_soft       #версия ПО текущего ПУ
        global meter_soft_list  #список версий ПО проверяемых на стенде ПУ
        global gsm_serial_number    #номер GSM модема (эталон). Порядок выбора:крышка
        global gsm_serial_number_list #список с номерами модулей связи ПУ, установленных 
        global rc_serial_number     #серийный номер ПДУ
        global rc_serial_number_list  #список с номерами ПДУ ПУ, установленных на стенде

        global meter_position_cur
        global number_of_meters
        global workmode

        global multi_com_opto_dic   #словарь со списками используемых COM-портов
        global multi_com_rs485_dic   #словарь со списками используемых COM-портов
        global com_opto             #COM-порт, к которому подключен оптопорт
        global com_rs485            #COM-порт, к которому подключен RS-485


        os.system("CLS")

        if meter_config_check=="3":
            a_txt=f"{bcolors.WARNING}Для автоматической проверки " \
                "конфигурации ПУ необходимо,\nчтобы была включена " \
                f"{bcolors.ATTENTIONWARNING} АНГЛИЙСКАЯ (EN) "\
                f"{bcolors.ENDC} {bcolors.WARNING}раскладка " \
                f"клавиатуры.\n"
            printColor(a_txt)

        res=getSutpToSaveDescript(sutp_to_save)
        sutp_to_save_descript=res[2]
        sutp_to_save_color=res[3]
        if  sutp_to_save[0]=="0" or sutp_to_save[0]=="1":
            print (f"{sutp_to_save_color}{sutp_to_save_descript}{bcolors.ENDC}\n")
        
        meter_position_cur=0
        meter_tech_number_list=[]

        meter_tech_number_start_list=[]

        meter_serial_number_start_list=[]

        meter_status_test_list=[]

        meter_serial_number_list=[]

        gsm_serial_number_list=[]

        rc_serial_number_list=[]

        meter_soft_list=[]

        default_value_dict = writeDefaultValue(default_value_dict)
        saveConfigValue('opto_run.json',default_value_dict)

        a_txt="Проверка счетчика."

        if meter_test_mode=="конфигурация":
             a_txt="Проверка только конфигурации ПУ."

        a_mode="номер ПУ"
            
        title_new="Аппаратная проверка ПУ"
        res = replaceTitleWindow("", title_new)
        
        while meter_position_cur < number_of_meters:
            com_opto=""
            com_rs485=""
            a_com_opto_list=multi_com_opto_dic["com_name"]
            a_com_rs485_list=multi_com_rs485_dic["com_name"]

            if len(a_com_opto_list)>meter_position_cur:
                com_opto=a_com_opto_list[meter_position_cur]

            if len(a_com_rs485_list)>meter_position_cur:
                com_rs485=a_com_rs485_list[meter_position_cur]

            default_value_dict = writeDefaultValue(default_value_dict)
            saveConfigValue('opto_run.json',default_value_dict)
            
            if number_of_meters>1:
                a_sep="="*49+"="*len(str(number_of_meters))+ \
                    "="*len(str(meter_position_cur+1))
                a_txt=f"{a_sep}\nПроверка {number_of_meters} " \
                    "счетчиков на стенде." \
                    f" {bcolors.ATTENTIONGREEN}Текущая позиция" \
                    f" {meter_position_cur+1}.{bcolors.ENDC}"
                a_txt=f"{a_txt}\n{a_sep}"

                a_mode="номер ПУ"

                title_new="Аппаратная проверка ПУ (текущая поз.: " \
                    f"{meter_position_cur+1})"
                replaceTitleWindow("", title_new)
                
            printGREEN(a_txt)

            if meter_test_mode=="конфигурация":
                a_mode="номер только счетчика"
            
            meter_tech_number=""
            meter_serial_number=""
            meter_soft=""
            gsm_serial_number=""
            rc_serial_number=""

            res=startMeterTest(a_mode)

            if res[0]!="1":
                if number_of_meters>1:
                    res=innerQuestionBreakTest()
                    if res[0]=="9":
                        for i in range(0, len(meter_status_test_list)):
                            meter_status_test_list[i]="прервано"
                        return

                    elif res[0]=="91":
                        continue
                    
                    elif res[0]=="92":
                        meter_tech_number=""
                        meter_tech_number_list.append(meter_tech_number)

                        meter_status_test_list.append("пропущен")

                        meter_serial_number=""
                        meter_serial_number_list.append(meter_serial_number)

                        meter_soft=""
                        meter_soft_list.append(meter_soft)

                        gsm_serial_number_list.append("")

                        rc_serial_number_list.append("")

                else:
                    return

            else:
                meter_tech_number_list.append(meter_tech_number)

                meter_status_test_list.append("проверяется")

                meter_serial_number_list.append(meter_serial_number)

                meter_soft_list.append(meter_soft)

                gsm_serial_number_list.append(gsm_serial_number)

                rc_serial_number_list.append(rc_serial_number)

                default_value_dict = writeDefaultValue(default_value_dict)
                saveConfigValue('opto_run.json',default_value_dict)
                
                res=getPathOptoRun(workmode)
                if res[0]=="":
                    return
                opto_run_path =res[2]
                multi_config_dir=res[3]

                opto_run_multi=f"opto_run_{meter_position_cur}.json"
                opto_run_multi_path=os.path.join(multi_config_dir, opto_run_multi)
                res=copyFile(opto_run_path, opto_run_multi_path, "0", "1")
                if res[0]=="0":
                    txt_err = f"{bcolors.FAIL}Ошибка при копировании ф.opto_run.json " \
                        f"в папку 'multi_config'.\n{bcolors.OKBLUE}Нажмите Enter."
                    inputSpecifiedKey("", txt_err, [], "\r", 1, "")
                    return
                
            meter_position_cur+=1

            meter_tech_number_start_list=meter_tech_number_list.copy()

            meter_serial_number_start_list=meter_serial_number_list.copy()

            default_value_dict = writeDefaultValue(default_value_dict)
            saveConfigValue('opto_run.json',default_value_dict)
            continue


        a_meter_on=False
        for a_tech in meter_tech_number_list:
            if a_tech!=None and a_tech!="":
                a_meter_on=True
                break

        if a_meter_on==False:
            return

        a_mode="конфигурация и далее"

        if meter_test_mode=="конфигурация":
            a_mode="конфигурация"

        res=startMeterTest(a_mode)

        if res[0]=="9":
            for i in range(0, len(meter_status_test_list)):
                if meter_status_test_list[i]=="проверяется":
                    meter_status_test_list[i]="прервано"

        if  number_of_meters>1:
            innerPrintTableResultAllTest(meter_test_mode)
            
        return
    

    def innerPrintTableResultAllTest(meter_test_mode: str):

        a_status_sutp_list=[""]*number_of_meters

        tbl_status_list=[]

        if data_exchange_sutp!="0":
            print ("Запрашиваю статус проверенных ПУ в СУТП...")

            a_mode="1"

            a_pw_encrypt=""

            for i in range(0, len(meter_tech_number_start_list)):
                a_tech=meter_tech_number_start_list[i]

                res=getInfoAboutDevice(a_tech, workmode, employee_id, 
                    a_pw_encrypt, a_mode)
                
                if res[0]!="0":
                    a_status_sutp_list[i]=res[7]
        
        summary_table=PrettyTable()
        summary_table.field_names=["Позиция", "Технический номер", 
            "Серийный номер", "Результат", "Статус"]
        
        a_status_dic={"ремонт": ["Дефект", bcolors.FAIL], 
            "годен": ["ОТК пройден", bcolors.OKGREEN],
            "проверяется": ["", bcolors.WARNING],
            "пропущен": ["", bcolors.WARNING],
            "прервано": ["", bcolors.WARNING]}
        
        a_color_def=bcolors.OKGREEN

        for i in range(0, len(meter_tech_number_start_list)):
            a_tech=meter_tech_number_start_list[i]
            if a_tech=="" or a_tech==None:
                continue

            a_serial=meter_serial_number_start_list[i]
            a_serial_space=a_serial[0:-7]+" "+a_serial[-7:len(a_serial)]

            a_status_test=meter_status_test_list[i]

            a_status_sutp=a_status_sutp_list[i]

            a_expected_val=a_status_dic[a_status_test][0]

            a_color_sutp=bcolors.OKGREEN

            a_color_test=a_status_dic[a_status_test][1]

            a_row=[str(i+1), a_tech, a_serial_space, f"{a_color_test}" \
                f"{a_status_test}{bcolors.ENDC}{a_color_def}", 
                f"{bcolors.WARNING}нет данных{bcolors.ENDC}{a_color_def}"]

            if a_status_sutp!="":
                if a_expected_val!=a_status_sutp:
                    a_color_sutp=bcolors.ATTENTIONWARNING

                a_row=[str(i+1), a_tech, a_serial_space, f"{a_color_test}" \
                    f"{a_status_test}{bcolors.ENDC}{a_color_def} ", \
                    f"{a_color_sutp}{a_status_sutp}{bcolors.ENDC}"\
                    f"{a_color_def}"]
                
            tbl_status_list.append(a_row)

        a_txt="Результаты ПОЛНОЙ проверки ПУ:"

        if meter_test_mode=="конфигурация":
            a_txt="Результаты проверки КОНФИГУРАЦИИ ПУ:"

        printGREEN(f"\n{a_txt}")

        summary_table.add_rows(tbl_status_list)
        print(f"{a_color_def}{summary_table}{bcolors.ENDC}")

        return


    os.system("CLS")
    
    readDefaultValue(default_value_dict) 
    
    com_current="com_opto"

    default_value_dict = writeDefaultValue(default_value_dict)
    saveConfigValue(file_name_in="opto_run.json", 
        var_config_dict=default_value_dict)

    multi_com_opto_dic=default_value_dict["multi_com_opto_dic"]
    multi_com_rs485_dic=default_value_dict["multi_com_rs485_dic"]
    multi_com_config_opto_dic=default_value_dict["multi_com_config_opto_dic"]
    multi_com_config_rs485_dic=default_value_dict["multi_com_config_rs485_dic"]

    number_of_meters=default_value_dict["number_of_meters"]
        
    while True:
        if print_number_big_font=="окно":
            a_title_list = ["Печать большим шрифтом", "otk_print_big_font"]
            for a_title in a_title_list:
                res = searchTitleWindow(a_title)
                if res[0] == "1":
                    res=actionsSelectedtWindow([a_title], None,"закрыть", "1")
                    if res[0]=="1":
                        break
      

        os.system("CLS")

        title_new="Аппаратная проверка ПУ"
        replaceTitleWindow("", title_new)

        mass_prod_vers=None

        restoreDefaultValue()

        saveConfigValue(file_name_in="opto_run.json", 
            var_config_dict=default_value_dict)
        
        updateFilesFromList("auto_upd_file.json", False , 
            workmode)
        
        keypress_request_no="1"
        
        com_current="com_opto"
        meter_pw_visible="0"
        default_value_dict = writeDefaultValue(default_value_dict)
        saveConfigValue(file_name_in="opto_run.json", 
            var_config_dict=default_value_dict)

        toPrintDefaultValue()

        res=innerCheckComPortInterface()
        com_opto_no=res[0]
        
        print()

        if workmode!="эксплуатация":
            print(f"{bcolors.WARNING}Программа работает в режиме 'ТЕСТ'.{bcolors.ENDC}")
        txt1="Выберите дальнейшее действие:"
        menu_item_list=["Отправить ПУ в ремонт", "Провести полную проверку ПУ",
            "Изменить конфигурацию программы или значения по умолчанию", 
            "Показать информацию и историю о ПУ", "Изменить количество ПУ",
            "Настройка соединения с ПУ"]
        menu_id_list=["ПУ в ремонт", "полная проверка", 
            "редактировать значения по умолчанию", "информация о ПУ", 
            "изменить количество ПУ", "настройка соединения",]
        
        if meter_config_check!="0":
            menu_item_list.append("Провести проверку только конфигурации ПУ")
            menu_id_list.append("проверка конфигурации ПУ")
        
        if com_opto_no:
            menu_item_list=["Отправить ПУ в ремонт", 
            "Изменить конфигурацию программы или значения по умолчанию", 
            "Показать информацию и историю о ПУ", "Изменить количество ПУ",
            "Настройка соединения с ПУ"]
            menu_id_list=["ПУ в ремонт",
                "редактировать значения по умолчанию", "информация о ПУ", 
                "изменить количество ПУ", "настройка соединения"]

        spec_list=["Выйти из режима полной проверки"]
        spec_keys=["/"]
        spec_id_list=["выход"]
        if workmode!="эксплуатация":
            spec_list.append("Перейти в режим 'ЭКСПЛУАТАЦИЯ'")
            spec_keys.append("*")
            spec_id_list.append("режим эксплуатация")
        
        spec_keys_hidden=["*11*", "*12*", "*13*"]

        a_time_wait=toformatNow()[3]
        oo=questionFromList(bcolors.OKBLUE, txt1, menu_item_list, menu_id_list,
            "", spec_list, spec_keys, spec_id_list, 1, 0, 0, spec_keys_hidden, "")
        print()
        if oo in ["ПУ в ремонт", "полная проверка"]:
            res=changeUserAfterWaiting(default_value_dict, a_time_wait,
                5, workmode)
            if res[0] in ["0","2"]:
                return
            elif res[0]=="1" and employee_id!=res[2]["employee_id"]:
                a_dic=res[2]
                readDefaultValue(a_dic)
                default_value_dict=writeDefaultValue(default_value_dict)
        
        
        if oo=="выход":
            return
        
        elif oo=="ПУ в ремонт":
            os.system("CLS")
            printWARNING("Отправка ПУ в ремонт.")

            startMeterTest("отправка ПУ в ремонт")
        
        elif oo=="редактировать значения по умолчанию":
            changeDefaultValue()
            continue
        
        elif oo=="информация о ПУ" or oo=="*12*":
            output_mode="0"
            if oo=="*12*":
                output_mode="1"
            
            os.system('CLS')

            first_meter=True

            cicl_info_meter=True
            while cicl_info_meter:
                a_txt="\nПолучение информации о ПУ и МС из СУТП."
                if output_mode=="1":
                    a_txt="\nПолучение информации о паролях доступа " \
                        "к ПУ из СУТП."
                print (f"{bcolors.OKGREEN}{a_txt}{bcolors.ENDC}\n")
                if data_exchange_sutp!="1":
                    print (f"{bcolors.WARNING}Отключен обмен данными с СУТП.{bcolors.ENDC}\n"
                        f"{bcolors.WARNING}Получение данных невозможно.{bcolors.ENDC}")
                else:
                    while True:
                        menu_ret="0"
                        txt1 = "Введите или отсканируйте технический или серийный номер"

                        a_dic={True:"ПУ", False: "следующего ПУ"}

                        txt1=f"{txt1} {a_dic[first_meter]}.\n" \
                            "Чтоб прервать ввод - нажмите '/'."
                    
                        len_num_list=[9, 13, 15]
                        spec_key_list=['/']
                        res=inputNumberDevice(len_num_list, "счетчик э/э", "Product1", 
                            txt1, spec_key_list, print_number_big_font, workmode)
                        if res[0] in ["0", "2", "3"]:
                            menu_ret="1"

                            break
                        
                        first_meter=False
                        
                        device_number=res[1]

                        user_id=""
                        user_pw_encrypt=""
                        query_mode="2"
                        if output_mode=="1":
                            user_id=employee_id
                            user_pw_encrypt=employee_pw_encrypt
                            query_mode="0"
                        res_device=getInfoAboutDevice(device_number, workmode, 
                            user_id, user_pw_encrypt, query_mode)
                        if res_device[0] in ["1", "2"]:
                            break
                        print(f"{bcolors.WARNING}Не удалось получить информацию о ПУ № {device_number}."
                            f"{bcolors.ENDC}")
                        a_dic={9: "серийный", 13: "технический", 15: "технический"}
                        a_txt=a_dic.get(len(device_number),"-")
                        print(f"{bcolors.WARNING}Попробуйте ввести {a_txt} номер ПУ.{bcolors.ENDC}")
                    if menu_ret=="1":
                        cicl_info_meter=False

                        keypress_request_no="0"
                        break

                    meter_tech_number=str(res_device[2])
                    meter_serial_number=res_device[3]

                    if output_mode=="1":
                        a_meter_pw_hl=res_device[10]
                        a_meter_pw_ll=res_device[19]
                        if a_meter_pw_hl!="" or a_meter_pw_ll!="":

                            os.system('CLS')
                            
                            print(f'{bcolors.OKGREEN}Информация о ПУ:{bcolors.ENDC}')
                            print (f'{bcolors.OKGREEN}- технический номер: {meter_tech_number}{bcolors.ENDC}')
                            print (f'{bcolors.OKGREEN}- серийный номер: {meter_serial_number}{bcolors.ENDC}')
                            print (f'{bcolors.OKGREEN}- пароль доступа к ПУ верхнего уровня: '
                                f'{a_meter_pw_hl}{bcolors.ENDC}')
                            print (f'{bcolors.OKGREEN}- пароль доступа к ПУ нижнего уровня: '
                                f'{a_meter_pw_ll}{bcolors.ENDC}')
                            if a_meter_pw_hl!="":
                                pyperclip.copy(a_meter_pw_hl)
                                print (f"{bcolors.OKGREEN}Пароль доступа к ПУ верхнего уровня " 
                                    f"скопирован в буфер обмена Windows.{bcolors.ENDC}")
                        
                    else:
                        a_status=res_device[7]
                        a_color=bcolors.OKGREEN
                        if a_status not in ["Гравировка пройдена", "Состыкован с МС"]:
                            a_color=bcolors.WARNING
                        
                        os.system('CLS')

                        print(f'{bcolors.OKGREEN}Информация о ПУ:{bcolors.ENDC}')
                        print (f'{bcolors.OKGREEN}- технический номер: {meter_tech_number}{bcolors.ENDC}')
                        print (f'{bcolors.OKGREEN}- серийный номер: {meter_serial_number}{bcolors.ENDC}')
                        print (f'{bcolors.OKGREEN}- модель: {res_device[12]}{bcolors.ENDC}')
                        print (f'{bcolors.OKGREEN}- версия ПО: {res_device[17]}{bcolors.ENDC}')
                        print (f'{bcolors.OKGREEN}- статус в СУТП:{bcolors.ENDC} {a_color}{a_status}{bcolors.ENDC}')
                        print (f'{bcolors.OKGREEN}- номер заказа: {str(res_device[8])}{bcolors.ENDC}')
                        print (f'{bcolors.OKGREEN}- описание заказа: {res_device[9]}{bcolors.ENDC}')
                        if res_device[13]!=None:
                            print(f'\n{bcolors.OKGREEN}Информация о состыкованном МС:{bcolors.ENDC}')
                            print(f'{bcolors.OKGREEN}- технический номер: {str(res_device[13])}{bcolors.ENDC}')
                            print(f'{bcolors.OKGREEN}- серийный номер: {str(res_device[14])}{bcolors.ENDC}')
                            print (f'{bcolors.OKGREEN}- модель: {res_device[15]}{bcolors.ENDC}')
                            print (f'{bcolors.OKGREEN}- версия ПО: {res_device[16]}{bcolors.ENDC}')
                        else:
                            print(f'{bcolors.OKGREEN}ПУ без МС.{bcolors.ENDC}')


                        print (f"\n{bcolors.OKGREEN}История присвоения серийных номеров ПУ № " \
                            f"{device_number}:{bcolors.ENDC}")
                        res=getMeterAllSN(device_number, workmode, "0")
                        if res[0]=="1":
                            txt=res[3]
                            if txt=="":
                                txt="нет данных"
                            print (txt)
                        

                        print (f"\n{bcolors.OKGREEN}История движения ПУ № {device_number}:{bcolors.ENDC}")
                        res=getDeviceHistory(device_number, employee_print="1", workmode=workmode)
                        if res[0]=="1":
                            txt=res[2]
                            if txt=="":
                                txt="нет данных"
                            print (txt)
                        
                        
                        print (f"\n{bcolors.OKGREEN}История ремонта ПУ № {device_number}:{bcolors.ENDC}")
                        res=getDeviceRepayHistory(device_number, workmode=workmode)
                        txt=""
                        if res[0]=="1":
                            txt=res[2]
                            if txt=="":
                                txt="нет данных"
                            print (txt)

                           
        elif oo == "изменить количество ПУ":
            changeNumOfMeters()
            continue
        
        elif oo=="настройка соединения":
            connectionMeterSetup()
            continue
        
        elif oo=="режим эксплуатация":
            workmode="эксплуатация"
            default_value_dict = writeDefaultValue(default_value_dict)
            saveConfigValue('opto_run.json',default_value_dict)
            continue
        
        elif oo=="полная проверка" or oo=="*11*":
            if oo=="*11*":
                meter_pw_visible="1"
           
            innerMeterTest("полная проверка")

        elif oo=="проверка конфигурации ПУ" or oo=="*11*":
          
            innerMeterTest("конфигурация")

        elif oo=="*13*":
            os.system('CLS')
            a_txt="\nРасшифровка строки"
            printGREEN(a_txt)

            while True:
                
                txt1 = "Введите строку с зашифрованным текстом.\n" \
                    "Для выхода в меню нажмите '/'."
                oo=inputSpecifiedKey(bcolors.OKBLUE, txt1, "", [0], ["/"], 0,
                    "", "откл", workmode)
                
                if oo=="/" or oo=="":
                    keypress_request_no="0"

                    break

                res=cryptStringSec("расшифровать", oo)
                if res[0]=="1":
                    a_pw=res[2]
                    printGREEN(f'\nРасшифрованный текст: {a_pw}')

                    pyperclip.copy(a_pw)
                    printGREEN ("Расшифрованный текст скопирован в буфер "
                                "обмена Windows.\n")
                    

        print()
        if keypress_request_no=="1":
            txt1="Для возврата в меню - нажмите Enter."
            keystrokeEnter(txt1)
            


def restoreDefaultValue():
    global default_value_dict

    a_list=["sutp_to_save", "meter_config_check"]
    for a_var in a_list:
        if len(globals()[a_var])==2:
            globals()[a_var]=globals()[a_var][1]
    
    if globals()["electrical_test_circuit"][0]== "0":
        globals()["electrical_test_circuit"]="1-0"

    default_value_dict = writeDefaultValue(default_value_dict)
    saveConfigValue('opto_run.json',default_value_dict)
    
    return



def startPrintBigFont():

    a_file_name="otk_print_big_font.bat"
    _, ans2, a_file_path = getUserFilePath(a_file_name,
        workmode=workmode)
    if a_file_path == "":
        return ["0", f"Ошибка в ПП getUserFilePath(): {ans2}"]
    
    txt1 = "Пожалуйста подождите, идет загрузка " \
        "программы для отображения вводимых номеров большим шрифтом...\n"
    printGREEN(txt1)
    subprocess.Popen(f"start {a_file_path}", shell=True)

    return ["1", "Программа успешно запущена"]



def startMeterTest(modeStartMeterTest: str):

    global default_value_dict  # словарь со зн-ями по умолчанию:
    global employees_name       #ФИО пользователя
    global employee_id          # таб. номер пользователя
    global employee_pw_encrypt  # зашифрованный пароль пользователя для доступа в СУТП
    global rep_copy_public      #метка возможности копирования протокола в общую папку:
    global speaker              #метка вкл/откл диктора. "0"-откл, "1"-вкл
    global res_ext_at_begin_test #метка записи результата внешнего осмотра ПУ в начале 
    global number_of_meters     #количество подключаемых ПУ
    global meter_position_cur   #номер текущей позиции ПУ на стенде
    global com_opto             #COM-порт, к которому подключен оптопорт
    global com_rs485            #COM-порт, к которому подключен RS-485
    global com_current          #вид активного порта для связи с ПУ:"com_opto","com_rs485"
    global com_config_opto      # COM-порт, к которому подключен оптопорт
    global com_config_rs485     # COM-порт, к которому подключен RS-485
    global com_config_current   #вид активного порта для связи с ПУ при проверке 
    global com_config_current_select    #способ выбора COM-порта для проверки
    global multi_com_opto_dic   #словарь со списками используемых COM-портов
    global multi_com_rs485_dic   #словарь со списками используемых COM-портов
    global multi_com_config_opto_dic   #словарь со списками используемых COM-портов
    global multi_com_config_rs485_dic   #словарь со списками используемых COM-портов
    global com_config_eqv_com   #метка использования для проверки конфигурации ПУ
    global config_send_mail     #отправка сообщения по электронной почте о
    global rep_err_send_mail    #отправка сообщения по электронной почте о
    global meter_color_body     #цвет корпуса ПУ по умолчанию
    global meter_config_check   #метод проверки конфигурации ПУ: "0"-откл.,
    global meter_config_res_list    #список с замечаниями по проверке конфигурации ПУ
    global meter_adjusting_clock    #корректировка часов счетчика: 
    global meter_type_def       # тип ПУ по умолчанию
    global meter_serial_number  #серийный номер ПУ (эталон). Порядок выбора:СУТП, электронный паспорт,
    global meter_serial_number_list #список серийных номеров ПУ, проверяемых на стенде
    global meter_sn_source      #источник получения эталонного серийного номера (СУТП, электронный паспорт, QR-код/надпись)                               
    global meter_sn_lbl         #серийный номер ПУ, считанный c QR-кода или на корпусе
    global meter_sn_ep          #серийный номер ПУ из электронного паспорта ПУ

    global meter_tech_number    #технический номер ПУ (эталон). Порядок выбора: наклейка на крышке, СУТП
    global meter_tech_number_list #список технических номеров ПУ, проверяемых на стенде
    global meter_tn_source      #источник получения технического номера (наклейка, СУТП)
    global meter_tn_lbl         #технический номер ПУ, считанный c наклейки
    global meter_status_test_list   #список статусов ПУ при прохождении проверки:
    
    global meter_pw_default     #пароль верхнего уровня по умолчанию для подключения к ПУ
    global meter_pw_encrypt     #текущий зашифрованный пароль подключения к ПУ
    global meter_pw_descript    #описание текущего пароля подключения к ПУ
    global meter_pw_level       # уровень текущего доступа к ПУ: "High", "Low"
    global meter_pw_default_descript  #описание пароля по умолчанию ("Стандартный верхнего уровня", "Карелия"...) для подключения к ПУ
    global meter_pw_low_encrypt     #зашифрованный пароль нижнего уровня подключения к ПУ
    global meter_pw_low_descript     #описание текущего пароля нижнего уровня подключения к ПУ
    global meter_pw_high_encrypt     #зашифрованный пароль верхнего уровня подключения к ПУ
    global meter_pw_high_descript     #описание пароля верхнего уровня подключения к ПУ
    global meter_pw_visible   #вывод на экран паролей доступа к ПУ: "0"-откл, "1"-вкл.
    global meter_phase          #число фаз у ПУ
    global meter_voltage_dic    #словарь с мгновенными значениями напряжения
    global meter_amperage_dic   #словарь с мгновенными значениями тока
    global meter_soft           #версия ПО ПУ


    global electrical_test_circuit  #схема подключения ПУ для проверки: 
    global ctrl_current_electr_test #контрольное значение тока в схеме подключения ПУ

    global modem_type_def       # тип модуля связи по умолчанию
    global modem_status         #статус модема по умолчанию: 
    global actions_no_mc        # действия при отсутствии обязательного модуля связи
    global SIMcard_status       #Статус SIM-карты по умолчанию: 
    global gsm_serial_number    #номер GSM модема (эталон). Порядок выбора:крышка
    global gsm_SIM_number       #номер SIM-карты ("0"-номер не требуется)

    global filename_rep             #имя файла-отчета без префикса "_отчет.txt"
    global default_filename_full     #имя отчета (протокола) вместе с именем тек. директории
    global workmode             #метка режима работы программы "тест" - режим теста1, 
    global data_exchange_sutp   #метка обмена данными с СУТП:"0"-откл.,"1"-вкл.
    global sutp_to_save         #способ записи рез-тов теста в БД СУТП ("0"-отключен, "1"-ручной,
    global order_control        # метка контроля принадлежности ПУ определенному заказу 
    global order_control_descript   # номер и описание контролируемого заказа
    global order_num            #номер заказа проверяемого ПУ
    global order_descript       #описание заказа проверяемого ПУ
    global order_ev
            
    global rep_err_list         #список с выявленными ошибками для записи в отчет
    global clipboard_err_list   #список выявленных ошибок для сохранения в буфере Windows
    global rep_remark_list      #список с примечаниями для записи в отчет
    global test_start_time      #время начала проведения проверки ПУ во внутреннем формате
    global duration_test        #продолжительность проверки ПУ, мин
    global rc_serial_number     #серийный номер ПДУ
    global print_number_big_font    # метка печать на экране вводимых номеров изделий
    global mass_prod_vers       # версия программы MassProd, которая должна использоваться 



    def innerCheckStatusSUTP():

        if data_exchange_sutp=="0":
            return ["4", "Обмен данными с СУТП отключен.", "", [], []]


        device_num=meter_tech_number
        if device_num=="":
            device_num=meter_serial_number
        if device_num=="":
            return["0","Отсутствуют номера ПУ.", "", [], []]
        
        err_list=[]

        rem_list=[]
        
        status_txt=""
        
        res=preChecksToGhangeStatusMeter(device_num, device_status_id=21,
            workmode=workmode, print_err_msg="0")
        status_txt=res[2]
        txt=f"Текущий статус ПУ в СУТП: {status_txt}"
        
        if res[0]=="1":
            if not status_txt in ["Гравировка пройдена",
                "Состыкован с МС"]:
                res=getDevicePw(device_num, employee_id, employee_pw_encrypt, 
                    workmode, "1")
                if res[1]=='Статус ИПУ не подразумевает получение пароля':
                    
                    if print_number_big_font=="окно":            
                        a_dic={"text_3": "ERR STATUS"}
                        saveConfigValue("print_big_font_line.json", a_dic,
                            workmode, "заменить часть")
                        
                        time.sleep(3)

                        a_title_list = ["Печать большим шрифтом", "otk_print_big_font"]
                        actionsSelectedtWindow(a_title_list, None,"закрыть", "1")

                    
                    err_txt=f"Статус ПУ '{status_txt}' не позволяет получить в СУТП " \
                        "пароли доступа к ПУ для проведения его проверки."
                    
                    err_list.append(err_txt)

                    txt="Нельзя будет получить из СУТП пароли доступа к ПУ."
                    menu_item_list=["Отправить ПУ на ремонт", 
                        "Продолжить проверку с фиксацией замечания в списке дефектов"]
                    menu_id_list=["ремонт", "продолжить"]
                    spec_list=["Прервать проверку"]
                    spec_keys=["/"]
                    spec_id_list=["прервать"]
                    oo=questionFromList(bcolors.OKBLUE, txt, menu_item_list, menu_id_list,
                        "", spec_list, spec_keys, spec_id_list, 1, start_list_num=1)
                    print()
                    if oo=="прервать":
                        printWARNING ("Проверка прервана.")
                        return ["9", "Проверка прервана пользователем.", status_txt, err_list]
                    
                    elif oo=="продолжить":
                        if print_number_big_font=="окно":
                            res=startPrintBigFont()
                            if res[0]=="0":
                                return ["0", res[1], err_list]

                    
                    elif oo=="ремонт":
                        return ["6", "Отправить ПУ в ремонт.", status_txt, err_list]

            printGREEN(txt)

            if len(err_list)>0:
                return ["5", "Продолжить проверку с фиксацией замечания в "
                    "списке дефектов.", status_txt, err_list]
            
            else:
                return ["1", "У ПУ допустимый статус.", status_txt, err_list]
        
        elif res[0]=="2":
            txt="Не удалось получить информацию о текущем статусе ПУ в СУТП."
            print(f"{bcolors.WARNING}{txt}{bcolors.ENDC}")
            with open(default_filename_full, "a", errors="ignore") as file:
                file.write(f"{txt}\n")

            if len(err_list)>0:
                return ["5", "Продолжить проверку с фиксацией замечания в "
                    "списке дефектов.", status_txt, err_list]
            
            return ["2", "Не удалось получить информацию о статусе ПУ.", "", []]
        
        elif res[0]=="3":
            if sutp_to_save[0]!="0":
                if print_number_big_font=="окно":            
                    a_dic={"text_3": "ERR STATUS"}
                    saveConfigValue("print_big_font_line.json", a_dic,
                        workmode, "заменить часть")
                    
                    time.sleep(3)

                    a_title_list = ["Печать большим шрифтом", "otk_print_big_font"]
                    actionsSelectedtWindow(a_title_list, None,"закрыть", "1")
                    
                printWARNING(txt)
                txt1_1=f"{bcolors.WARNING}Нельзя будет внести результаты проверки ПУ в СУТП.{bcolors.ENDC}\n" \
                    f"{bcolors.OKBLUE}Прервать проверку? 0 - нет, 1 - да.{bcolors.ENDC}"
                oo = questionSpecifiedKey("", txt1_1,["0","1"],"")
                print()
                if oo=="1":
                    return ["9", "Пользователь прервал дальнейшую проверку.", 
                        status_txt, err_list]
                
                else:
                    if print_number_big_font=="окно":
                        res=startPrintBigFont()
                        if res[0]=="0":
                            return ["0", res[1], err_list]
                    
                    if len(err_list)>0:
                        return ["5", "Продолжить проверку с фиксацией замечания в "
                            "списке дефектов.", status_txt, err_list]
                    
            return ["3", "Статус ПУ нельзя будет изменить.", status_txt, err_list]
                

    
    def innerCheckConfigFilename():

        if data_exchange_sutp=="0":
            return ["3", "Обмен данными с СУТП отключен.", ""]

        err_txt=""

        if  meter_tech_number==None or meter_tech_number=="" or \
            meter_config_check[0]=="0":
            return ["5", "Сверка имен файлов конфигурации ПУ не проводилась.", ""]
        
        res=getMeterConfigFilePath(meter_tech_number, "1", workmode, 
            "0", None)

        if res[0]!="0":
            printGREEN("Имя файла конфигурации ПУ, указанного в заказе: "
                f"{res[3]}.")
            
            printGREEN("Имя файла, использованного при конфигурировании ПУ: " \
                    f"{res[4]}.")
            
        if res[0] in ["0", "3"]:

            if print_number_big_font=="окно":            
                a_dic={"text_3": "ERR CONFIG"}
                saveConfigValue("print_big_font_line.json", a_dic,
                    workmode, "заменить часть")
                
                time.sleep(3)

                a_title_list = ["Печать большим шрифтом", "otk_print_big_font"]
                actionsSelectedtWindow(a_title_list, None,"закрыть", "1")

            err_txt="Не удалось загрузить файл конфигурации из СУТП."

            if res[0]=="3":
                err_txt="Имя файла, примененного для конфигурирования ПУ " \
                    f"'{res[4]}' отличается от имени файла, указанного в заказе " \
                    f"'{res[3]}'."

            menu_item_list=["Отправить ПУ на ремонт", 
                "Продолжить проверку с фиксацией замечания в списке дефектов",
                "Продолжить проверку с фиксацией замечания в списке примечаний"]
            menu_id_list=["ремонт", "продолжить", "записать в примечания"]

            spec_list=["Прервать проверку"]
            spec_keys=["/"]
            spec_id_list=["прервать"]
            oo=questionFromList(bcolors.OKBLUE, err_txt, menu_item_list, menu_id_list,
                "", spec_list, spec_keys, spec_id_list, 1, start_list_num=1)
            print()
            if oo=="прервать":
                printWARNING ("Проверка прервана.")
                return ["9", "Проверка прервана пользователем.", err_txt]
            
            elif oo=="продолжить":
                if print_number_big_font=="окно":
                    res=startPrintBigFont()
                    if res[0]=="0":
                        return ["0", res[1], err_txt]
                
                return ["2", "Записать замечание в список с дефектами.", err_txt]

            elif oo=="записать в примечания":
                return ["3", "Записать замечание в список с примечаниями.", err_txt]
            
            elif oo=="ремонт":
                return ["4", "Отправить ПУ в ремонт.", err_txt]
            
        elif res[0]=="1":
            return ["1", "Имена файлов конфигурации ПУ совпадают", ""]


    
    def innerMakeFileName(meter_sn_in):
        nonlocal default_filename_old
        global filename_rep
        nonlocal filename_report
        global default_filename_full
        nonlocal filename_ext
        nonlocal default_filename_ext

        meter_sn="-"
        meter_tn="-"
        if meter_sn_in!="":
            meter_sn=meter_sn_in
        if meter_tech_number!="":
            meter_tn=meter_tech_number

        
        a_now_full=now_full
        if a_now_full=="":
            a_now_full=toformatNow()[0]

        pc_time2 = a_now_full.replace(" ", "_").replace(":", "")
        default_filename_old=default_filename_full
        filename_rep=f"{meter_sn}_{meter_tn}_{pc_time2}"

        filename_report=filename_rep+"_отчет.txt"
        default_filename_full=os.path.join(default_dirname, filename_report)
        filename_ext=filename_rep+"_внеш.осмотр.txt"
        default_filename_ext =os.path.join(default_dirname, filename_ext)
        return



    def innerReadMeter(read_mode="1"):

        global default_value_dict   #словарь со значениями по умолчанию
        global meter_sn_ep          #серийный номер ПУ из электронного паспорта
        global com_current          #имя активного порта
        global com_opto             #номер порта для оптопорта
        global com_rs485            #номер порта для RS-485
        global meter_phase          #число фаз у ПУ
        global meter_voltage_dic    #словарь с мгновенными значениями напряжения
        global meter_amperage_dic   #словарь с мгновенными значениями тока
        global meter_soft           #версия ПО ПУ

        _, ans2, file_name_check = getUserFilePath('otk_check_opto.py',
            workmode=workmode)
        if file_name_check == "":
            txt1=f"{bcolors.WARNING}Не удалось найти путь к ф.'otk_check_opto.py'.{bcolors.ENDC}" \
                f"{bcolors.WARNING}Продолжение проверки невозможно.{bcolors.ENDC}\n" \
                f"{bcolors.OKBLUE}Нажмите 'Enter'.{bcolors.ENDC}"
            spec_keys=["\r"]
            oo=inputSpecifiedKey("", txt1, "", [0], spec_keys, 1)
            return "0"


        a_com_dic={"com_opto": "оптопорт", "com_rs485": "RS-485"}
        print(f"Запрашиваю у ПУ данные через {a_com_dic[com_current]}...")
        a_mode=int(read_mode)
        code_exit = os.system(f"python {file_name_check} 0 {a_mode}")
        res_1 = ExchangeBetweenPrograms(operation="read",
            recipient="otk_menu_result")
        if res_1[0] == "1":
            rec_dict_1 = res_1[2]
            dt_second = float(rec_dict_1.get("dateTime"))
            source_1 = rec_dict_1.get("source")
            operation = rec_dict_1.get("operation")
            result = rec_dict_1.get("result")
            pc_second = toformatNow()[3]
            a_ret=0
            if pc_second-dt_second < 30 and source_1 == "otk_check_opto" and \
                    operation == "checkPassword": 
                if result=="3":    
                    res=readGonfigValue(file_name_in="opto_run.json",
                        var_name_list=[],default_value_dict=default_value_dict)
                    if res[0]!="1":
                        txt1=f"{bcolors.WARNING}При получении серийного номера из файла " \
                            f"возникла ошибка.{bcolors.ENDC}\n" \
                            f"{bcolors.OKBLUE}Нажмите Enter{bcolors.ENDC}"
                        print (txt1)
                        oo=questionSpecifiedKey("","",["\r"],"",1)
                        return "0"
                    default_value_dict=res[2]
                    com_opto=default_value_dict.get("com_opto","")
                    com_rs485=default_value_dict.get("com_rs485","")
                    meter_sn_ep=default_value_dict["meter_sn_ep"]

                    if read_mode=="2":
                        meter_phase=default_value_dict["meter_phase"]
                        meter_voltage_dic=default_value_dict["meter_voltage_dic"]
                        meter_amperage_dic=default_value_dict["meter_amperage_dic"]
                        meter_soft=default_value_dict["meter_soft"]

                    a_ret=1
            if a_ret==0:
                txt1=f"{bcolors.WARNING}Не удалось получить данные из памяти ПУ: " \
                    f"серийный номер, мгновенное значение напряжений/тока." \
                    f"{bcolors.ENDC}\n" \
                    f"{bcolors.OKBLUE}Нажмите Enter{bcolors.ENDC}"
                print (txt1)
                oo=questionSpecifiedKey("","",["\r"],"",1)
                return "0"
            return "1"


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
        global rep_err_send_mail

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
        dt_sec=""

        reestr_clipboard_err_txt=",".join(clipboard_err_list)
    
        rep_remark_txt=", ".join(rep_remark_list)
        delta_pc_minus_device_txt=str(delta_pc_minus_device)

        reestr_key_list=["dt_sec","test_start_time", "meter_type",
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
            "reestr_clipboard_err_txt", "rep_remark_txt", "order_ev"]

        reestr_val_list=[dt_sec,test_start_time, meter_type,
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
            reestr_clipboard_err_txt, rep_remark_txt, order_ev]

        reestr_dic=dict.fromkeys(reestr_key_list, "")
        i=0
        for key in reestr_key_list:
            reestr_dic[key]=reestr_val_list[i]
            i+=1

        return reestr_dic

    
    
    def innerSendForRepair(err_txt=""):

        global meter_serial_number  #серийный номер ПУ

        if data_exchange_sutp=="1":
            device_number=meter_tech_number
            a_mode="1"
            print ("Запрашиваю из БД СУТП информацию о ПУ...")
            res=getInfoAboutDevice(device_number, workmode, "", 
                "", a_mode)
            if res[0] in ["1", "2"]:
                meter_serial_number=res[3]

            else:
                print(f"\n{bcolors.WARNING}Не удалось получить информацию " 
                      f"о ПУ из СУТП.{bcolors.ENDC}")
                return
            
        else:
            while True:
                txt1 = "Введите или отсканируйте серийный номер ПУ с его крышки." \
                    "\nЧтобы вернуться в меню нажмите '/'."
                spec_key_list=["/"]
                len_num_list=[13, 15]
                res=inputNumberDevice(len_num_list, "счетчик э/э", "Product1", 
                    txt1, spec_key_list, print_number_big_font, workmode)
                if res[0] in ["0", "3"]:
                    return

                elif res[0]=="2" and res[1]=="/":
                    return

                meter_serial_number=res[1]
                break


        innerMakeFileName(meter_serial_number)
        os.rename(default_filename_old, default_filename_full)

        err_list=[]
        if err_txt!=None and err_txt!="":
            err_list=err_txt.split("\n")
        res=innerSelectActions("", err_list, "9", "прочие", [], [], [])
        
        return

 
    
    def innerCheckPassProdVers(check_meter_soft: str, check_sn: str):

        global mass_prod_vers
    
        a_mass_prod=None
        
        res_info=toGetProductInfo2(check_sn, "Product1")
        if res_info[0]=="0":
            a_err_txt = f"Ошибка при получении информации о ПУ № {check_sn} из " \
                "ф.ProductNumber.xlsx."
            printFAIL(a_err_txt)
            return ["0", a_err_txt]
    

        a_res_cmp=cmpVers(check_meter_soft, "4.14.10")
        if a_res_cmp=="<":
            a_mass_prod=res_info[32]
        
        else:
            a_mass_prod=res_info[33]

        if a_mass_prod==None or a_mass_prod=="":
            a_err_txt="В ф.ProductNumber.xlsx отсутствует информация о " \
                'версии программы "MassProdAutoConfig.exe" для проверки ' \
                f'конфигурации\n ПУ № {check_sn} .'
            printFAIL(a_err_txt)
            return ["0", a_err_txt]
        
        if mass_prod_vers!=None:
            if number_of_meters>1 and mass_prod_vers!=a_mass_prod:
                a_err_txt=f"Конфигурацию ПУ № {check_sn} нельзя будет проверить " \
                    "в одно время с предыдущими ПУ,\nт.к. у них отличаются " \
                    "версии программы 'MassProdAutoConfig.exe'."
                printWARNING(a_err_txt)
                return ["2", a_err_txt]
        
        else:
            mass_prod_vers=a_mass_prod
        
        return ["1", "Версия программы 'MassProdAutoConfig.exe' определена."]


   
    _, _, default_dirname=getUserFilePath(file_name="otk_report",
        only_dir="1",workmode=workmode)
    if default_dirname == "":
        sys.exit()
    
    default_filename_full = os.path.join(default_dirname, 'otk2.txt')
    
    _, _, work_dirname=getUserFilePath(file_name="work_dirname",
        only_dir="1",workmode=workmode)
    if work_dirname == "":
        sys.exit()
    
    _, _, dirname_sos=getUserFilePath(file_name="sharedFolder",
        only_dir="1",workmode=workmode)
    if dirname_sos=="":
        sys.exit()
    
    with open(default_filename_full, "w", errors="ignore") as file:
        pass

    modem_status=""

    res=readGonfigValue("opto_run.json",[],{}, workmode, "1")
    if res[0]!="1":
        printFAIL("При загрузке данных из конфигурационного файла " 
            "'opto_run.json' произошла ошибка.")
        printBLUE("Нажмите Enter.")
        questionSpecifiedKey("", ["\r"], "", 1)
        return ["0", "Ошибка при загрузке конфигурационны данных."]
        
    default_value_dict=res[2]
    readDefaultValue(default_value_dict)

    meter_serial_number_list=default_value_dict["meter_serial_number_list"]

    meter_tech_number=""        #технический номер ПУ (эталон)
    meter_tn_source=""
    meter_tn_lbl=""
    
    meter_serial_number=""
    meter_type=""
    meter_sn_source=""
    meter_sn_lbl=""
    meter_sn_ep=""
    meter_tn_sutp=""
    meter_sn_sutp=""

    meter_status_name=""

    meter_type=""
    meter_date_of_manufacture=""

    meter_pw_high_encrypt=""
    meter_pw_high_descript=""
    meter_pw_low_encrypt=""
    meter_pw_low_descript=""
    
    gsm_serial_number=""
    gsm_soft=""
    mc_on_board_set=None
    gsm_product_type=""
    delta_pc_minus_device=0
    otnoshenie_str=""
    meter_grade=""
    rc_serial_number=""
    rc_soft=""
    meter_soft_sutp=""
    meter_model_sutp=""
    gsm_docked_tn_sutp=""
    gsm_docked_sn_sutp=""
    gsm_docked_model_sutp=""
    gsm_docked_soft_sutp=""
    duration_test=""
    meter_config_res_list=[]

    a_dic={"meter_config_res_list": []}
    saveConfigValue("opto_run.json", a_dic, workmode)

    com_current="com_opto"
    default_value_dict = writeDefaultValue(default_value_dict)
    saveConfigValue('opto_run.json',default_value_dict)


    order_num=""
    order_descript=""
    order_ev="0"
    order_pw_dic={"pw_type_reader_id": None,
        "pw_type_reader_descript": "",
        "pw_reader_assigned": "",
        "pw_reader_encrypt": "", 
        "pw_type_config_id": None, 
        "pw_type_config_descript": "",
        "pw_config_assigned": "",
        "pw_config_encrypt": ""}
    

    default_filename_old=""
    filename_report=""
    filename_ext=""
    default_filename_ext =""

    rep_err_list=[]
    rep_remark_list=[]
    clipboard_err_list=[]

    mass_res_multi_dic={}

    now_full=""
    test_start_time=""

    if modeStartMeterTest in ["номер ПУ", "номер только счетчика", "отправка ПУ в ремонт"]:
        if print_number_big_font=="окно":
            res=startPrintBigFont()
            if res[0]=="0":
                return ["0", res[1]]

        while True:
            txt1=f"{bcolors.OKBLUE}Введите или отсканируйте технический номер ПУ.{bcolors.ENDC}"
            spec_key_list=['/']

            if modeStartMeterTest!="отправка ПУ в ремонт":
                txt1_1 = f"{bcolors.OKGREEN}Установите оптопорт, подключенный к " \
                    f"{bcolors.ATTENTIONGREEN} {com_opto} {bcolors.ENDC} " \
                    f"{bcolors.OKGREEN}на ПУ.\n"
                
                if com_rs485!="" and com_rs485!=None:
                    txt1_1 =txt1_1+f"{bcolors.OKGREEN}Для ПУ серии СПЛИТ - присоедините к ПУ " \
                    f"преобразователь RS-485, подключенный к " \
                    f"{bcolors.ATTENTIONGREEN} {com_rs485} {bcolors.ENDC}{bcolors.OKGREEN}.\n"

                txt1_1=txt1_1+f"{bcolors.OKGREEN}Подайте напряжение на ПУ.{bcolors.ENDC}\n"
                
                if not "номер" in modeStartMeterTest:
                    txt1_1=txt1_1+f"{bcolors.OKBLUE}Если это невозможно, то введите 0 и " \
                    f"нажмите Enter.\nПри этом нельзя будет проверить " \
                    f"конфигурацию ПУ и сохранить результаты проверки " \
                    f"в СУТП.{bcolors.ENDC}"
                    spec_key_list=['0', '/']
                
                txt1=f"{txt1_1}{bcolors.OKBLUE}Чтобы прервать проверку ПУ - нажмите '/'.\n" \
                    f"{txt1}"
                
            else:
                txt1=f"{txt1}\n{bcolors.OKBLUE}Чтобы прервать ввод - нажмите '/'."
                
            len_num_list=[9]

            a_time_wait=toformatNow()[3]

            if print_number_big_font=="окно":
                a_dic={"input_end_list": ["\r"],
                    "input_spec_list": spec_key_list,
                    "input_max_numb_list": len_num_list,
                    "text_1": "T/N  METER", 
                    "text_2": "", 
                    "text_3": "",
                    "input_text": ""}
                
                if number_of_meters>1:
                    a_dic["text_1"]=f"pos. {meter_position_cur+1}\n"+ \
                        a_dic["text_1"]
                    
                saveConfigValue("print_big_font_line.json", a_dic,
                    workmode, "заменить часть")
                a_title = 'Печать большим шрифтом'
                actionsSelectedtWindow([a_title], None,"показать+активировать", 
                    "1")
            
            res=inputNumberDevice(len_num_list, "счетчик э/э", "Product1", 
                txt1, spec_key_list, print_number_big_font, workmode)
            print()

            if res[0]!="1" and  print_number_big_font=="окно":    
                a_title_list = ["Печать большим шрифтом", "otk_print_big_font"]
                actionsSelectedtWindow(a_title_list, None,"закрыть", "1")

            if res[0] in ["0", "3"]:
                return ["0", res[1]]

            elif res[0]=="2" and res[1]=="/":
                return ["9", "Прервали проверку ПУ."]

            elif res[0]=="2" and res[1]=="0":
                meter_tn_lbl=""
                print (f"\n{bcolors.WARNING}У ПУ не введен технический номер.{bcolors.ENDC}\n" \
                    f"{bcolors.WARNING}Проверка конфигурации ПУ будет отключена и нельзя{bcolors.ENDC}\n" \
                    f"{bcolors.WARNING}будет из программы сохранить результаты проверки в СУТП.{bcolors.ENDC}")
                a_dic={"0":"0", "1":"01", "2":"02"}
                sutp_to_save=a_dic.get(sutp_to_save,"0")

                a_dic={"0":"0", "1":"01", "2":"02"}
                meter_config_check = a_dic.get(meter_config_check, "0")
                default_value_dict = writeDefaultValue(default_value_dict)
                saveConfigValue('opto_run.json',default_value_dict)
                res=getSutpToSaveDescript(sutp_to_save)

                printColor (f"{res[3]}{res[2]}\n")
                break
        
            elif res[0]=="1":
                meter_tn_lbl=res[1]
                meter_tech_number=meter_tn_lbl
                meter_tn_source="наклейка"
                break
            

        res=changeUserAfterWaiting(default_value_dict, a_time_wait,
            5, workmode)
        if res[0] =="0":
            return  ["0", res[1]]
        
        elif res[0]=="2":
            return ["9", res[1]]

        elif res[0]=="1" and employee_id!=res[2]["employee_id"]:
            a_dic=res[2]
            readDefaultValue(a_dic)
            default_value_dict=writeDefaultValue(default_value_dict)
        

        now_full=toformatNow()[0]
        test_start_time=now_full
        
        
        if modeStartMeterTest=="отправка ПУ в ремонт":
            innerSendForRepair()
            return ["1", "ПУ успешно отправлен в ремонт."]
        
        if meter_tech_number!="":
            res=innerCheckConfigFilename()
            if res[0] in ["0","9"]:
                return [res[0], res[1]]
            
            elif res[0]=="4":
                innerSendForRepair(res[2])
                return ["2", "Плановый выход из ПП."]
            
            elif res[0]=="2":
                rep_err_list.append(res[2])
                printGREEN("Список с дефектами дополнен")
                
            elif res[0]=="3":
                rep_remark_list.append(res[2]) 
                printGREEN("Список с примечаниями дополнен")

            
            if modeStartMeterTest!="номер только счетчика":
                res=innerCheckStatusSUTP()
                if res[0] in ["0","9"]:
                    return [res[0], res[1]]
                
                elif res[0]=="5":
                    rep_err_list.extend(res[3])
                    clipboard_err_list.extend(res[3])

                elif res[0]=="6":
                    a_err_txt="\n".join(res[3])
                    innerSendForRepair(a_err_txt)
                    return ["2", "Плановый выход из ПП."]
                    
                meter_status_name=res[2]

                if print_number_big_font=="окно":            
                    a_dic={"text_3": "Ok"}
                    saveConfigValue("print_big_font_line.json", a_dic,
                        workmode, "заменить часть")
        

        while True:
            txt1 = "Введите или отсканируйте серийный номер ПУ с его крышки."
            spec_key_list=["/"]
            if meter_tech_number!="":
                txt1=txt1+"\nВведите 0 и Enter, если серийный номер ПУ отсутствует."
                spec_key_list.append("0")
            txt1=txt1+"\nЧтобы прервать проверку ПУ - нажмите '/'."
            len_num_list=[13, 15]

            if print_number_big_font=="окно":
                a_dic={"input_end_list": ["\r"],
                    "input_spec_list": spec_key_list,
                    "text_1": "S/N  METER", 
                    "text_2": "", 
                    "text_3": "",
                    "input_text": ""}
                
                if number_of_meters>1:
                    a_dic["text_1"]=f"pos. {meter_position_cur+1}\n"+ \
                        a_dic["text_1"]

                saveConfigValue("print_big_font_line.json", a_dic,
                    workmode, "заменить часть")
                a_title = 'Печать большим шрифтом'
                actionsSelectedtWindow([a_title], None,"показать+активировать", 
                    "1")
                
            res=inputNumberDevice(len_num_list, "счетчик э/э", "Product1", 
                txt1, spec_key_list, print_number_big_font, workmode)
            
            if res[0]!="1" and  print_number_big_font=="окно":    
                a_title_list = ["Печать большим шрифтом", "otk_print_big_font"]
                actionsSelectedtWindow(a_title_list, None,"закрыть", "1")

            if res[0] in ["0", "3"]:
                return ["0", res[1]]
            
            elif res[0]=="2" and res[1]=="0":
                a_err_list=[]
                err_txt=""
                innerMakeFileName("")
                os.rename(default_filename_old, default_filename_full)
                err_txt="на корпусе отсутствует серийный номер ПУ."
                a_err_list.append(err_txt)
                header=f"\n{bcolors.FAIL}На корпусе ПУ отсутствует серийный номер." \
                        f"{bcolors.ENDC}"
                res=innerSelectActions(header, a_err_list, "3", "сн ПУ", [], [])
                return ["2", "Плановый выход из ПП."]

            elif res[0]=="2" and res[1]=="/":
                return ["9", "Проверка ПУ прервана."]

            if print_number_big_font=="окно":            
                a_dic={"text_3": "Ok"}
                saveConfigValue("print_big_font_line.json", a_dic,
                    workmode, "заменить часть")
                
            meter_sn_lbl=res[1]
            meter_serial_number=meter_sn_lbl
            meter_sn_source="надпись"

            if meter_config_check[0]=="3" and number_of_meters>1 and \
                meter_position_cur>0:
                a_mask_file_name=None
                a_sn_list=listCopy(meter_serial_number_list)
                a_sn_list.append(meter_serial_number)
                for a_sn in a_sn_list:
                    if a_sn=="":
                        continue

                    res=toGetProductInfo2(a_sn, "Product1")
                    if res[0]=="0":
                        a_err_txt = f"Ошибка при получении информации о ПУ № {a_sn} из " \
                            "ф.ProductNumber.xlsx."
                        printFAIL(a_err_txt)
                        return ["0", a_err_txt]
                
                    a_name=res[28]
                    if a_name==None or a_name=="":
                        a_err_txt="В ф.ProductNumber.xlsx отсутствует информация о " \
                            "файле для настройки программы проверки конфигурации\n" \
                            f' ПУ № {a_sn} "MassProdAutoConfig.exe".'
                        printFAIL(a_err_txt)
                        return ["0", a_err_txt]
                    
                    if a_mask_file_name!=None and a_name!=a_mask_file_name:
                        a_err_txt=f"ПУ № {a_sn} нельзя будет проверить в одно время с " \
                            "предыдущими ПУ с помощью программы 'MassProdAutoConfig.exe',\n" \
                            "т.к. у них отличаются имена mask-файлов."
                        printWARNING(a_err_txt)
                        return ["0", a_err_txt]
                    
                    elif a_mask_file_name==None:
                        a_mask_file_name=a_name

            break
        
        
        sheet_name="Product1"
        res=toGetProductInfo2(meter_serial_number, sheet_name, workmode)
        if res[0]!='1' or (res[21] in ["", None, "None"]) or \
            (res[27] in ["", None, "None"]) or \
            (res[19] in ["", None, "None"]) or \
            (res[20] in ["", None, "None"]) or \
            (modeStartMeterTest=="конфигурация" and 
             res[31] in ["", None, "None"]):
            txt1=f"{bcolors.WARNING}При получении информации о ПУ из " \
                f'ф."ProductNumber.xlsx" возникла ошибка.{bcolors.ENDC}\n'
            if res[21] in ["", None, "None"]:
                txt1=txt1+f"{bcolors.WARNING}В таблице не указан способ " \
                    "проверки ПУ.{bcolors.ENDC}\n"
            
            if res[27] in ["", None, "None"]:
                txt1 = txt1+f"{bcolors.WARNING}В таблице не указано " \
                    f"наличие/отсутствие МС.{bcolors.ENDC}\n"

            if res[19] in ["", None, "None"]:
                txt1 = txt1+f"{bcolors.WARNING}В таблице не указана " \
                    f"возможность считывания серийного номера МС.{bcolors.ENDC}\n"

            if res[20] in ["", None, "None"]:
                txt1 = txt1+f"{bcolors.WARNING}В таблице не указана " \
                    f"возможность считывания серийного номера ПДУ.{bcolors.ENDC}\n"
            
            if modeStartMeterTest=="конфигурация" and res[31] in ["", None, "None"]:
                txt1 = txt1+f"{bcolors.WARNING}В таблице не указан " \
                    f"вид применяемого интерфейса для проверки только " \
                    f"конфигурации ПУ.{bcolors.ENDC}\n"

            txt1=txt1+f"{bcolors.WARNING}Продолжение проверки невозможно.{bcolors.ENDC}\n" \
                f"{bcolors.OKBLUE}Нажмите 'Enter'.{bcolors.ENDC}"
            spec_keys=["\r"]
            oo=inputSpecifiedKey("", txt1, "", [0], spec_keys, 1)
            return ["0", "Ошибка в ф.ProductNumber.xlsx."]

        mc_num_set=res[19]
        rc_num_set=res[20]
        verification_method=res[21]
        mc_on_board_set=res[27]
        interface_type_check_config=res[31]

        if modeStartMeterTest=="номер только счетчика":
            a_dic={"оптопорт": "com_opto", "RS-485": "com_rs485"}
            com_current=a_dic.get(interface_type_check_config, "com_opto")

            default_value_dict = writeDefaultValue(default_value_dict)
            saveConfigValue(file_name_in="opto_run.json", 
                var_config_dict=default_value_dict) 
        
        if modeStartMeterTest!="номер только счетчика":    
            if modem_status=="0" and mc_on_board_set=="обязательно":

                innerMakeFileName(meter_serial_number)
                os.rename(default_filename_old, default_filename_full)
                default_filename_old=default_filename_full

                a_txt = f'{bcolors.WARNING}По умолчанию указан статус модема ' \
                    f'"не будет устанавливаться",\n' \
                    f'при этом у данного ПУ ' \
                    f'должен быть установлен МС.{bcolors.ENDC}'
                menu_item_add_list=[]
                menu_id_add_list=[]
                if actions_no_mc=="1":
                    menu_item_add_list=['Изменить статус модема на "будет рабочий"',
                        'Изменить статус модема на "будет тестовый"']
                    menu_id_add_list=['1', '2']
                a_err_txt="У ПУ отсутствует обязательный к установке МС."
                res=innerSelectActions(a_txt, [a_err_txt], "3", "", menu_item_add_list,
                    menu_id_add_list, rep_err_list)
                if res[0]=="2":
                    return ["2", "Плановый выход из ПП."]
                
                elif res[0]=="9":
                    return ["9", "Проверка ПУ прервана."]

                elif res[0]=="4": 
                    modem_status = res[3]
                    a_list = ["устанавливаться не будет", "будет рабочим", 
                            "будет тестовым"]
                    print (f'{bcolors.OKGREEN}Статус МС изменен на ' \
                        f'"{a_list[int(modem_status)]}".{bcolors.ENDC}')


            elif  modem_status in ["1", "2"] and mc_on_board_set=="нет":
                a_modem_status_dic = {"0":"устанавливаться не будет", 
                    "1": "рабочий", "2": "тестовый"}
                a_txt = f'{bcolors.WARNING}По умолчанию указан статус модема ' \
                    f'"{a_modem_status_dic[modem_status]}", при этом у данного ПУ ' \
                    f'МС должен отсутствовать.{bcolors.ENDC}'
                menu_item_add_list=['Изменить статус модема на ' 
                    '"устанавливаться не будет"']
                menu_id_add_list=['0']

                res=innerSelectActions(a_txt, [], "4", "", menu_item_add_list,
                    menu_id_add_list)
                if res[0]=="9":
                    return ["9", "Проверка ПУ прервана."]
                
                elif res[0]=="4": 
                    modem_status = res[3]
                    a_list = ["устанавливаться не будет", "будет рабочим", 
                            "будет тестовым"]
                    print (f'{bcolors.OKGREEN}Статус МС изменен на ' \
                        f'"{a_list[int(modem_status)]}".{bcolors.ENDC}')
            
            
            if modem_status in ["1", "2"] and mc_num_set=="да":
                cicl=True
                while cicl:
                    txt1 = "Введите или отсканируйте номер модуля связи, " \
                            "указанный на корпусе." \
                        "\nЧтобы прервать проверку ПУ - нажмите '/'."
                    spec_key_list=["/"]
                    len_num_list=[13, 15]

                    if print_number_big_font=="окно":
                        a_dic={"input_end_list": ["\r"],
                            "input_spec_list": spec_key_list,
                            "input_max_numb_list": len_num_list,
                            "text_1": "S/N  MC", 
                            "text_2": "", 
                            "text_3": "",
                            "input_text": ""}   

                        if number_of_meters>1:
                            a_dic["text_1"]=f"pos. {meter_position_cur+1}\n"+ \
                                a_dic["text_1"]

                        saveConfigValue("print_big_font_line.json", a_dic,
                            workmode, "заменить часть")
                        
                    res=inputNumberDevice(len_num_list, "модуль связи", "Product1", 
                        txt1, spec_key_list, print_number_big_font, workmode)
                    
                    if res[0]!="1" and  print_number_big_font=="окно":    
                        a_title_list = ["Печать большим шрифтом", "otk_print_big_font"]
                        actionsSelectedtWindow(a_title_list, None,"закрыть", "1")

                    if res[0] in ["0", "3"]:
                        return ["0", "Ошибка."]
                    
                    elif res[0]=="2":
                        return ["9", "Проверка ПУ прервана."]

                    if print_number_big_font=="окно":            
                        a_dic={"text_3": "Ok"}
                        saveConfigValue("print_big_font_line.json", a_dic,
                            workmode, "заменить часть")
                        
                    gsm_serial_number=res[1]
                    mc_on_board="1"
                    break


                
                gsm_SIM_number="0" #номер SIM-карты не нужен
                if SIMcard_status=="1":
                    txt1 = "Введите или отсканируйте номер SIM-карты.\n" \
                            "Если SIM-карты не будет, то нажмите '0' и Enter.\n" \
                            "Если SIM-карта тестовая, то нажмите '+' и Enter.\n" \
                            "Если номер SIM-карты не нужен, то нажмите '-' и Enter.\n"
                    txt2 = "Номер SIM-карты должен состоять из 21 символа."
                    oo = inputSpecifiedKey(bcolors.OKBLUE, txt1, txt2, [21], ["0","+","-"])
                    if oo=="0" or oo=="+":
                        SIMcard_status ="0"
                        if oo=="+":
                            SIMcard_status ="2"
                        default_value_dict = writeDefaultValue(default_value_dict)
                        saveConfigValue('opto_run.json',default_value_dict)
                        oo="0"
                    else:
                        oo="0"
                    gsm_SIM_number=oo
            
            if rc_num_set == "да":
                while True:
                    txt1 = "Введите или отсканируйте серийный номер пульта " \
                        "управления ПУ, указанный на корпусе." \
                        "\nЧтобы вернуться в меню нажмите '/'."
                    spec_key_list = ["/"]
                    len_num_list = [13, 15]

                    if print_number_big_font=="окно":
                        a_dic={"input_end_list": ["\r"],
                            "input_spec_list": spec_key_list,
                            "text_1": "S/N  RC", 
                            "text_2": "", 
                            "text_3": "",
                            "input_text": ""}

                        if number_of_meters>1:
                            a_dic["text_1"]=f"pos. {meter_position_cur+1}\n"+ \
                                a_dic["text_1"]

                        saveConfigValue("print_big_font_line.json", a_dic,
                            workmode, "заменить часть")
                        
                    res=inputNumberDevice(len_num_list, "пульт управления ПУ", "Product1", 
                        txt1, spec_key_list, print_number_big_font, workmode)
                    
                    if res[0]!="1" and  print_number_big_font=="окно":    
                        a_title_list = ["Печать большим шрифтом", "otk_print_big_font"]
                        actionsSelectedtWindow(a_title_list, None,"закрыть", "1")

                    if res[0] in ["0", "3"]:
                        return ["0", "Ошибка."]
                    
                    elif res[0]=="2":
                        return ["9", "Проверка ПУ прервана."]
                    
                    if print_number_big_font=="окно":            
                        a_dic={"text_3": "Ok"}
                        saveConfigValue("print_big_font_line.json", a_dic,
                            workmode, "заменить часть")
                        
                    rc_serial_number = res[1]
                    break

            
        if print_number_big_font=="окно":    
            a_title_list = ["Печать большим шрифтом", "otk_print_big_font"]
            actionsSelectedtWindow(a_title_list, None,"закрыть", "1")

        res=cryptStringSec("зашифровать",meter_pw_default)
        meter_pw_high_encrypt=res[2]
        meter_pw_high_descript=meter_pw_default_descript

        res=cryptStringSec("зашифровать","1234567898765432")
        meter_pw_config_standart_encrypt=res[2]

        res=cryptStringSec("зашифровать","654321")
        meter_pw_low_encrypt=res[2]
        meter_pw_reader_standart_encrypt=meter_pw_low_encrypt
        meter_pw_low_descript="Стандартный нижнего уровня"

        
        if data_exchange_sutp=="0":
            a_txt="Обмен данными с СУТП отключен."
            rep_remark_list.append(a_txt)
        
        
        if data_exchange_sutp!="0":
            device_number=meter_tech_number
            if meter_tech_number=="":
                if meter_sn_lbl!="":
                    device_number=meter_sn_lbl
                else:
                    print (f"{bcolors.WARNING}Отсутствуют технический и серийный номера ПУ." 
                        f"{bcolors.ENDC}")
                    return ["2", "Плановый выход из ПП."]
            
            a_mode="0"
            print ("Запрашиваю из БД СУТП информацию о ПУ...")
            res=getInfoAboutDevice(device_number, workmode, employee_id, 
                employee_pw_encrypt, a_mode)
            if res[0]=="0":
                print(f"\n{bcolors.WARNING}Не удалось получить информацию о ПУ из СУТП.{bcolors.ENDC}")
            elif res[0] in ["1", "2"]:
                meter_tn_sutp=str(res[2])
                meter_sn_sutp=res[3]
                order_num=str(res[8])
                order_descript=res[9]
                order_ev=res[21]
                order_pw_dic=res[20]
                order_pw_reader_encrypt=order_pw_dic.get("pw_reader_encrypt", "")
                order_pw_config_encrypt=order_pw_dic.get("pw_config_encrypt", "")
                meter_pw_high_encrypt=res[11]
                meter_pw_low_encrypt=res[18]
                meter_pw_high_descript=f"Пароль верхнего уровня из БД СУТП для ПУ № {meter_tn_sutp}"
                a_txt="Используем пароль верхнего уровня из БД СУТП."
                a_color=bcolors.OKGREEN
                if meter_pw_high_encrypt=="":
                    a_msg=f"{bcolors.WARNING}Не удалось получить пароль верхнего уровня из СУТП.\n" \
                        f"Для дальнейшей проверки ПУ можно попытаться подобрать пароль из списка " \
                        f"по умолчанию.{bcolors.ENDC}"
                    res=innerSelectActions(a_msg,[],"4","прочие","")
                    if res[0]=="9":
                        return ["9", "Проверка ПУ прервана."]
                    a_msg="Не удалось получить пароль верхнего уровня из СУТП."
                    rep_remark_list.append(a_msg)
                    a_txt=f"Используем пароль верхнего уровня по умолчанию '{meter_pw_default_descript}'."
                    meter_pw_high_descript=meter_pw_default_descript
                    a_color=bcolors.WARNING
                    res=cryptStringSec("зашифровать",meter_pw_default)
                    meter_pw_high_encrypt=res[2]
                print (f'{a_color}{a_txt}{bcolors.ENDC}')

                meter_pw_low_descript=f"Пароль нижнего уровня из БД СУТП для ПУ № {meter_tn_sutp}"
                a_txt="Используем пароль нижнего уровня из БД СУТП."
                a_color=bcolors.OKGREEN
                if meter_pw_low_encrypt=="":
                    a_msg=f"{bcolors.WARNING}Не удалось получить пароль нижнего уровня из СУТП.\n" \
                        f"Для дальнейшей проверки ПУ можно попытаться подобрать пароль из списка " \
                        f"по умолчанию.{bcolors.ENDC}"
                    res=innerSelectActions(a_msg,[],"4","прочие","")
                    if res[0]=="9":
                        return ["9", "Проверка ПУ прервана."]
                    a_msg="Не удалось получить пароль нижнего уровня из СУТП."
                    rep_remark_list.append(a_msg)
                    a_txt=f"Используем пароль нижнего уровня по умолчанию."
                    meter_pw_low_descript="Стандартный нижнего уровня"
                    a_color=bcolors.WARNING
                    res=cryptStringSec("зашифровать","654321")
                    meter_pw_low_encrypt=res[2]
                print (f'{a_color}{a_txt}{bcolors.ENDC}')

        
        if meter_tech_number=="" and meter_tn_sutp!="":
            print (f"{bcolors.OKGREEN}Из СУТП получен технический номер: "
                f"{meter_tn_sutp}{bcolors.ENDC}")
            meter_tech_number=meter_tn_sutp
            meter_tn_source="СУТП"

        if meter_sn_sutp!="":
            print (f"{bcolors.OKGREEN}Из СУТП получен серийный номер: "
                f"{meter_sn_sutp}{bcolors.ENDC}")
            meter_serial_number=meter_sn_sutp
            meter_sn_source="СУТП"

        
        innerMakeFileName(meter_serial_number)
        os.rename(default_filename_old, default_filename_full)


        if order_num!="" and order_num!="None" and order_control=="1":
            a_descript=f"№ {order_num}: {order_descript}"
            if order_control_descript=="":
                order_control_descript=a_descript
            a_descript1=a_descript.replace(" ","").upper()
            a_order=order_control_descript.replace(" ","").upper()
            if a_order!=a_descript1:
                txt1_1=f"{bcolors.WARNING}ПУ № {meter_tech_number} отсутствует в " \
                    f"контролируемом заказе '{order_control_descript}'.{bcolors.ENDC}\n" \
                    f"{bcolors.WARNING}Он включен в заказ '{a_descript}'.{bcolors.ENDC}"
                menu_item_list=["Заменить контролируемый заказ на заказ проверяемого ПУ",
                    "Продолжить проверку"]
                menu_id_list=["заменить заказ", "продолжить проверку"]
                if order_num=="None":
                    txt1_1=f"{bcolors.WARNING}На ПУ № {meter_tech_number} отсутствует " \
                        f"заказ.{bcolors.ENDC}"
                    menu_item_list=["Продолжить проверку"]
                    menu_id_list=["продолжить проверку"]
                print(txt1_1)
                txt1=f"Выберите дальнейшее действие"
                spec_list=["Прервать проверку"]
                spec_keys=["/"]
                spec_id_list=["прервать"]
                oo=questionFromList(bcolors.OKBLUE, txt1, menu_item_list, menu_id_list,
                    "", spec_list, spec_keys, spec_id_list, start_list_num=1)
                print()
                if oo=="прервать":
                    return ["9", "Проверка ПУ прервана."]
                
                elif oo=="заменить заказ":
                    order_control_descript=a_descript
                    default_value_dict = writeDefaultValue(default_value_dict)
                    saveConfigValue('opto_run.json',default_value_dict)
                    print(f"{bcolors.OKGREEN}Установлен новый заказ для контроля: " 
                        f"'{order_control_descript}'.{bcolors.ENDC}")


        print()
        a_dic={0:"meter_pass_low.json", 1:"meter_pass.json"}
        a_txt_dic={0:"нижнего", 1:"верхнего"}

        pw_type_config_id=order_pw_dic["pw_type_config_id"]
        pw_type_reader_id=order_pw_dic["pw_type_reader_id"]
        a_pw_config_dic={0: meter_pw_config_standart_encrypt, 
                1: order_pw_dic["pw_config_encrypt"],
                2: meter_pw_high_encrypt}
        a_pw_reader_dic={0: meter_pw_reader_standart_encrypt, 
                1: order_pw_dic["pw_reader_encrypt"],
                2: meter_pw_low_encrypt}
        order_pw_low_encrypt=a_pw_reader_dic.get(pw_type_reader_id,"")
        order_pw_high_encrypt=a_pw_config_dic.get(pw_type_config_id,"")
        order_pw_type_low_descript=order_pw_dic.get("pw_type_reader_descript","")
        order_pw_type_high_descript=order_pw_dic.get("pw_type_config_descript","")
        
        a_pw_dic={0: ["Low", meter_pw_low_encrypt, meter_pw_low_descript,
            order_pw_low_encrypt, order_pw_type_low_descript],
            1:["High", meter_pw_high_encrypt, meter_pw_high_descript,
            order_pw_high_encrypt, order_pw_type_high_descript]}
        a_err_list=[]

        for step in range(2):
            print(f"{bcolors.OKGREEN}Проверяем пароль {a_txt_dic[step]} уровня.{bcolors.ENDC}")
            a_name=a_dic.get(step,"")
            meter_pw_level=a_pw_dic.get(step,"")[0]
            meter_pw_encrypt=a_pw_dic.get(step,"")[1]
            meter_pw_descript=a_pw_dic.get(step,"")[2]
            order_pw_encrypt=a_pw_dic.get(step,"")[3]
            order_pw_type_descript=a_pw_dic.get(step,"")[4]

            if meter_pw_visible=="1":
                a_meter=None
                a_order=None
                res=cryptStringSec("расшифровать", meter_pw_encrypt)
                if res[0]=="1":
                    a_meter=res[2]
                res=cryptStringSec("расшифровать", order_pw_encrypt)
                if res[0]=="1":
                    a_order=res[2]
                print(f"Пароль {a_txt_dic[step]} уровня из БД СУТП: " \
                        f"{a_meter}")
                print(f"Пароль {a_txt_dic[step]} уровня из заказа: " \
                        f"{a_order}")

            if  order_num!="None" and order_num!=None and order_pw_encrypt!="":
                a_meter=cryptStringSec("расшифровать", meter_pw_encrypt)[2]
                a_order=cryptStringSec("расшифровать", order_pw_encrypt)[2]            
                if a_meter!=a_order:
                    a_err_txt=f"Пароль {a_txt_dic[step]} уровня в БД СУТП отличается от пароля, " \
                            f"указанного в заказе: {order_pw_type_descript}."
                    print(f"{bcolors.FAIL}{a_err_txt}{bcolors.ENDC}")
                    a_err_list.append(a_err_txt)

            meter_pw_descript_old=meter_pw_descript
            res=checkPW(a_name, order_pw_type_descript)
            ans=res[0]
            ans_pw=res[2]
            ans_pw_descript=res[3]
            if ans=="5":
                if meter_pw_level=="High":
                    meter_pw_high_encrypt=ans_pw
                    meter_pw_high_descript=ans_pw_descript
                else:
                    meter_pw_low_encrypt=ans_pw
                    meter_pw_low_descript=ans_pw_descript
                default_value_dict = writeDefaultValue(default_value_dict)
                saveConfigValue('opto_run.json',default_value_dict)
                
            elif ans in ["4", "6"]:
                print(f"{bcolors.FAIL}При проверке пароля подключения к ПУ " \
                    f"{a_txt_dic[step]} уровня возникла ошибка.{bcolors.ENDC}")
                print(f"{bcolors.FAIL}Продолжение проверки ПУ невозможно.{bcolors.ENDC}")
                return ["0", "Ошибка."]
            
            elif ans=="0":
                a_txt1="Нет связи с ПУ через оптопорт."
                txt1=f"{bcolors.FAIL}{a_txt1}{bcolors.ENDC}"
                a_err_list=["нет связи с ПУ через оптопорт"]
                res=innerSelectActions(txt1, a_err_list, "3")
                return ["0", "Ошибка."]

            if (ans in ["2", "3", "5"]) and ("ИЗ БД СУТП" in meter_pw_descript_old.upper()) \
                or (ans=="7"):
                a1_dic={
                    "2": f"Пароль {a_txt_dic[step]} уровня в ПУ отличается от пароля, " \
                        f"указанного в БД СУТП. Не удалось подобрать пароль.",
                    "3": f"Пароль {a_txt_dic[step]} уровня в ПУ отличается от пароля, " \
                        f"указанного в БД СУТП. Подобран пароль '{ans_pw_descript}'.",
                    "5": f"Пароль {a_txt_dic[step]} уровня в ПУ отличается от пароля, " \
                        f"указанного в БД СУТП. Подобран пароль '{ans_pw_descript}'.",
                    "7": "Сгенерированный пароль верхнего уровня совпадает со " \
                        "стандартным значением."}
                a_err_list.append(a1_dic[ans])
            
            if step==1 and (ans in ["2", "3"]):
                a_dic={"2": "Пароль верхнего уровня отличается от пароля в БД СУТП.",
                    "3": "Пароль верхнего уровня отличается от пароля в БД СУТП. " \
                        f"Подобран пароль '{ans_pw_descript}'.",
                    "7": "Сгенерированный пароль верхнего уровня совпадает со " \
                        "стандартным значением."}
                err_txt=a_dic[ans]
                print(f"{bcolors.WARNING}{err_txt}{bcolors.ENDC}\n"
                    f"{bcolors.WARNING}Продолжение проверки ПУ невозможно.{bcolors.ENDC}")
                if len(a_err_list)==0:
                    a_err_list.append(err_txt)
                err_txt="\n".join(a_err_list)
                header=(f"{bcolors.WARNING}При проверке паролей доступа были выявлены " \
                        f"следующие ошибки:\n{err_txt}{bcolors.ENDC}")
                res=innerSelectActions(header, a_err_list, "3")
                return ["2", "Плановый выход из ПП."]
                
        
        if len(a_err_list)>0:
            err_txt="\n".join(a_err_list)
            print (f"{bcolors.WARNING}При проверке паролей доступа были выявлены " \
                f"замечания:\n{err_txt}{bcolors.ENDC}")
            res=innerSelectActions("", a_err_list, "1")
            if res[0]=="1":
                rep_err_list.extend(res[2])
                clipboard_err_list.extend(res[2])
            
            elif res[0]=="9":
                return ["9", "Проверку ПУ прервали."]
            
            else:
                return ["2", "Плановый выход из ПП."]
            

        default_value_dict = writeDefaultValue(default_value_dict)
        saveConfigValue('opto_run.json',default_value_dict)

 
        a_txt="Получение серийного номера, мгновенных значений " \
            "напряжения/тока и версии ПО из ПУ."
        print(f"{bcolors.OKGREEN}{a_txt}{bcolors.ENDC}")
        while True:
            res=innerReadMeter("2")
            if res!="1":
                return ["0", "Ошибка."]
            
            
            a_v_list=list(meter_voltage_dic.values())
            a_phase_volt=0
            a_v_str="/".join(a_v_list)
            if meter_phase=="1":
                a_v_str=a_v_list[0]
            for a_v in a_v_list:
                if int(a_v)>50:
                    a_phase_volt+=1
            
            a_amp_list=list(meter_amperage_dic.values())
            a_phase_amp=0
            a_amp_str="/".join(a_amp_list)
            if meter_phase=="1":
                a_amp_str=a_amp_list[0]
            for a_amp in a_amp_list:
                if int(a_amp)>ctrl_current_electr_test*100:
                    a_phase_amp+=1

            a_schem=f"{a_phase_volt}-{a_phase_amp}"

            a_dic={"1-0":"напряжение - 1 фаза, ток - нет", 
                    "1-1":"напряжение - 1 фаза, ток - 1 фаза",
                    "2-0":"напряжение - 2 фазы, ток - нет",
                    "2-1":"напряжение - 2 фазы, ток - 1 фаза",
                    "2-2":"напряжение - 2 фазы, ток - 2 фазы",
                    "3-0":"напряжение - 3 фазы, ток - нет",
                    "3-1":"напряжение - 3 фазы, ток - 1 фаза",
                    "3-2":"напряжение - 3 фазы, ток - 2 фазы",
                    "3-3":"напряжение - 3 фазы, ток - 3 фазы"}
            
            if not electrical_test_circuit in a_dic:
                electrical_test_circuit="1-0"
                default_value_dict = writeDefaultValue(default_value_dict)
                saveConfigValue('opto_run.json',default_value_dict)

            a_schem_def_descript=a_dic.get(electrical_test_circuit, "не указана")
            a_schem_cur_descript=a_dic.get(a_schem, "неизвестная")

            if electrical_test_circuit!=a_schem:
                txt1=f"По умолчанию указана схема подключения ПУ для " \
                    f"проверки: {a_schem_def_descript}.\n" 
                if electrical_test_circuit in ["1-1", "2-1", "2-2", "3-1",
                    "3-2", "3-2", "3-3"]:
                    txt1=txt1+f"Контрольное значение тока {ctrl_current_electr_test*1000} мА.\n"
                
                txt1=txt1+ f"Согласно измеренным значениям напряжения и тока " \
                    f"({a_v_str} В, {a_amp_str} мА) текущая схема\n" \
                    f"подключения: {a_schem_cur_descript}."
                
                menu_item_list=["Заменить схему подключения по умолчанию на текущую",
                    "Повторно проверить схему подключения ПУ (контакт восстановлен)",
                    "Продолжить проверку"]
                menu_id_list=["заменить схему", "повторить", "продолжить проверку"]
                
                if a_phase_volt==0 or a_schem_cur_descript=="неизвестная":
                    txt1="В измерительных цепях напряжения ПУ отсутствует напряжение.\n" \
                        "Подайте напряжение."
                    if a_schem_cur_descript=="неизвестная":
                        txt1="Неизвестная схема подключения ПУ."
                        
                    menu_item_list=["Повторно проверить схему подключения ПУ", 
                        "Отправить ПУ в ремонт"]
                    menu_id_list=["повторить", "ремонт"]

                printFAIL(txt1)

                txt1=f"Выберите дальнейшее действие:"
                spec_list=["Прервать проверку"]
                spec_keys=["/"]
                spec_id_list=["прервать"]
                oo=questionFromList(bcolors.OKBLUE, txt1, menu_item_list, menu_id_list,
                    "", spec_list, spec_keys, spec_id_list)
                print()
                if oo=="прервать":
                    return ["9", "Проверку ПУ прервали."]
                
                elif oo=="заменить схему":
                    electrical_test_circuit=a_schem
                    default_value_dict = writeDefaultValue(default_value_dict)
                    saveConfigValue('opto_run.json',default_value_dict)
                    print(f"{bcolors.OKGREEN}Изменена схема подключения ПУ для проверки: " 
                            f"'{a_schem_cur_descript}'.{bcolors.ENDC}")
                    break

                elif oo=="продолжить проверку":
                    break

                elif oo=="повторить":
                    print (f"{bcolors.OKGREEN}Повторное получение мгновенных значений " \
                        f"напряжения/тока из ПУ.{bcolors.ENDC}")
                    
                elif oo=="ремонт":
                    header="При проверке схемы подключения ПУ на стенде выявлено замечание."
                    a_err_txt="Неисправны измерительные цепи:\n" \
                        "при схеме подключения ПУ на испытательном стенде " \
                        f"'{a_schem_def_descript}', ПУ сообщает о следующих измеренных" \
                        f"значениях напряжения и тока: {a_v_str} В, {a_amp_str} мА."
                    innerSelectActions(header, [a_err_txt], "9",
                        err_no_edit_list=rep_err_list)
                    return ["2", "Плановый выход из ПП."]
                    
            else:
                break

        
        
        if meter_config_check[0]=="3":
            res=innerCheckPassProdVers(meter_soft, meter_serial_number)
            if res[0]!="1":
                return ["0", res[1]]

        
        if meter_sn_ep=="":
            txt1="В электронном паспорте отсутствует серийный номер ПУ."
            print (f"{bcolors.FAIL}{txt1}{bcolors.ENDC}")
        else:
            a_color=bcolors.OKGREEN
            if int(meter_sn_ep)==0:
                a_color=bcolors.FAIL
        print(f"{a_color}В электронном паспорте записан серийный номер ПУ: "\
            f"{meter_sn_ep}{bcolors.ENDC}")

        
        if meter_sn_source!="СУТП" and int(meter_sn_ep)!=0:
            meter_serial_number=meter_sn_ep
            meter_sn_source="электронный паспорт"

        default_value_dict = writeDefaultValue(default_value_dict)
        saveConfigValue('opto_run.json',default_value_dict)



        if verification_method=="сплит":
            txt1=f"{bcolors.OKGREEN}Проверка наличия связи с ПУ через RS-485.{bcolors.ENDC}"
            print(txt1)
            a_dic={"meter_sn_ep":""}
            saveConfigValue('opto_run.json',a_dic)

            com_current="com_rs485"
            
            res=checkComPort(com_opto=com_rs485, print_msg="")
            if res!="1":
                res=restoreCOMPort(com_current)    
                if res=="3":
                    return ["9", "Проверку ПУ прервали."]
                
                elif res!="1":
                    return ["0", "Ошибка."]

            default_value_dict = writeDefaultValue(default_value_dict)
            saveConfigValue('opto_run.json',default_value_dict)

            res_rs485=innerReadMeter("1")


            if res_rs485!="1":
                txt1=f"{bcolors.WARNING}Не удалось установить связь " \
                    f"с ПУ через RS-485.{bcolors.ENDC}"
                a_err_list=["Нет связи с ПУ через RS-485"]
                a_err_list=rep_err_list+a_err_list
                res=innerSelectActions(txt1, a_err_list, "1")
                if res[0]=="1":
                    rep_err_list.extend(a_err_list)
                    clipboard_err_list.extend(a_err_list)
                    verification_method="станд"
                
                elif res[0]=="9":
                    return [res[0], res[1]]
                
                else:
                    return ["2", "Плановый выход из ПП."]
                
            else:
                printWARNING ("\nПроверка связи с ПУ через оптопорт " 
                    "закончена.\nОптопорт можно снять.")
                printGREEN ("Проверка связи с ПУ через RS-485 "
                    "пройдена успешно.")
                
        
        
        innerMakeFileName(meter_serial_number)
        os.rename(default_filename_old, default_filename_full)

        
        device_repay_history=""
        if data_exchange_sutp!="0" and meter_tech_number!="":
            res=getDeviceRepayHistory(device_number=meter_tech_number, workmode=workmode)
            if res[0]=="1":
                device_repay_history=res[2]
                if device_repay_history!="":
                    txt=f"\n{bcolors.WARNING}Список ранее выявленных несоответствий:{bcolors.ENDC}\n" \
                        f"{bcolors.WARNING}{device_repay_history}{bcolors.ENDC}"
                    print(txt)
                else:
                    print (f"{bcolors.OKGREEN}Информации о ремонте ПУ нет.{bcolors.ENDC}")


        if res_ext_at_begin_test=="1":
            print()
            a_err_txt="\n".join(rep_err_list)
            txt1=f"{bcolors.OKGREEN}Запись результатов внешнего осмотра.{bcolors.ENDC}"
            if len(rep_err_list)>0:
                print(f"{txt1}\n" \
                    f"{bcolors.WARNING}Ранее были внесены следующие замечания:{bcolors.ENDC}\n" \
                    f"{bcolors.WARNING}{a_err_txt}{bcolors.ENDC}")
            res=innerSelectActions(txt1, rep_err_list, "7")
            if res[0]=="2":
                return ["2", "Плановый выход из ПП."]
            
            elif res[0]=="9":
                return [res[0], res[1]]
            
            elif res[0]=="3" and len(res[2])>0 :
                rep_err_list=res[2].copy()
                clipboard_err_list=rep_err_list.copy()


        txt1 = f"1. Дата и время проведения проверки: {now_full}\n" \
            f"   Статус ПУ в СУТП: {meter_status_name}\n" \
            f"   Описание пароля доступа к ПУ нижнего уровня: {meter_pw_low_descript}\n" \
            f"   Описание пароля доступа к ПУ верхнего уровня: {meter_pw_high_descript}\n"
        a_txt=f"   Заказ: № {order_num} '{order_descript}'."
        if order_ev=="1":
            a_txt=f"{a_txt} Заказ относится к розничной продаже (EV)."
        a_txt=f"{a_txt}\n"
        if order_num=="None" or order_num==None :
            a_txt="   Заказ: нет данных\n"
        a_low_txt=order_pw_type_low_descript
        a_high_txt=order_pw_type_high_descript
        if a_low_txt=="":
            a_low_txt= "нет данных"
        if a_high_txt=="":
            a_high_txt= "нет данных"
        a_txt=a_txt+f"   Тип пароля нижнего уровня, указанный в заказе: " \
            f"{a_low_txt}\n" \
            f"   Тип пароля верхнего уровня, указанный в заказе: " \
            f"{a_high_txt}"
        txt1=txt1+a_txt
        print(f"{bcolors.OKGREEN}{txt1}{bcolors.ENDC}")
        fileWriter(default_filename_full, "a", "", f"{txt1}\n", \
            "Сохранение в отчет даты и серийного номера ПУ",join="on")
        

        txt1=f"2. Интерфейсы:\n" \
            f"   Оптопорт: связь установлена."
        if verification_method=="сплит":
            txt1=txt1+f"\n   RS-485: связь установлена."
        else:
            txt1=txt1+f"\n   RS-485: тест связи не проводился."

        print(f"{bcolors.OKGREEN}{txt1}{bcolors.ENDC}")
        fileWriter(default_filename_full, "a", "", f"{txt1}\n", \
            "Сохранение в отчет информации об интерфейсах.",join="on")

        
        a_color_sn_lbl=bcolors.OKGREEN
        a_color_sn_ep=bcolors.OKGREEN
        a_err_txt_lbl=""
        a_err_txt_ep=""
        if meter_sn_lbl!=meter_serial_number:
            a_color_sn_lbl=bcolors.FAIL
            a_err_txt_lbl=f"Серийный номер на крышке ПУ {meter_sn_lbl} отличается " \
                f"от эталонного серийного номера {meter_serial_number} ({meter_sn_source})."
        if meter_sn_ep!=meter_serial_number:
            a_color_sn_ep=bcolors.FAIL
            a_err_txt_ep=f"Серийный номер в электронном паспорте ПУ {meter_sn_ep} отличается " \
                f"от эталонного серийного номера {meter_serial_number} ({meter_sn_source})."


        txt1=f"{bcolors.OKGREEN}3.1. Номера ПУ:\n" \
            f"     Серийный номер в электронном паспорте ПУ:{bcolors.ENDC} " \
                f"{a_color_sn_ep}{meter_sn_ep}{bcolors.ENDC}\n" \
            f"     {bcolors.OKGREEN}Серийный номер, указанный на крышке ПУ (QR-код):{bcolors.ENDC} " \
                f"{a_color_sn_lbl}{meter_sn_lbl}{bcolors.ENDC}\n" \
            f"     {bcolors.OKGREEN}Серийный номер из СУТП: {meter_sn_sutp}\n" \
            f"     Эталонный серийный номер ({meter_sn_source}): {meter_serial_number}\n" \
            f"     Эталонный технический номер ({meter_tn_source}): {meter_tech_number}" \
                f"{bcolors.ENDC}"
        rep_txt1=f"3.1. Номера ПУ:\n" \
            f"     Серийный номер в электронном паспорте ПУ: {meter_sn_ep}\n" \
            f"     Серийный номер, указанный на крышке ПУ (QR-код): {meter_sn_lbl}\n" \
            f"     Серийный номер из СУТП: {meter_sn_sutp}\n" \
            f"     Эталонный серийный номер ({meter_sn_source}): {meter_serial_number}\n" \
            f"     Эталонный технический номер ({meter_tn_source}): {meter_tech_number}"
        a_err_list=[]
        if a_err_txt_lbl!="":
            txt1=f"{txt1}\n     {bcolors.FAIL}{a_err_txt_lbl}{bcolors.ENDC}"
            rep_txt1=f"{rep_txt1}\n     {a_err_txt_lbl}"
            a_err_list.append(a_err_txt_lbl)
        if a_err_txt_ep!="":
            txt1=f"{txt1}\n     {bcolors.FAIL}{a_err_txt_ep}{bcolors.ENDC}"
            rep_txt1=f"{rep_txt1}\n     {a_err_txt_ep}"
            a_err_list.append(a_err_txt_ep)
        print(f"{txt1}")

        if len(a_err_list)>0:
            a_txt="\n".join(a_err_list)
            txt1=(f"{bcolors.WARNING}При проверке серийного номера ПУ были выявлены " \
                f"следующие ошибки:\n{a_txt}{bcolors.ENDC}")
            res=innerSelectActions(txt1, a_err_list, "1", "сн ПУ", [], [],
                rep_err_list)
            if res[0]=="1":
                rep_err_list.extend(a_err_list)
                clipboard_err_list.extend(a_err_list)
            
            elif res[0]=="9":
                return ["9", "Проверку ПУ прервали."]
            
            else:
                return ["2", "Плановый выход из ПП"]

        fileWriter(default_filename_full, "a", "", f"{rep_txt1}\n", \
            "Сохранение в отчет даты и номеров ПУ",join="on")
        
        rep_err_list=delItemList(rep_err_list)
        clipboard_err_list=delItemList(clipboard_err_list)
        rep_remark_list=delItemList(rep_remark_list)

        default_value_dict = writeDefaultValue(default_value_dict)
        saveConfigValue('opto_run.json',default_value_dict)
        
        return ["1", "Номера ПУ получены, пароли проверены."]

    else:
        res=getPathOptoRun(workmode)
        if res[0]=="":
            return ["0", res[1]]
        opto_run_path =res[2]
        multi_config_dir=res[3]

        opto_run_path_last=f"{opto_run_path}_last"
        res=copyFile(opto_run_path, opto_run_path_last, "0", "1")
        if res[0]=="0":
            a_err_txt="Ошибка при дублировании ф. opto_run.json."
            printFAIL(a_err_txt)
            keystrokeEnter()
            return ["0", a_err_txt]

        
        for i in range(0,len(meter_tech_number_list)):
            if meter_tech_number_list[i]!=None and \
                meter_tech_number_list[i]!="":
                a_num=meter_serial_number_list[i]
                break
        
        else:
            a_err_txt="Список технических номеров ПУ для проверки пуст."
            printFAIL(a_err_txt)
            keystrokeEnter()
            return ["0", a_err_txt]
                

        if meter_config_check[0]=="3" :
            res=toGetProductInfo2(a_num, "Product1", workmode)
            if res[0]!='1' or (res[21] in ["", None, "None"]):
                txt1=f"{bcolors.FAIL}При получении информации о ПУ из " \
                    "ф.'ProductNumber.xlsx' возникла ошибка.\n" \
                    "Продолжение проверки невозможно.\n" \
                    f"{bcolors.OKBLUE}Нажмите 'Enter'.{bcolors.ENDC}"
                spec_keys=["\r"]
                oo=inputSpecifiedKey("", txt1, "", [0], spec_keys, 1)
                return ["0", "Ошибка в ф.ProductNumber.xlsx."]
            
            verification_method=res[21]

            if verification_method=="сплит" and com_config_current_select=="1":
                com_config_current="com_config_rs485"

            elif verification_method!="сплит" and com_config_current_select=="1":
                a_dic={"com_opto": "com_config_opto",
                    "com_rs485":"com_config_rs485"}
                com_config_current=a_dic[com_current] 

            default_value_dict = writeDefaultValue(default_value_dict)
            saveConfigValue('opto_run.json',default_value_dict)

            while True:
                _, ans2, file_name = getUserFilePath('getLogMassConfig.py')
                if file_name == "":
                    return ["0", "Ошибка."]
                
                
                printGREEN("Для проверки конфигурации ПУ используем программу " 
                    f"MassProdAutoConfig.exe с версией {mass_prod_vers}.")
                
                a_dic = {"mass_prod_vers": mass_prod_vers}
                saveConfigValue("mass_config.json",a_dic, workmode) 

                txt1=f"Пожалуйста подождите, идет загрузка программы для " \
                    "проверки конфигурации ПУ...\n"
                printGREEN(txt1)

                code_exit = os.system(f"python {file_name} '0'")


                title_new="Аппаратная проверка ПУ"
                res = replaceTitleWindow("", title_new)

                res=readGonfigValue("mass_config.json",[],{}, workmode, "1")
                if res[0]!="1":
                    return ["0", res[1]]
                
                mass_responce=res[2]["mass_responce"]
                
                if mass_responce=="1":
                    res=readGonfigValue("mass_log_line_multi.json", [], {},
                        workmode, "1")
                    if res[0]!="1":
                        printWARNING("Не удалось получить результаты проверки "
                            "конфигурации ПУ.")
                        meter_config_check="23"
                        break 

                    mass_res_multi_dic=res[2]
                    
                    meter_mass_res_dic={}
                    for i in range(0, len(meter_tech_number_list)):
                        a_meter_tech_number=meter_tech_number_list[i]
                        
                        if a_meter_tech_number==None or a_meter_tech_number=="":
                            continue

                        a_save_change=False
                        
                        meter_mass_res_dic=mass_res_multi_dic.get(a_meter_tech_number, {})

                        meter_analisys_res_0=meter_mass_res_dic["analisys_res_0"]
                        
                        meter_config_res_list=meter_mass_res_dic["err_in_log_list"]
                        a_except_list=meter_mass_res_dic["except_in_log_list"]
                        a_sub_no_found_list=meter_mass_res_dic["no_substrings_found_list"]
                        
                        a_name=f"opto_run_{str(i)}.json"
   
                        path_opto_run_multi=os.path.join(multi_config_dir, a_name)
                        
                        res=readWriteFile(path_opto_run_multi, "r-json", "", "utf-8", "1")
                        if res[0]=="0":
                            return ["0","Ошибка при чтении файла конфигурации."]
                        
                        a_def_dic=res[2]

                        a_rep_err_list=a_def_dic["rep_err_list"]
                        a_clipboard_err_list=a_def_dic["clipboard_err_list"]
                        a_rep_remark_list=a_def_dic["rep_remark_list"]
                        a_meter_config_check=a_def_dic["meter_config_check"]
                        
                        if len(meter_config_res_list)>0 or \
                            len(a_except_list)>0 or len(a_sub_no_found_list)>0:
                            a_save_change=True

                            if len(meter_config_res_list)>0:

                                a_txt="Конфигурация ПУ не соответствует заказу:"
                                a_rep_err_list.append(a_txt)
                                a_rep_err_list.extend(meter_config_res_list)

                                a_clipboard_err_list.append(a_txt)
                                a_clipboard_err_list.extend(meter_config_res_list)
                            
                            if len(a_except_list)>0:
                                a_txt="Список параметров ПУ, которые включены в " \
                                    "список исключений для данной версии ПО:"
                                a_rep_remark_list.append(a_txt)
                                a_rep_remark_list.extend(a_except_list)

                            if len(a_sub_no_found_list)>0:
                                a_txt="Список контрольных подстрок (блоков), которые " \
                                    "не были найдены в log-файле:" 
                                a_rep_remark_list.append(a_txt)
                                a_rep_remark_list.extend(a_sub_no_found_list)
                        
                        if meter_analisys_res_0=="":
                            a_meter_config_check="03"

                            a_save_change=True
                                
                        if a_save_change:
                            a_def_dic["meter_config_res_list"]=meter_config_res_list
                            a_def_dic["rep_err_list"]=a_rep_err_list
                            a_def_dic["clipboard_err_list"]=a_clipboard_err_list
                            a_def_dic["rep_remark_list"]=a_rep_remark_list
                            a_def_dic["meter_config_check"]=a_meter_config_check

                            res=readWriteFile(path_opto_run_multi, "w-json", a_def_dic, 
                                "utf-8", "1")
                            if res[0]=="0":
                                return ["0","Ошибка при записи файла конфигурации."]
                        
                    break
                    
                elif mass_responce in ["0", "2"]:
                    a_txt=f"{bcolors.OKBLUE}Выберите дальнейшее действие:"
                    if mass_responce=="2":
                        a_txt=f"{bcolors.WARNING}Не удалось провести проверку " \
                            "конфигурации ПУ в автоматическом режиме.\n" \
                            "Проверьте:\n" \
                            "1) 'Caps Lock' должен быть отключен.\n" \
                            "2) Должна быть включена " \
                            f"{bcolors.ATTENTIONWARNING} АНГЛИЙСКАЯ (EN) "\
                            f"{bcolors.ENDC} {bcolors.WARNING}раскладка " \
                            f"клавиатуры.\n{a_txt}"
                    menu_item_list=["Повторить попытку проверки конфигурации ПУ "
                        "в автоматическом режиме", "Провести проверку конфигурации "
                        "ПУ в ручном режиме", "Пропустить проверку конфигурации ПУ"]
                    menu_id_list=["повторить", "ручная проверка", "пропустить"]
                    
                    a_meters_fact=0
                    for i in range(0, len(meter_tech_number_list)):
                        a_meter_tech_number=meter_tech_number_list[i]
                        if a_meter_tech_number!=None and a_meter_tech_number!="":
                            a_meters_fact+=1
                    
                    if a_meters_fact>1:
                        menu_item_list=["Повторить попытку проверки конфигурации ПУ "
                        "в автоматическом режиме", "Провести проверку конфигурации "
                        "ПУ в ручном режиме"]
                        menu_id_list=["повторить", "ручная проверка"]

                    oo=questionFromList(bcolors.OKBLUE, a_txt, menu_item_list, 
                        menu_id_list, "", ["Прервать проверку ПУ"], ["/"],
                        ["прервать"], 1, 1,1)
                    print()
                    if oo=="прервать":
                        return ["9", "Проверка ПУ прервана."]
                    
                    elif oo=="повторить":
                        continue

                    elif oo=="ручная проверка":
                        meter_config_check="23"
                        break

                    elif oo=="пропустить":
                        meter_config_check = "03"
                        break
                
                elif mass_responce=="8":
                    meter_config_check="03"
                    break
                
                elif mass_responce=="9":
                    return ["9", "Проверка ПУ прервана."]
                
                elif mass_responce in ["3", "4"]:
                    meter_config_check="23"
                    break

                else:
                    printWARNING("Неизвестный код ответа от программы "
                        f"'{mass_responce}'.")
                    return ["0", "Ошибка."]

        
        default_value_dict = writeDefaultValue(default_value_dict)
        saveConfigValue(file_name_in="opto_run.json", 
            var_config_dict=default_value_dict)
        
        res=copyFile(opto_run_path, opto_run_path_last, "0", "1")
        if res[0]=="0":
            a_err_txt="Ошибка при копировании ф. opto_run.json."
            printFAIL(a_err_txt)
            keystrokeEnter()
            return ["0", a_err_txt]
        
        
        meter_tech_number_list_last=meter_tech_number_list.copy()
        meter_status_test_list_last=meter_status_test_list.copy()
        meter_serial_number_list_last=meter_serial_number_list.copy()
        
        if  meter_config_check[0]=="2":
            menu_item_list=[]
            menu_id_list=[]
            meter_pos_config_err_list=[]
            
            for i in range(0, len(meter_tech_number_list)):
                a_meter_tech_number=meter_tech_number_list[i]
                if a_meter_tech_number==None or a_meter_tech_number=="":
                    continue
                
                menu_item_list.append(f"поз.{i+1}: {a_meter_tech_number}")
                
                menu_id_list.append(str(i))

                meter_pos_config_err_list.append(str(i))

            
            spec_list=["у всех ПУ в списке положительный результат",
                "у всех ПУ в списке отрицательный результат", 
                "прервать проверку всех ПУ"]
            spec_keys=["+", "-", "/"]
            spec_id = ["положительно", "отрицательно",
                "прервать"]
            
            while len(menu_item_list)>0:
                a_cur=0
                while a_cur<len(menu_item_list):
                    if menu_item_list[a_cur]=="":
                        del menu_item_list[a_cur]
                        del menu_id_list[a_cur]
                        a_cur-=1
                    
                    a_cur+=1

                if len(menu_item_list)==0:
                    break

                if len(menu_item_list)==1:
                    spec_list=["у всех ПУ в списке положительный результат",
                    "прервать проверку всех ПУ"]
                    spec_keys=["+", "/"]
                    spec_id = ["положительно", "прервать"]

                txt1="Выберите ПУ для внесения результата проверки конфигурации:"
                
                oo="0"

                if number_of_meters>1:
                    oo=questionFromList(bcolors.OKBLUE, txt1, menu_item_list,
                        menu_id_list, "", spec_list, spec_keys, spec_id, 1, 1, 1,
                        [])
                    print()
                
                if oo=="прервать":
                    return ["9", "Проверка ПУ прервана."]
                
                elif oo=="положительно":
                    meter_pos_config_err_list=[]

                    break

                elif oo=="отрицательно":
                    pass
                
                else:
                    meter_pos_config_err_list=[oo]

                
                for i in range(0, len(meter_pos_config_err_list)):
                    meter_pos=meter_pos_config_err_list[i]

                    a_name=f"opto_run_{meter_pos}.json"

                    opto_run_multi_path=os.path.join(multi_config_dir, a_name)
                    
                    res=copyFile(opto_run_multi_path, opto_run_path, "0", "1")
                    if res[0]=="0":
                        a_err_txt=f"Ошибка при копировании ф. '{a_name}' из папки " \
                            "'multi_config' в рабочую папку."
                        printFAIL(a_err_txt)
                        keystrokeEnter()
                        return ["0", a_err_txt]
                    
                    res = readGonfigValue("opto_run.json", [], {}, workmode, "1")
                    if res[0] != "1":
                        return ["0", "Ошибка при чтении конфигурационных данных из файла."]
                    
                    default_value_dict = res[2]

                    readDefaultValue(default_value_dict)

                    innerMakeFileName(meter_serial_number)

                    print()
                    a_meter_tech_number=meter_tech_number_list_last[int(meter_pos)]
                    txt1=f"{bcolors.OKGREEN}Запись результатов проверки конфигурации " \
                        f"ПУ № {a_meter_tech_number}"
                    if number_of_meters>1:
                        txt1=txt1+f" (поз.{int(meter_pos)+1})"

                    txt1=txt1+f"."

                    res=innerSelectActions(txt1, [] , "7", "конфигурация ПУ",
                        [], [], rep_err_list)
                    if res[0]=="2":

                        for a_ind in range(0, len(menu_item_list)):
                            if meter_tech_number in menu_item_list[a_ind]:
                                menu_item_list[a_ind]=""

                        meter_tech_number_list_last[int(meter_pos)]=""
                        meter_status_test_list_last[int(meter_pos)]="ремонт"
                        meter_serial_number_list_last[int(meter_pos)]=""

                    
                    elif res[0]=="1":
                        if number_of_meters==1:
                            menu_item_list[0]=""
                            

                    
                    elif res[0]=="0":
                        break

                    elif res[0]=="9":
                        return ["9", "Проверка ПУ прервана."]

                    elif res[0]=="3" and len(res[2])>0 :
                        meter_config_res_list=res[2]
                        rep_err_list.extend(meter_config_res_list)
                        clipboard_err_list.extend(meter_config_res_list)

                        default_value_dict = writeDefaultValue(default_value_dict)
                        saveConfigValue(file_name_in="opto_run.json", 
                            var_config_dict=default_value_dict)
                        
                        for a_ind in range(0, len(menu_item_list)):
                            if meter_tech_number in menu_item_list[a_ind]:
                                menu_item_list[a_ind]=""

                        res=copyFile(opto_run_path, opto_run_multi_path, "0", "1")
                        if res[0]=="0":
                            a_err_txt="Ошибка при копировании ф. opto_run.json из " \
                                "рабочей папки в 'multi_config'."
                            printFAIL(a_err_txt)
                            keystrokeEnter()
                            return ["0", a_err_txt]


        res=copyFile(opto_run_path_last, opto_run_path, "0", "1")
        if res[0]=="0":
            a_err_txt="Ошибка при восстановлении ф. opto_run.json."
            printFAIL(a_err_txt)
            keystrokeEnter()
            return ["0", a_err_txt]
        
        res = readGonfigValue("opto_run.json", [], {}, workmode, "1")
        if res[0] != "1":
            return ["0", res[1]]
        
        default_value_dict=res[2]

        readDefaultValue(default_value_dict)

        meter_tech_number_list=meter_tech_number_list_last.copy()
        meter_status_test_list=meter_status_test_list_last.copy()
        meter_serial_number_list=meter_serial_number_list_last.copy()

        default_value_dict = writeDefaultValue(default_value_dict)
        saveConfigValue(file_name_in="opto_run.json", 
            var_config_dict=default_value_dict)
        
        res=copyFile(opto_run_path, opto_run_path_last, "0", "1")
        if res[0]=="0":
            a_err_txt="Ошибка при дублировании ф. opto_run.json."
            printFAIL(a_err_txt)
            keystrokeEnter()
            return ["0", a_err_txt]

        
        for a_tech in meter_tech_number_list:
            if a_tech!="":
                break
        
        else:
            return ["2", "Плановый выход из ПП."]
        

        _, ans2, file_name = getUserFilePath('otk_opto.py')
        if file_name == "":
            return ["0", f"Ошибка в ПП getUserFilePath(): {ans2}"]
        txt1=f"\n{bcolors.WARNING}Пожалуйста подождите, идет загрузка программы для " \
            f"дальнейшей проверки ПУ...{bcolors.ENDC}\n"
        print(txt1)

        command_line=f"python {file_name} 0"

        if modeStartMeterTest=="конфигурация":
            command_line=f"python {file_name} 1"
        
        code_exit = os.system(f"{command_line}")
    
        
        default_value_dict=optoRunVarRead()

        readDefaultValue(default_value_dict)

        return ["1", "Проверка ПУ успешно завершена."]
    
 



if  __name__ ==  '__main__' :   

    global var_all_value_dic
    global default_value_dict 

    os.system('CLS')

    title_old="otk_menu.bat"
    title_new="Аппаратная проверка ПУ"
    res = replaceTitleWindow(title_old, title_new)
    
    default_value_dict=getDefaultValue()
    
    workmode="эксплуатация"
    res = readGonfigValue("opto_run.json", [], default_value_dict, workmode, "1")
    if res[0] != "1":
        sys.exit()

    default_value_dict=res[2] 
    workmode=res[2]["workmode"]

    res = saveConfigValue("opto_run.json", default_value_dict, 
        workmode, "заменить часть")
    if res[0] != "1":
        sys.exit()

    res = readGonfigValue("var_all_value.json", [], {}, workmode, "1")
    if res[0] != "1":
        sys.exit()
    
    var_all_value_dic=res[2]

    res = changeUser(workmode)
    if res[0]!="1":
        print(f"\n{bcolors.WARNING}Проверка прервана.{bcolors.ENDC}")
        sys.exit()
    
    
    res = readGonfigValue("opto_run.json", [], default_value_dict, 
        workmode, "1")
    if res[0] != "1":
        sys.exit()


    default_value_dict=res[2]

    readDefaultValue(default_value_dict)
    
    com_current="com_opto"
    default_value_dict = writeDefaultValue(default_value_dict)
    saveConfigValue(file_name_in="opto_run.json", 
        var_config_dict=default_value_dict)

    serial_num=""

    print_number_big_font=default_value_dict["print_number_big_font"]

    cicl=True
    while cicl:
        os.system('CLS')

        if print_number_big_font=="окно":
            filename = 'otk_print_big_font.bat'
            closeProgram(filename)

        if speaker=="1":
            playsound("speech\\hello.mp3", block=False)

        menuMain()
        os.system('CLS')
        sys.exit()
