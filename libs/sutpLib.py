
def getAboutSutpLib():
    version = "13.04.2024 07:09"
    descript = "Библиотека ПП для обмена данными с СУТП"
    return [version, descript]


import time
import requests     #pip install requests для записи о результатах проверки в БД СУТП
from libs.sutpRequestTest import requestSUTPTest     #для получения тестовых данных

import json

from datetime import datetime, timedelta


class bcolors:
    HEADER = '\033[95m' #пурпурный
    OKBLUE = '\033[94m' #синий
    ATTENTIONBLUE = '\033[37m\033[44m'  #белый текст на синем фоне
    OKCYAN = '\033[96m' #цвет морской волны
    OKGREEN = '\033[92m'
    OKRESULT='\033[30m\033[42m' #для вывода результата проверки: черный текст на зеленом фоне
    WARNING = '\033[93m' #оранжевый
    FAIL = '\033[37m\033[41m' #белый текст на красном фоне
    ENDC = '\033[0m' #код сброса
    BOLD = '\033[1m'
    WHITE = '\033[97m'
    MAGENTA = '\033[95m' #пурпурный
    UNDERLINE = '\033[4m'


def request_sutp(type_request, url1, body_query1, err_txt="", workmode="эксплуатация",
                 print_err="1"):
    # url_base = "http://api.sutp.promenergo.local"  # базовый адрес РАБОЧИЙ
    url_base = "http://test.api.sutp.promenergo.local"  # базовый адрес РАБОЧИЙ
    if "test-api" in workmode:
        url_base = "http://test.api.sutp.promenergo.local"  # базовый адрес для ТЕСТОВ
    elif "тест" in workmode:
        res = requestSUTPTest(url1)
        if res[0]=="1":
            return ["1","", res[2]]
        
    url2 = f"{url_base}{url1}"

    request_header={"X-Client-Key": "technical-control-department-software",}

    ret = ["0","",[]]
    if err_txt != "":
        err_txt = f"{err_txt}\n"
    try:
        if type_request=="GET":
            response = requests.get(url2, headers=request_header, timeout=10)

        elif type_request=="GET_BEARER":
            request_header.update(body_query1)
            response = requests.get(url2, headers=request_header, timeout=10)

        else:
            response = requests.post(url2, headers=request_header, json=body_query1,
                timeout=10)
            
        if response.status_code!=200:
            serv_msg=response.reason
            msg_err=f"{str(response.status_code)}: {serv_msg}"
            err_color=bcolors.FAIL
            if response.status_code==404:
                err_color=bcolors.WARNING
            if print_err=="1":
                print(f"{err_color}Сервер сообщил об ошибке {msg_err}.{bcolors.ENDC}")
            return ["3",str(response.status_code),[]]
        try:
            val_json=response.json()
        except Exception:
            val_json=[]
        return ["1", response, val_json]
    
    except requests.exceptions.Timeout:
        err_txt = f"{err_txt}Ошибка: время ожидания ответа от сервера истекло"
        ret=["2","Timeout",[]]

    except requests.exceptions.HTTPError as err:
        err_txt = f"{err_txt}Ошибка HTTP: {err.response.status_code}"
        ret = ["0", str(err.response.status_code),[]]
    
    except Exception as e:
        err_txt = f"{err_txt}Ошибка при обмене данными с сервером СУТП: {e.args[0]}"
        ret = ["0", str(e.args[0]),[]]
    
    if print_err=="1":
        print(f"{bcolors.FAIL}{err_txt}{bcolors.ENDC}")
    return ret


def correctDateTime(dt: str, dt_format_in='%Y-%m-%dT%H:%M:%S',
        dt_format_out='%d.%m.%Y %H:%M',dont_to_correct="откл",
        weeks=0,days=0,hours=0,minutes=0,seconds=0):

    dt_datetime = datetime.strptime(dt, dt_format_in)
    dt_datetime_correct = dt_datetime
    if dont_to_correct=="откл":
        dt_datetime_correct = dt_datetime+ \
            timedelta(weeks=weeks, days=days,
                      hours=hours, minutes=minutes, 
                      seconds=seconds)
    date_time_txt = dt_datetime_correct.strftime(dt_format_out)
    return date_time_txt




def getNameEmployee(id_employee,workmode="эксплуатация"):
    url1=f"/api/Employee/{id_employee}"
    response=request_sutp("GET",url1,[],"",workmode)
    if response[0]=="1":
        val = response[2]
        fam=val["surName"]
        fio=fam
        nam=""
        patronymic = ""
        if val["name"] != "None" and \
            val["name"] != None:
            nam = val["name"]
            fio = f"{fio} {nam[0]}."
            if val["secondName"] != "None" and \
                val["secondName"] != None:
                patronymic=val["secondName"]
                fio = f"{fio}{patronymic[0]}."
        return ["1",fam,nam,patronymic, fio]
    elif response[0] == "2" or response[0] == "0":
        return ["2", response[1]]
    else:
        return ["0", response[1]]

def findNameMeterStatus(device_status_id,workmode="эксплуатация"):
    url1="/api/NSI/DeviceStatus/all"  
    response=request_sutp("GET",url1,[],"",workmode)
    if response[0]!="1":
        return ["0", response[1]]
    val = response[2]
    for i in val:
        if i["id"]==device_status_id:
            name_status=i["name"]
            return ["1",name_status]
    return ["0"]     


def getInfoAboutDockedMC(device_number:str, workmode="эксплуатация"):

    gsm_docked_tn_sutp=None
    gsm_docked_sn_sutp=""
    gsm_docked_model_sutp=""
    gsm_docked_soft_sutp=""
    device_number = str(device_number)
    techNumber = device_number

    url1 = f"/api/GsmModule/ByDockedDevice/ByTechNumber?deviceTechNumber={techNumber}"
    response = request_sutp(
        "GET", url1, [], err_txt="", workmode=workmode)
    if response[0] != "1":
        print(f"{bcolors.WARNING}Ошибка при получении информации о " \
            f"состыкованном МС из СУТП.{bcolors.ENDC}")
        return ["0", "Ошибка при получении информации из СУТП", None, 
                "", "", ""]
    else:
        mc_docked_dic=response[2]
        gsm_docked_tn_sutp=mc_docked_dic.get("techNumber", 0)
        gsm_docked_sn_sutp=mc_docked_dic.get("serialNumber", "")
        a_dic=mc_docked_dic.get("productIndividualType", {})
        gsm_docked_model_sutp=a_dic.get("name","")
        mc_docked_id_firmware=mc_docked_dic.get("firmwareId","")
        if mc_docked_id_firmware!="":
            a_txt="Ошибка при получении информации о " \
                "версии ПО МС из СУТП."
            res=getSoftDevice(mc_docked_id_firmware, a_txt, workmode)
            if res[0]!="1":
                return ["0", "Ошибка при получении информации о версии "
                        "ПО МС из СУТП", None, "", "", ""]
            gsm_docked_soft_sutp=res[2]
        return ["1", "Информация получена успешно.", gsm_docked_tn_sutp,
                gsm_docked_sn_sutp, gsm_docked_model_sutp, 
                gsm_docked_soft_sutp]



def getSoftDevice(id_firmware: str, txt_err="", workmode="эксплуатация"):

    url1 = f"/Firmware/Get/Id={id_firmware}"
    response = request_sutp(
        "GET", url1, [], err_txt="", workmode=workmode)
    if response[0] != "1":
        if txt_err!="":
            print(f"{bcolors.WARNING}{txt_err}{bcolors.ENDC}")
        return ["0", "Ошибка при получении информации о версии ПО.", ""]
    a_list=response[2].get("items",[])
    a_dic=a_list[0]
    soft_sutp=a_dic.get("version","")
    return ["1", "Версия ПО устройства получена.", soft_sutp]



def getInfoAboutDevice(device_number, workmode="эксплуатация", user_id="", 
    user_pw_encrypt="", query_mode="0"):



    from libs.otkLib import cryptStringSec  #для зашифровки пароля подклчения к ПУ

    ret="1"
    err_list=[]
    serialNumber=""
    order_descript=""
    order_ev="0"
    order_pw_type_reader_descript=""
    order_pw_reader_assigned=""
    order_pw_type_reader_id=None
    pw_reader_encrypt=""
    order_pw_type_config_descript=""
    order_pw_config_assigned=""
    order_pw_type_config_id=None
    pw_config_encrypt=""
    order_pw_dic={}
    device_config_pw=""
    device_pw_encrypt=""
    device_reader_pw=""
    device_reader_pw_encrypt=""
    gsm_docked_tn_sutp=None
    gsm_docked_sn_sutp=""
    gsm_docked_model_sutp=""
    gsm_docked_soft_sutp=""
    device_number = str(device_number)
    techNumber = device_number
    url1 = f"/api/Device/{device_number}"
    if len(device_number)==15 or len(device_number)==13:
        url1 = f"/api/Device/BySerialNumber/{device_number}"
        serialNumber = device_number
        techNumber=""
    response = request_sutp("GET", url1, [], "", workmode)
    if response[0]!="1":
        return ["0",response[0]]
    val = response[2]
    techNumber=val["techNumber"]        #технический номер счетчика
    serialNumber=val['serialNumber']  #серийный номер счетчика
    meter_soft_sutp=""
    if val['firmware']!=None:
        meter_soft_sutp=val['firmware']['version']
    deviceType=val["productIndividualType"]
    model_device=deviceType.get("name", "")
    generalDeviceType=val["productGeneralType"]
    formOfDeviceTypeId=val["productGeneralTypeId"]
    orderNumber=val['orderNumber']
    order_dic={}
    if orderNumber!=None:
        order_dic=val.get("order",{})
    order_descript=order_dic.get("orderName", "")

    a_dic={False:"0", True:"1"}
    a_ev=order_dic.get("isRetailOrder", False)
    order_ev=a_dic.get(a_ev, "0")

    order_pw_type_reader_id=order_dic.get("passwordTypeLLSId", "")
    order_pw_type_config_id=order_dic.get("passwordTypeHLSId", "")
    
    deviceStatusId=val["productStatusId"]   
    deviceStatusName=""
    
    res = findNameMeterStatus(deviceStatusId,workmode)
    if res[0]!="1":
        ret="2"
        err_list.append(res[1])
    else:
        deviceStatusName=res[1]
    
    if query_mode=="1":
        return [ret, "\n".join(err_list), techNumber, serialNumber, deviceType, formOfDeviceTypeId, deviceStatusId, 
        deviceStatusName, orderNumber, order_descript, device_config_pw, device_pw_encrypt,
        model_device, gsm_docked_tn_sutp, gsm_docked_sn_sutp, gsm_docked_model_sutp, 
        gsm_docked_soft_sutp, meter_soft_sutp, device_reader_pw_encrypt, device_reader_pw,
        order_pw_dic, order_ev]
    

    res=getInfoAboutDockedMC(techNumber, workmode)
    if res[0] != "1":
        ret="2"
        err_list.append(res[1])
    else:
        gsm_docked_tn_sutp=res[2]
        gsm_docked_sn_sutp=res[3]
        gsm_docked_model_sutp=res[4]
        gsm_docked_soft_sutp=res[5]
        

    if query_mode=="2":
        return [ret, "\n".join(err_list), techNumber, serialNumber, deviceType, formOfDeviceTypeId, deviceStatusId, 
        deviceStatusName, orderNumber, order_descript, device_config_pw, device_pw_encrypt,
        model_device, gsm_docked_tn_sutp, gsm_docked_sn_sutp, gsm_docked_model_sutp, 
        gsm_docked_soft_sutp, meter_soft_sutp, device_reader_pw_encrypt, device_reader_pw, 
        order_pw_dic, order_ev]
    
    if user_id!="" and user_pw_encrypt!="":
        res=getDevicePw(techNumber, user_id, user_pw_encrypt, workmode, "1")
        if res[0]!="1":
            print(f"{bcolors.WARNING}Ошибка при получении информации о " \
                f"пароле подключения к ПУ из СУТП.{bcolors.ENDC}")
            ret="2"
            err_list.append(res[1])
        else:
            device_pw_dic=res[2]
            device_config_pw=device_pw_dic.get("configkey","")
            device_reader_pw=device_pw_dic.get("readerkey","")
            if device_config_pw=="":
                txt1='Сервер вернул пустой пароль высокого уровня.'
                print (f'{bcolors.WARNING}{txt1}{bcolors.ENDC}')
            if device_reader_pw=="":
                txt1='Сервер вернул пустой пароль низкого уровня.'
                print (f'{bcolors.WARNING}{txt1}{bcolors.ENDC}')

        if orderNumber!=None:
            res=getOrderPw(orderNumber, user_id, user_pw_encrypt, workmode, "1")
            if res[0] != "1":
                ret="2"
                err_list.append(res[1])
            else:
                a_dic=res[2]
                order_pw_config_assigned=a_dic.get("passwordHLS", "")
                order_pw_reader_assigned=a_dic.get("passwordLLS", "")
                a_dic={0: "заводской пароль", 1: "пароль заказчика", 2: "сгенерированный пароль"}
                order_pw_type_reader_descript=a_dic.get(order_pw_type_reader_id, "")
                order_pw_type_config_descript=a_dic.get(order_pw_type_config_id, "")
                if order_pw_type_reader_id==1 and order_pw_reader_assigned!=None:
                    res=cryptStringSec("зашифровать", order_pw_reader_assigned)
                    pw_reader_encrypt=res[2]
                
                if order_pw_type_config_id==1 and order_pw_config_assigned!=None:
                    res=cryptStringSec("зашифровать", order_pw_config_assigned)
                    pw_config_encrypt=res[2]

    order_pw_dic={"pw_type_reader_id": order_pw_type_reader_id, 
        "pw_type_reader_descript": order_pw_type_reader_descript, 
        "pw_reader_assigned": order_pw_reader_assigned,
        "pw_reader_encrypt": pw_reader_encrypt, 
        "pw_type_config_id": order_pw_type_config_id, 
        "pw_type_config_descript": order_pw_type_config_descript,
        "pw_config_assigned": order_pw_config_assigned,
        "pw_config_encrypt": pw_config_encrypt}
    

    if device_config_pw!="":
        res=cryptStringSec("зашифровать", device_config_pw)
        device_pw_encrypt=res[2]
    
    if device_reader_pw!="":
        res=cryptStringSec("зашифровать", device_reader_pw)
        device_reader_pw_encrypt=res[2]

    return [ret, "\n".join(err_list), techNumber, serialNumber, deviceType, formOfDeviceTypeId, deviceStatusId, 
        deviceStatusName, orderNumber, order_descript, device_config_pw, device_pw_encrypt,
        model_device, gsm_docked_tn_sutp, gsm_docked_sn_sutp, gsm_docked_model_sutp, 
        gsm_docked_soft_sutp, meter_soft_sutp, device_reader_pw_encrypt, device_reader_pw,
        order_pw_dic, order_ev]




def preChecksToGhangeStatusMeter(device_number: int,
        device_status_id:int, workmode="эксплуатация", print_err_msg="1"):

    res = getInfoAboutDevice(device_number, workmode, "", "", "1")
    if res[0]=="0":
        return ["0","Сервер вернул ошибку.","",0]
    tech_number = res[2]
    serial_number=res[3]
    meter_status_id=res[6]
    meter_status_name=res[7]

    url1 = f"/api/device/preChecks"
    body_query1 = {
        "statusIds": [
            device_status_id
        ],
        "techNumber": tech_number,
    }
    
    response = request_sutp("POST", url1, body_query1, "", workmode=workmode)
    if response[0] == "3":
        if print_err_msg=="1":
            txt1_1 = f"{bcolors.FAIL}На данном этапе с устройством № {tech_number} ({serial_number})" \
                f" проводить работы нельзя.{bcolors.ENDC}\n" \
                f"{bcolors.FAIL}Текущий статус ПУ: {meter_status_name}.{bcolors.ENDC}"
            print(f"{txt1_1}")
        return ["3","Операцию с ПУ проводить нельзя.",meter_status_name,meter_status_id]
    elif response[0] == "2":
        return ["2","Таймаут превышен.","",0]
    elif response[0] == "0":
        return ["0","Сервер вернул ошибку.","",0]
    return ["1","Операция выполнена успешно.", meter_status_name,meter_status_id]



def savetToSUTP2(device_number: int, createdEmployeeId, operation: str, 
    device_status_id=21, defect_description="", workmode="эксплуатация",
    preChecks="1", species_of_device="счетчик"):
    

    from libs.otkLib import getUserFilePath, getDataFromXlsx, findStrInList
    
    def innerGetDefectCategoryList(defect_description: str, 
        createdEmployeeIdInt):

        device_defects_sutp_list=[]
        category_defect_sutp_list=[]
        defect_sutp_dic={}
        
        if defect_description=="" or defect_description==None:
            return []

        res=getUserFilePath("QR_list.xlsx", "0", workmode)
        file_path=res[2]
        if file_path == "":
            return []

        sheet_name="Defects"
        res=getDataFromXlsx(file_path, sheet_name, {}, "")
        if res[0]!="1":
            return []
        
        xls_list=res[2]

        defects_list=defect_description.split("\n")

        keywords_list=[]
        keywords_str=""

        id_defect_category_def="c3007fbb-9ea6-4a87-a005-c47d2ca1f82d"

        for i in range(0, len(xls_list)):
            if len(defects_list)==0:
                break

            row_cur_dic=xls_list[i]

            keywords_str=row_cur_dic.get("keywords", "")

            id_defect_category=row_cur_dic.get("idDefectCategory", 
                id_defect_category_def)
            
            if keywords_str!="" and keywords_str!=None:
                keywords_list=keywords_str.split(",")
                for keyword in keywords_list:
                    res=findStrInList(keyword, defects_list, "0", "0")
                    if res[0]=="1":
                        category_defect_sutp_list=defect_sutp_dic.get(id_defect_category, [])
                        category_defect_sutp_list.append(res[2])
                        defect_sutp_dic[id_defect_category]=category_defect_sutp_list

                        del defects_list[res[3]]

                        if len(defects_list)==0:
                            break
        
        if len(defects_list)!=0:
            a_str="\n".join(defects_list)
            category_defect_sutp_list=defect_sutp_dic.get(id_defect_category_def, [])
            category_defect_sutp_list.append(a_str)
            defect_sutp_dic[id_defect_category_def]=category_defect_sutp_list
        
        a_list=defect_description.split("\n")
        a_id="5f0745b8-029b-4559-a711-7b5b65a3390f"
        defect_sutp_dic[a_id]=a_list
        
        id_defect_category_list=list(defect_sutp_dic.keys())
        for id_defect_category in id_defect_category_list:
            a_list=defect_sutp_dic[id_defect_category]
            a_str="\n".join(a_list)
            a_dic={
                "description": a_str,
                "defectId": id_defect_category
                }
            device_defects_sutp_list.append(a_dic)

        return device_defects_sutp_list


    
    techNumber = ""
    res = getInfoAboutDevice(device_number, workmode, "", "", "1")
    if res[0]=="0":
        return "0"
    techNumber = res[2]
    formOfDeviceTypeId=res[5]

    createdEmployeeIdInt=int(createdEmployeeId)

    if operation == "0":
        res=innerGetDefectCategoryList(defect_description, createdEmployeeIdInt)
        if len(res)==0:
            return "0"
        
        device_defects_list=res

        url1 = "/RepairCycle/Create"
        if species_of_device=="счетчик":
            body_query1={
                "techNumber": techNumber,
                "statusToSet": 399,
                "statusWasToSet": 21,
                "createdEmployeeId": createdEmployeeIdInt,
                "productDefects": device_defects_list
                }
            err_txt="Отправка счетчика в ремонт."
        else:
            body_query1={
                "techNumber": techNumber,
                "statusToSet": 399,
                "statusWasToSet": 21,
                "createdEmployeeId": createdEmployeeIdInt,
                "deviceDefects": [],
                "productDefects": device_defects_list
                }
            err_txt="Отправка модуля связи в ремонт."

        print ("Отправляю запрос серверу об отправке изделия в ремонт...")

        response = request_sutp("POST", url1, body_query1,
                                err_txt=err_txt, workmode=workmode)
        
        if response[0] in ["3", "2", "1"]:
            return response[0]
        else:
            return "0"

    elif operation == "1":
        if preChecks=="1":
            res=preChecksToGhangeStatusMeter(device_number=device_number,
                device_status_id=device_status_id,workmode=workmode,print_err_msg="1")
            if res[0]!="1":
                ret=res[0]
                if ret=="3":
                    ret="4"
                return ret  
        url1 = "/api/device/changeStatus"
        body_query1 = {
            "employeeId": createdEmployeeId,
            "techNumber": techNumber,
            "deviceStatusId": device_status_id
        }
        print ("Отправляю запрос серверу об изменении статуса изделия...")
        response = request_sutp("POST", url1, body_query1,
                                err_txt="", workmode=workmode)
        if response[0] == "1":
            return "1"
        elif response[0] == "2":
            res = getInfoAboutDevice(techNumber, workmode, "", "", "1")
            if res[0]=="1":
                device_status_id_SUTP=res[6]
                if device_status_id != "" and device_status_id_SUTP == device_status_id:
                    return "1"
                else:
                    return "0"
            elif res[0]=="2":
                return "2"
            else:
                return "0"
        else:
            return "0"
    else:
        return "0"





def getDeviceHistory(device_number: str, employee_print="0",workmode="эксплуатация"):

    device_history_dic={}

    device_number=str(device_number)
    url1 = f"/DeviceHistory/Get/TechNumber={device_number}"
    if len(device_number) == 15 or len(device_number)==13:
        url1 = f"/DeviceHistory/Get/SerialNumber={device_number}"
    response = request_sutp("GET", url1, [], "", workmode)
    if response[0] != "1":
        return ["0","Ошибка при обмене данными с сервером"]
    val = response[2]
    items = val['items']
    device_history=""
    for item in items:
        version =str(item["version"])
        dt_from = item["validFrom"][0:19]
        date_time_txt = correctDateTime(dt_from, '%Y-%m-%dT%H:%M:%S',
            '%d.%m.%Y %H:%M',"откл",hours=3)
        date_time_full_txt = correctDateTime(dt_from, '%Y-%m-%dT%H:%M:%S',
            '%d.%m.%Y %H:%M:%S',"откл",hours=3)
        device_status_dic = item["productStatus"]
        device_status_id = str(device_status_dic["id"])
        device_status_name = device_status_dic["name"]
        employee_dic = item["employee"]
        employee_id=str(employee_dic.get("id",""))

        employee_name="Ф.И.О. не указаны"

        a_name=employee_dic.get("name","")
        if a_name!=None and len(a_name.lstrip(" "))>0:
            employee_name_1 = a_name

        employee_name_2=""
        a_name=employee_dic.get("surName","")
        if a_name!=None and len(a_name.lstrip(" "))>0:
            employee_name_2 = a_name

        employee_name_3=""
        a_name=employee_dic.get("secondName","")
        if a_name!=None and len(a_name.lstrip(" "))>0:
            employee_name_3 = a_name

        if employee_name_2!="":
            employee_name = employee_name_2
            
            if employee_name_1!="":
                employee_name=f"{employee_name} {employee_name_1[0]}."

            if employee_name_3!="":
                employee_name = f"{employee_name}{employee_name_3[0]}."

        elif employee_name_1!="":
            employee_name=employee_name_1

        elif employee_name_3!="":
            employee_name=employee_name_3
        
        rec = f"{version}. {date_time_txt} {device_status_name}"
        if employee_print=="1":
            rec=f"{rec} ({employee_id} {employee_name})"
        device_history = f"{device_history}{rec}\n"

        device_history_dic[version]={"date_time": date_time_full_txt,
            "device_status_id": device_status_id,
            "device_status": device_status_name,
            "employee_id": employee_id, 
            "employee_name": employee_name}

    return ["1", "История ПУ сформирована", device_history, device_history_dic]


    

def getDeviceRepayHistory(device_number: str, workmode="эксплуатация"):
    
    device_number=str(device_number)
    if len(device_number) == 15 or len(device_number)==13:
        res=getInfoAboutDevice(device_number,workmode,"","", "1")
        if res[0]=="0":
            return ["0","Ошибка при получении технического номера ПУ."]
        device_number=res[2]
    size = 50
    page = 0
    cicl=True
    ret_txt=""
    rec_number=1
    while cicl:
        url1=f"/RepairCycle/Get/TechNumber={device_number}/Page={page}/Size={size}"
        response = request_sutp(
            "GET", url1, [], err_txt="", workmode=workmode)
        if response[0] != "1":
            print ("Сервер вернул сообщение об ошибке.")
            return ["0","Ошибка при получении информации о ремонте ПУ."]
        val = response[2]
        count_rec=val['count']
        ret_txt = ""
        ret_txt_list=[]
        for item in val['items']:
            device_defects_list=item['productDefects']
            device_defect_description=""
            for j in range(0,len(device_defects_list)):
                device_defects_dic=device_defects_list[j]
                defect_dic=device_defects_dic['defect']
                defect_name=defect_dic['name']
                if device_defects_dic['description']!=None:
                    device_defect_description=device_defects_dic['description']
                if defect_name=="Другое":
                    defect_name=device_defect_description
                created_employee_id=device_defects_dic['createdEmployeeId']
                created_employee_name=""
                res = getNameEmployee(created_employee_id, workmode)
                if res[0]=="1":
                    created_employee_name=res[4]
                create_date=device_defects_dic['createDate'][0:19]
                create_date = correctDateTime(
                    create_date, '%Y-%m-%dT%H:%M:%S',
                    '%d.%m.%Y %H:%M',"откл",hours=3)
                fixed_employee_id=0
                fix_date=""
                if device_defects_dic['fixedEmployeeId']!=None:
                    fixed_employee_id=device_defects_dic['fixedEmployeeId']
                    fix_date=device_defects_dic['fixDate'][0:19]
                    fix_date = correctDateTime(
                        fix_date, '%Y-%m-%dT%H:%M:%S',
                        '%d.%m.%Y %H:%M',"откл",hours=3)
                rec_number_len = len(str(rec_number))
                space_num = " "*(rec_number_len+3)
                if fix_date=="":
                    fix_date="настоящее время"
                txt_1 = f"Несоответствие: {defect_name} "\
                    f"({created_employee_name})\n" \
                    f"{space_num}В ремонте: с {create_date} по {fix_date}\n"
                if txt_1 not in ret_txt_list:
                    ret_txt_list.append(txt_1)
                    ret_txt=f"{ret_txt} {rec_number}. {txt_1}"
                    rec_number+=1
        if count_rec<50:
            break
        page+=1
    return ["1","Информация получена",ret_txt]
    


def getDevicePw(device_tech_number:str, user_id: str, 
    user_pw_encrypt: str, workmode="эксплуатация", print_err="1"):

    from libs.otkLib import cryptStringSec  #для расшифровки пароля сотрудника

    if user_id=="" or user_pw_encrypt=="":
        txt1="Получение паролей доступа к ПУ: отсутствует табельный номер " \
            "или пароль сотрудника."
        if print_err=="1":
            print(f"{bcolors.FAIL}{txt1}{bcolors.ENDC}")
        return ["0", txt1, {}]
    
    res=cryptStringSec("расшифровать", user_pw_encrypt)
    user_pw=res[2] 
    
    url1="/api/Auth/login"
    body_query1={"id": int(user_id), "password": user_pw}
    response=request_sutp("POST", url1, body_query1, "", workmode, print_err)
    if response[0] != "1":
        return ["0", response[1], {}]
    token=response[2]["token"]

    device_pw_dic={}

    url1=f"/PasswordStorage/Get/TechNumber={int(device_tech_number)}"
    headers = {'Authorization': f'Bearer {token}',}
    response=request_sutp("GET_BEARER", url1, headers, "", workmode, print_err)
    if response[0] != "1":
        return ["0", response[1], {}]
    if 'Error' in response[2]["status"]:
        msg_err=response[2]["message"]
        if print_err=="1":
            print(f"{bcolors.FAIL}Сервер сообщил об ошибке: {msg_err}.{bcolors.ENDC}")
        return ["3", msg_err, {}]
    elif 'Success' in response[2]["status"]:
        pw_dic=response[2]["value"]
        device_pw_dic['configkey']=pw_dic.get("configkey","")
        device_pw_dic['readerkey']=pw_dic.get("readerKey","")
        return ["1", "Пароли для устройства получены.", device_pw_dic]
    msg_err=f'Сервер вернул неизвестный статус: {response[2]["status"]}.'
    if print_err=="1":
        print(f"{bcolors.FAIL}{msg_err}{bcolors.ENDC}")
    return ["0", msg_err, {}]



261124
def getOrderPw(order_number:str, user_id: str, 
    user_pw_encrypt: str, workmode="эксплуатация", print_err="1"):

    from libs.otkLib import cryptStringSec  #для расшифровки пароля сотрудника

    if user_id=="" or user_pw_encrypt=="":
        txt1="Получение паролей доступа из заказа: отсутствует табельный номер " \
            "или пароль сотрудника."
        if print_err=="1":
            print(f"{bcolors.FAIL}{txt1}{bcolors.ENDC}")
        return ["2", txt1, {}]
    
    if order_number=="" or order_number==None or order_number=="None":
        txt1="Отсутствует номер заказа."
        if print_err=="1":
            print(f"{bcolors.FAIL}Получение паролей доступа из заказа: "
                  f"{txt1}{bcolors.ENDC}")
        return ["2", txt1, {}]

    res=cryptStringSec("расшифровать", user_pw_encrypt)
    user_pw=res[2] 
    
    url1="/api/Auth/login"
    body_query1={"id": int(user_id), "password": user_pw}
    response=request_sutp("POST", url1, body_query1, "", workmode, print_err)
    if response[0] != "1":
        return ["0", response[1], {}]
    token=response[2]["token"]

    device_pw_dic={}

    url1=f"/api/Order/GetPasswordsByOrderNumber/{int(order_number)}"
    headers = {'Authorization': f'Bearer {token}',}
    response=request_sutp("GET_BEARER", url1, headers, "", workmode, print_err)
    if response[0] != "1":
        return ["0", response[1], {}]
    pw_dic=response[2]
    return ["1", "Пароли из заказа получены.", pw_dic]



def getMeterAllSN_OLD(device_number:str, workmode="эксплуатация",
    print_err="1"):

    device_number=str(device_number)
    if len(device_number) == 15 or len(device_number)==13:
        res=getInfoAboutDevice(device_number,workmode,"","", "1")
        if res[0]=="0":
            return ["0","Ошибка при получении технического номера ПУ.",
                    [], ""]
        device_number=res[2]

    device_all_sn_list=[]
    url1=f"/DeviceHistory/{device_number}/oldSerialNumbers"
    res = request_sutp("GET", url1, [], 
        err_txt="", workmode=workmode)
    if res[0] != "1":
        if print_err=="1":
            print ("Сервер вернул сообщение об ошибке.")
        return ["0","Ошибка при получении всех серийных номеров ПУ.", 
                [], ""]
    device_all_sn_list = res[2]
    all_sn_list=[]
    for sn_dic in device_all_sn_list:
        sn=sn_dic['serialNumber']
        dt=sn_dic['dateTime'][0:19]
        dt = correctDateTime(dt, '%Y-%m-%dT%H:%M:%S',
            '%d.%m.%Y %H:%M:%S',"откл",hours=3)
        all_sn_list.append(f"{dt}: {sn}")
    all_sn_txt="\n".join(all_sn_list)
    return ["1", "Данные о серийных номерах сформированы.", device_all_sn_list,
            all_sn_txt]



def getMeterAllSN(device_number_str:str, workmode="эксплуатация",
    print_err="1"):

    device_number_str=str(device_number_str)

    device_all_sn_list=[]
    res = getProductHistoryByAnyKey(device_number_str, workmode, print_err)
    
    if res[0] != "1":
        return ["0","Ошибка при получении всех серийных номеров ПУ.", 
                [], ""]
    
    device_all_sn_list = res[2].get("serialNumberHistory", [])

    all_sn_list=[]
    for sn_dic in device_all_sn_list:
        sn=sn_dic['serialNumber']
        dt=sn_dic['validFrom'][0:19]
        dt = correctDateTime(dt, '%Y-%m-%dT%H:%M:%S',
            '%d.%m.%Y %H:%M:%S',"откл",hours=3)
        all_sn_list.append(f"{dt}: {sn}")
    all_sn_txt="\n".join(all_sn_list)
    
    return ["1", "Данные о серийных номерах сформированы.", device_all_sn_list,
            all_sn_txt]


    
def getProductHistoryByAnyKey(device_number_str:str, workmode="эксплуатация",
    print_err="1"):

    device_number_str=str(device_number_str)

    device_number=int(device_number_str)

    url1=f"/api/ProductHistories/ByAnyKey"

    body_query1={"techNumber": device_number}

    if len(device_number_str) == 15 or len(device_number_str)==13:
        body_query1={"serialNumber": device_number_str}

    err_txt="Получение истории об изделии."

    res = request_sutp("POST", url1, body_query1,
        err_txt=err_txt, workmode=workmode)
    
    if res[0] != "1":
        if print_err=="1":
            print ("Сервер вернул сообщение об ошибке.")
        return ["0","Ошибка при получении истории о ПУ.", {}]

    return ("1", "Информация об изделии успешно получена.", res[2])



def getUrlMeterConfigFile(meter_tech_number:str, print_err="1",
                          workmode="эксплуатация"):

    meter_tech_number=int(meter_tech_number)
    
    resource_type_id ="5a923bd7-9352-49dd-b86d-5fa21ef193e2"
    url1=f"/Resource/Get/TechNumber={meter_tech_number}/" \
        f"ResourceTypeId={resource_type_id}/Page=0/Size=1000"
    res = request_sutp("GET", url1, [], 
        err_txt="", workmode=workmode)
    if res[0] != "1":
        if print_err=="1":
            print ("Сервер вернул сообщение об ошибке.")
        return ["0","Ошибка при получении от сервера СУТП ссылки "
                "на файл конфигурации ПУ.", "", ""]
    
    elif len(res[2]["items"])==0:
        err_txt="Сервер не предоставил ссылку на файл конфигурации."
        if print_err=="1":
            print (f"{bcolors.FAIL}{err_txt}{bcolors.ENDC}")
        return ["0", err_txt, "", ""]
    
    items_list = res[2]["items"]
    url=items_list[0]["url"]
    sha256=items_list[0]["checksum"]

    return ["1", "Ссылка на файл конфигурации ПУ успешно получена.",
            url, sha256]



def getMeterConfigFileName(meter_tech_number:str, print_err="1",
    workmode="эксплуатация"):

    meter_tech_number=int(meter_tech_number)
    
    url1=f"/DeviceConfigurationInfo/Get/DeviceTechNumber={meter_tech_number}"

    res = request_sutp("GET", url1, [], 
        err_txt="", workmode=workmode)
    if res[0] != "1":
        if print_err=="1":
            print ("Сервер вернул сообщение об ошибке.")
        return ["0","Ошибка при получении от сервера СУТП имени файла, "
                "который применялся при конфигурировании ПУ.", "", ""]
    
    elif len(res[2]["items"])==0:
        err_txt="Сервер не предоставил имя файла, " \
                "который применялся при конфигурировании ПУ."
        if print_err=="1":
            print (f"{bcolors.FAIL}{err_txt}{bcolors.ENDC}")
        return ["0", err_txt, "", ""]
    
    items_list = res[2]["items"]
    file_name=items_list[0]["fileName"]

    return ["1", "Имя файла конфигурации ПУ успешно получено.",
            file_name]



def downloadFileURL(url:str, file_path:str, print_err_msg="1", 
                    sha256=None, max_tries=2):

    
    import os
    from libs.otkLib import pause_ui
    from libs.otkLib import checksumFile
    from libs.otkLib import printWARNING

    import hashlib
    dest_folder=os.path.split(file_path)[0]

    if not os.path.exists(dest_folder):
        try:
            os.makedirs(dest_folder)
        except Exception as e:
            a_txt=f"При создании папки '{dest_folder}' возникла ошибка: " \
                f"{e}"
            if print_err_msg=="1":
                print(f"{bcolors.FAIL}{a_txt}{bcolors.ENDC}")
            return["0", a_txt]

    url_base="http://api.sutp.promenergo.local"
    if not "http://" in url:
        url=f"{url_base}{url}"
    
    time_pause=3
    tries=1
    while True:
        res = requests.get(url, stream=True)
        if res.ok:
            try:
                with open(file_path, 'wb') as f:
                    for chunk in res.iter_content(chunk_size=1024 * 8):
                        if chunk:
                            f.write(chunk)
                            f.flush()
                            os.fsync(f.fileno())
                break
            except Exception as e:
                a_txt=f"При сохранении файла '{file_path}' возникла ошибка: " \
                    f"{e}"
                if print_err_msg=="1":
                    print(f"{bcolors.FAIL}{a_txt}{bcolors.ENDC}")
                return["0", a_txt]
        else:  # HTTP status code 4XX/5XX
            a_txt=f"При загрузке файла '{url}' возникла ошибка: " \
                    f"{str(res.status_code)} {res.text}"
            if print_err_msg=="1":
                print(f"{bcolors.FAIL}{a_txt}{bcolors.ENDC}")
            if tries<max_tries:
                if print_err_msg=="0":
                    print(f"{bcolors.WARNING}{a_txt}{bcolors.ENDC}")
                tries+=1
                print (f"Попытка загрузки файла №{tries} из {max_tries}.")
                pause_ui(time_pause)
                continue
            return["0", a_txt]

    if sha256!=None and sha256!="":
        res=checksumFile(file_path, sha256, print_err_msg)
        if res[0]=="1":
            return ["1", "Файл успешно загружен."]
        
        elif res[0]=="2":
            printWARNING("Контрольная сумма загруженного из СУТП " \
                "файла отличается от эталонного значения.")
            
        return["2", res[1]]
    
    return ["1", "Файл успешно загружен."]