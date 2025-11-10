

import sys
from libs.otkLib import *
import hashlib

from win32gui import SetForegroundWindow, ShowWindow, GetWindowRect, MoveWindow
from win32con import SW_MINIMIZE, SW_MAXIMIZE
from tqdm import tqdm   #pip install tqdm   для отображения progress bar

import keyboard     #для имитации нажатия клавиш

def executeMassProdAutoConfig_OLD(print_msg_err="1", 
    time_wait_exec=2):

    filename = 'MassProdAutoConfig.exe'
    closeProgram(filename)
    

    txt1 = "Запускаем программу MassProdAutoConfig.exe..."
    wait_key = ""
    res = openFileWaitKey(filename, "", txt1, "0",
        wait_key, "1", time_wait_exec)
    if res[0] == "0":
        a_err_txt = 'При запуске программы "MassProdAutoConfig.exe" ' \
            'возникла ошибка.'
        printMsgWait(a_err_txt, bcolors.WARNING, 
            print_msg_err )
        return ["0", a_err_txt]
    
    return ["1", "Программа запущена."]



def executeMassProdAutoConfig(meter_tech_number_list:list, 
    employee_id:str, pw: str, print_msg_err="1", 
    time_wait_exec=2):

    filename = 'MassProdAutoConfig.exe'

    closeProgram(filename)
    

    filename = f'MassProdAutoConfig.exe l:{employee_id} pass:{pw} StartProcess'

    for meter_tech_number in meter_tech_number_list:
        if meter_tech_number!=None and meter_tech_number!="":
            filename=f"{filename} tnums:{meter_tech_number}"
    
    filename=f"{filename} tnums:{0}"

    txt1 = "Запускаем программу MassProdAutoConfig.exe..."
    wait_key = ""
    res = openFileWaitKey(filename, "", txt1, "0",
        wait_key, "1", time_wait_exec)
    if res[0] == "0":
        a_err_txt = 'При запуске программы "MassProdAutoConfig.exe" ' \
            'возникла ошибка.'
        printMsgWait(a_err_txt, bcolors.WARNING, 
            print_msg_err )
        return ["0", a_err_txt]
    
    return ["1", "Программа запущена."]



def authMassProdAutoConfig(employee_id:str, pw:str, 
        workmode="эксплуатация", print_msg_err="1",
        time_wait_open_window=2):

    employee_id=(5-len(employee_id))*'0'+employee_id
    
    _,txt1, image_dir = getUserFilePath('mass_config_image_dir',
        only_dir="1", workmode=workmode)
    if image_dir=="":
        a_err_txt = "Не найден путь к папке с файлом-рисунком " \
            "для входа в программу MassProdAutoConfig.exe."
        printMsgWait(a_err_txt, bcolors.WARNING, 
            print_msg_err)
        return ["0", a_err_txt]


    windows_list=[["AuthGUI", "AuthGUI.png", 0]]

    for window_start in windows_list:
        window_title=window_start[0]
        w_image=window_start[1]
        image_path = os.path.join(image_dir, w_image)
        err_next_step=window_start[2]
        
        res=actionsSelectedtWindow([window_title], None,
            "показать+активировать","1")
        if res[0]!="1":
            if err_next_step=="0":
                a_err_txt=f"Не удалось найти окно '{window_title}'."
                printMsgWait(a_err_txt, bcolors.WARNING, print_msg_err)
                return ["3", a_err_txt]
        
            elif err_next_step=="1":
                continue
        

        pause_ui(time_wait_open_window)

        res=searchImageWindow(image_path, window_title, "1")
        if res[0]!="1":
            a_err_txt=f"Не удалось найти изображение '{w_image}'."
            printMsgWait(a_err_txt, bcolors.WARNING, print_msg_err)
            return ["3", a_err_txt]


        window_position=res[2]
        image_location=res[3]
        image_wh=res[4]
            
        mouse_x=image_location[0] + window_position[0]+image_wh[0]/2
        mouse_y= image_location[1] + window_position[1]+image_wh[1]/2

        res=mouseClickGraf(mouse_x, mouse_y, "1")
        if res[0]=="0":
            a_err_txt=f"Не удалось кликнуть мышкой."
            printMsgWait(a_err_txt, bcolors.WARNING, print_msg_err)
            return ["3", a_err_txt]

        if window_title=="AuthGUI":

                print("Ввожу регистрационные данные...")
                keys_list=list(employee_id)+list(pw)+['tab','space']

                for key in keys_list:
                    keyboard.send(key)
                    time.sleep(0.2)

                pause_ui(time_wait_open_window)

                a_title_window="MassProdAutoConfigGUI"
                res=searchTitleWindow(a_title_window)
                if res[0]=="1":
                    break

                elif res[0]!="2":
                    a_err_txt=f"При регистрации в программе возникла ошибка"
                    printMsgWait(a_err_txt, bcolors.WARNING, print_msg_err)
                    return ["0", a_err_txt]
                
                elif res[0]=="2":
                    keyboard.send("space")
                    a_err_txt=f"Не удалось зарегистрироваться в программе."
                    return ["2", a_err_txt]

    return ["1", "Регистрация прошла успешно."]



def cycleOnlineReadLog_OLD(file_log_path: str, meter_tech_number_list_in: list,
    meter_status_test_list_in: list, meter_serial_number_list_in: list, 
    mass_number_of_meter: int, meter_soft_list_in: list,
    workmode="эксплуатация", print_msg_err="1"):


    global programm_status 
  
    meter_tech_number_list=meter_tech_number_list_in.copy()

    meter_status_test_list=meter_status_test_list_in.copy()

    meter_serial_number_list=meter_serial_number_list_in.copy()

    meter_soft_list=meter_soft_list_in.copy()    

    meter_tech_number_list_old=meter_tech_number_list.copy()

    meter_status_test_list_old=meter_status_test_list.copy()

    meter_serial_number_list_old=meter_serial_number_list.copy()

    res=readGonfigValue("mass_config.json",[],{}, workmode, "1")
    if res[0]!="1":
        a_err_txt="Не удалось прочитать конфигурационные данные."
        return ["0",a_err_txt]

    mass_config_dic=res[2]

    time_interval_limit=mass_config_dic.get("time_interval_limit",10)

    mass_log_split_print=mass_config_dic.get("mass_log_split_print","индикатор")

    mass_log_print_analysis=mass_config_dic.get("mass_log_print_analysis","индикатор")
    
    time_interval=mass_config_dic.get("time_interval",1)

    window_title="MassProdAutoConfigGUI"

    moveResizeWindow([window_title], -1, 10, -1, 663)

    log_file_folder=os.path.split(file_log_path)[0]

    mass_log_line_dic={}

    count_meter_stage_ok=0

    meter_number_analisys_err_list=[]
    
    log_analiz_mode="2"
    
    stage=0
    while stage<2:

        if stage==0:
            mass_log_line_dic={}
            for meter_tech_number in meter_tech_number_list_old:
                mass_log_line_dic[meter_tech_number]={
                    "err_in_log_list":[],
                    "except_in_log_list":[],
                    "no_substrings_found_list":[],
                    "analisys_res_0":"",
                    "analisys_res_1":"",
                    "log_line_file_list":[]
                }
           
        for i in range(0, mass_number_of_meter):
            if meter_tech_number_list[i]!=None and meter_tech_number_list[i]!="":
                log_meter_path=os.path.join(log_file_folder, 
                    f"log_{i}.txt")
                with open(log_meter_path, "w", errors="ignore", encoding='utf-8') as file:
                    pass
        
        count_meter_stage_ok=0

        meter_number_analisys_err_list=[]
        
        err_in_log_list=[]

        except_in_log_list=[]

        log_line_file_list=[]

        no_substrings_found_list=[]

        control_val_dic=mass_config_dic["stages_dic"][str(stage)]

        stage_name = control_val_dic["stage_name"]

        if stage==0:
            with open(file_log_path, "w", errors="ignore") as file:
                pass

        send_start_command_list=control_val_dic["send_start_command_list"]

        send_stop_command_list=control_val_dic["send_stop_command_list"]

        if log_analiz_mode=="2":
            res=sendCommandShortKey(window_title, send_start_command_list, 
                file_log_path, stage_name, "1", workmode)
        

            if res[0] in ["0","2"]:
                printWARNING("Не удалось запустить процесс проверки в "
                    "автоматическом режиме.")
                res=massSelectActions("1")
                if res[0]=="1" and res[2]=="6":
                    continue

                if res[0]=="9":
                    return ["9", res[1]]
                
                a_txt=f"{bcolors.OKGREEN}Проведите проверку " \
                    "конфигурации ПУ вручную.\n" \
                    f"{bcolors.OKBLUE}По окончании - " \
                    "нажмите Enter."
                log_analiz_mode="1"

                questionSpecifiedKey("", a_txt, ["\r"], "", 1)

            if log_analiz_mode=="1":
                a_txt="Провести анализ log-файла? 0-нет, 1-да:"
                oo=questionSpecifiedKey(bcolors.OKBLUE, a_txt, 
                    ["0", "1"], "", 1)
                print()
                log_analiz_mode=oo

            if log_analiz_mode=="0":
                return ["2", "Анализ log-файла проведен не был."]
    
        if log_analiz_mode=="2":
            keyboard.on_press(onPress)

        res_read_multi=onlineReadMulti(file_log_path, control_val_dic,
            mass_number_of_meter, meter_tech_number_list, stage, time_interval, 
            time_interval_limit, mass_log_split_print,
            print_msg_err)
            
        if log_analiz_mode=="2":
            keyboard.unhook_all()

        if len(send_stop_command_list)>0 and log_analiz_mode=="2":
            sendCommandShortKey(window_title, send_stop_command_list, 
                file_log_path, stage_name, "0", workmode)
            
        if res_read_multi[0]=="0":
            menu_item_add_list=["Пропустить проверку конфигурации ПУ"]
            menu_id_add_list=["пропустить тест"]
            res = massSelectActions("3", menu_item_add_list, menu_id_add_list)
            if res[0]=="9":
                return ["9", res[1]]
            
            if res[0]=="1":
                if res[2] == "6":
                    stage=0
                    continue

                if res[2] in ["3", "4"]:
                    return [res[2], res[1]]

            elif res[0]=="2":
                if res[2]=="пропустить тест":
                    return ["8", "Пропустить тест."]
                
        elif res_read_multi[0]=="8":
            return ["8", res_read_multi[1]]

        elif res_read_multi[0]=="9":
            return ["9", res_read_multi[1]]
        
        elif res_read_multi[0] in ["3", "4"]:
            return [res_read_multi[0], res_read_multi[1]] 
        
        elif res_read_multi[0] == "5":
            stage=0
            continue         


        for i in range(0, len(meter_tech_number_list)):
            meter_tech_number=meter_tech_number_list[i]
            meter_serial_number=meter_serial_number_list[i]
            meter_serial_space=meter_serial_number[0:-7]+" " \
                +meter_serial_number[-7:len(meter_serial_number)]

            if meter_tech_number==None or meter_tech_number=="":
                continue

            meter_soft=meter_soft_list[i]

            param_exceptions_list=[]
        
            res = getParamExceptionsList(meter_soft, workmode)
            if res[0]!="1":

                a_err_txt="Ошибка при получении списка исключений для " \
                    f"версии ПО {meter_soft} для ПУ № {meter_tech_number} " \
                    f"({meter_serial_space})."
                printFAIL(a_err_txt)
                return ["0", a_err_txt]

            param_exceptions_list=res[2]

            log_meter_path=os.path.join(log_file_folder, 
                f"log_{i}.txt")
            if not os.path.exists(log_meter_path):
                a_err_txt=f"Отсутствует log-файл {log_meter_path} " \
                    f"для ПУ № {meter_tech_number} " \
                    f"({meter_serial_space})."
                printFAIL(a_err_txt)
                return ["0", a_err_txt]
            
             
            res=meterLogAnalysis(log_meter_path, control_val_dic, 
                param_exceptions_list, "1", mass_log_print_analysis)
            
            err_in_log_list=res[2]

            except_in_log_list=res[3]

            log_line_file_list=res[4]

            no_substrings_found_list=res[5]
            
            if res[0]=="1":
                printGREEN(f"ПУ № {meter_tech_number} ({meter_serial_space}) " \
                    f"успешно прошел этап '{stage_name}'.")
                
                count_meter_stage_ok+=1

                a_mass_dic=mass_log_line_dic.get(meter_tech_number,{})
                a_mass_dic["err_in_log_list"]=err_in_log_list
                a_mass_dic["except_in_log_list"]=except_in_log_list
                a_mass_dic["no_substrings_found_list"]=no_substrings_found_list
                a_mass_dic[f"analisys_res_{str(stage)}"]="ok"
                a_mass_dic["log_line_file_list"]=log_line_file_list
                mass_log_line_dic[meter_tech_number]=a_mass_dic.copy()
                
                continue

            elif res[0]=="0":
                a_txt=f"При прохождении этапа '{stage_name}' " \
                    f"для ПУ № {meter_tech_number} ({meter_serial_space}) " \
                    f"была выявлена ошибка в работе программы."
                printWARNING(a_txt)
                a_mass_dic=mass_log_line_dic.get(meter_tech_number,{})
                a_mass_dic["err_in_log_list"]=err_in_log_list
                a_mass_dic["except_in_log_list"]=except_in_log_list
                a_mass_dic["no_substrings_found_list"]=no_substrings_found_list
                a_mass_dic[f"analisys_res_{str(stage)}"]="error"
                a_mass_dic["log_line_file_list"]=log_line_file_list
                mass_log_line_dic[meter_tech_number]=a_mass_dic.copy()
                
                meter_number_analisys_err_list.append(meter_serial_number)

            elif res[0]=="2":
                meter_number_analisys_err_list.append(meter_serial_number)

                a_defects="\n".join(err_in_log_list)
                a_txt=f"При прохождении этапа '{stage_name}' " \
                    f"для ПУ № {meter_tech_number} ({meter_serial_space}) " \
                    f"были выявлены замечания:\n{a_defects}"
                if len(no_substrings_found_list)>0:
                    a_sub="\n".join(no_substrings_found_list)
                    if len(err_in_log_list)>0:
                        a_txt=f"{a_txt}\nТакже в log-файле не найдены " \
                            f"следующие подстроки:\n{a_sub}"
                    
                    else:
                        a_txt=f"В log-файле не найдены " \
                            f"следующие подстроки:\n{a_sub}"
                
                printFAIL(f"\n{a_txt}")
                if stage==1:
                    a_mass_dic=mass_log_line_dic.get(meter_tech_number,{})
                    a_mass_dic["err_in_log_list"]=err_in_log_list
                    a_mass_dic["except_in_log_list"]=except_in_log_list
                    a_mass_dic["no_substrings_found_list"]=no_substrings_found_list
                    a_mass_dic[f"analisys_res_{str(stage)}"]="bad"
                    a_mass_dic["log_line_file_list"]=log_line_file_list
                    mass_log_line_dic[meter_tech_number]=a_mass_dic.copy()
                        
        
        if stage==0 and count_meter_stage_ok!=mass_number_of_meter:
            a_str=", ".join(meter_number_analisys_err_list)
            a_txt=f"При прохождении этапа '{stage_name}' " \
                f"у ПУ № {a_str} были выявлены замечания."            
            printWARNING(a_txt)

        
            menu_item_add_list=["Пропустить проверку конфигурации ПУ"]
            menu_id_add_list=["пропустить тест всех ПУ"]

            if mass_number_of_meter>1 and count_meter_stage_ok==0:
                menu_item_add_list=["Пропустить проверку конфигурации всех ПУ"]
                menu_id_add_list=["пропустить тест всех ПУ"]

            elif mass_number_of_meter>1 and count_meter_stage_ok!=0:
                menu_item_add_list=["Пропустить проверку конфигурации у ПУ с ошибкой", 
                    "Пропустить проверку конфигурации всех ПУ"]
                menu_id_add_list=["пропустить тест ПУ","пропустить тест всех ПУ"]
            res = massSelectActions("3", menu_item_add_list, menu_id_add_list)
            if res[0]=="9":
                return ["9", res[1]]
            
            if res[0]=="1":
                if res[2] == "6":
                    meter_tech_number_list=meter_tech_number_list_old.copy()
                    meter_status_test_list=meter_status_test_list_old.copy()
                    meter_serial_number_list=meter_serial_number_list_old.copy()
                    stage=0
                    continue

                if res[2] in ["3", "4"]:
                    return [res[2], res[1]]

            elif res[0]=="2":
                if res[2]=="пропустить тест всех ПУ":
                    return ["8", "Пропустить тест."]
                
                elif res[2]=="пропустить тест ПУ":
                    a_str=", ".join(meter_number_analisys_err_list)
                    a_txt=f"Для ПУ № {a_str} будет пропущена дальнейшая " \
                        "проверка конфигурации."
                    printWARNING(a_txt)
                    for i in range (0, len(meter_number_analisys_err_list)):
                        ind=meter_serial_number_list.index(meter_number_analisys_err_list[i])
                        meter_tech_number_list[ind]=""
                        meter_status_test_list[ind]="пропущен"
                        meter_serial_number_list[ind]=""

        stage+=1
    

    return ["1", "Анализ log-файла проведен.", mass_log_line_dic]



def cycleOnlineReadLog(file_log_path: str, meter_tech_number_list_in: list,
    meter_status_test_list_in: list, meter_serial_number_list_in: list, 
    mass_number_of_meter: int, meter_soft_list_in: list,
    workmode="эксплуатация", print_msg_err="1"):


    global programm_status 
  
    meter_tech_number_list=listCopy(meter_tech_number_list_in)

    meter_status_test_list=listCopy(meter_status_test_list_in)

    meter_serial_number_list=listCopy(meter_serial_number_list_in)

    meter_soft_list=listCopy(meter_soft_list_in)

    meter_tech_number_list_old=listCopy(meter_tech_number_list)

    meter_status_test_list_old=listCopy(meter_status_test_list)

    meter_serial_number_list_old=listCopy(meter_serial_number_list)

    res=readGonfigValue("mass_config.json",[],{}, workmode, "1")
    if res[0]!="1":
        a_err_txt="Не удалось прочитать конфигурационные данные."
        return ["0",a_err_txt]

    mass_config_dic=res[2]

    time_interval_limit=mass_config_dic.get("time_interval_limit",10)

    mass_log_split_print=mass_config_dic.get("mass_log_split_print","индикатор")

    mass_log_print_analysis=mass_config_dic.get("mass_log_print_analysis","индикатор")
    
    time_interval=mass_config_dic.get("time_interval",1)

    window_title="MassProdAutoConfigGUI"

    moveResizeWindow([window_title], -1, 10, -1, 663)

    log_file_folder=os.path.split(file_log_path)[0]

    mass_log_line_dic={}

    count_meter_stage_ok=0

    meter_number_analisys_err_list=[]
    
    log_analiz_mode="2"

    skip_send_short_key="1"
    
    stage=0
    while stage<2:

        if stage==0:
            mass_log_line_dic={}
            for meter_tech_number in meter_tech_number_list_old:
                mass_log_line_dic[meter_tech_number]={
                    "err_in_log_list":[],
                    "except_in_log_list":[],
                    "no_substrings_found_list":[],
                    "analisys_res_0":"",
                    "analisys_res_1":"",
                    "log_line_file_list":[]
                }
           
        for i in range(0, mass_number_of_meter):
            if meter_tech_number_list[i]!=None and meter_tech_number_list[i]!="":
                log_meter_path=os.path.join(log_file_folder, 
                    f"log_{i}.txt")
                with open(log_meter_path, "w", errors="ignore", encoding='utf-8') as file:
                    pass
        
        count_meter_stage_ok=0

        meter_number_analisys_err_list=[]
        
        err_in_log_list=[]

        except_in_log_list=[]

        log_line_file_list=[]

        no_substrings_found_list=[]

        control_val_dic=mass_config_dic["stages_dic"][str(stage)]

        stage_name = control_val_dic["stage_name"]

        if stage==0:
            with open(file_log_path, "w", errors="ignore") as file:
                pass

        send_start_command_list=control_val_dic["send_start_command_list"]

        send_stop_command_list=control_val_dic["send_stop_command_list"]

        if log_analiz_mode=="2":
            if skip_send_short_key=="0":
                res=sendCommandShortKey(window_title, send_start_command_list, 
                    file_log_path, stage_name, "1", workmode)
        

                if res[0] in ["0","2"]:
                    printWARNING("Не удалось запустить процесс проверки в "
                        "автоматическом режиме.")
                    res=massSelectActions("1")
                    if res[0]=="1" and res[2]=="6":
                        closeProgram('MassProdAutoConfig.exe')
                        return ["4", "Перезапустить программу 'MassProdAutoConfig.exe'"]

                    if res[0]=="9":
                        return ["9", res[1]]
                    
                    a_txt=f"{bcolors.OKGREEN}Проведите проверку " \
                        "конфигурации ПУ вручную.\n" \
                        f"{bcolors.OKBLUE}По окончании - " \
                        "нажмите Enter."
                    log_analiz_mode="1"

                    questionSpecifiedKey("", a_txt, ["\r"], "", 1)

            if log_analiz_mode=="1":
                a_txt="Провести анализ log-файла? 0-нет, 1-да:"
                oo=questionSpecifiedKey(bcolors.OKBLUE, a_txt, 
                    ["0", "1"], "", 1)
                print()
                log_analiz_mode=oo

            if log_analiz_mode=="0":
                return ["2", "Анализ log-файла проведен не был."]
    
        if log_analiz_mode=="2":
            keyboard.on_press(onPress)

        res_read_multi=onlineReadMulti(file_log_path, control_val_dic,
            mass_number_of_meter, meter_tech_number_list, stage, time_interval, 
            time_interval_limit, mass_log_split_print,
            print_msg_err)
            
        if log_analiz_mode=="2":
            keyboard.unhook_all()

        if len(send_stop_command_list)>0 and log_analiz_mode=="2" and \
            skip_send_short_key=="0":
            sendCommandShortKey(window_title, send_stop_command_list, 
                file_log_path, stage_name, "0", workmode)
            
        if res_read_multi[0]=="0":
            a_mode="3"
            res = searchTitleWindow(window_title)
            if res[0] in ["0", "2"]:
                a_mode="5"

            menu_item_add_list=["Пропустить проверку конфигурации ПУ"]
            menu_id_add_list=["пропустить тест"]
            res = massSelectActions(a_mode, menu_item_add_list, 
                menu_id_add_list)
            if res[0]=="9":
                return ["9", res[1]]
            
            if res[0]=="1":
                if res[2] == "6":
                    skip_send_short_key="0"

                    stage=0
                    continue

                if res[2] in ["3", "4"]:
                    return [res[2], res[1]]

            elif res[0]=="2":
                if res[2]=="пропустить тест":
                    return ["8", "Пропустить тест."]
                
        elif res_read_multi[0]=="8":
            return ["8", res_read_multi[1]]

        elif res_read_multi[0]=="9":
            return ["9", res_read_multi[1]]
        
        elif res_read_multi[0] in ["3", "4"]:
            return [res_read_multi[0], res_read_multi[1]] 
        
        elif res_read_multi[0] == "5":
            skip_send_short_key="0"

            stage=0
            continue         


        for i in range(0, len(meter_tech_number_list)):
            meter_tech_number=meter_tech_number_list[i]
            meter_serial_number=meter_serial_number_list[i]
            meter_serial_space=meter_serial_number[0:-7]+" " \
                +meter_serial_number[-7:len(meter_serial_number)]

            if meter_tech_number==None or meter_tech_number=="":
                continue

            meter_soft=meter_soft_list[i]

            param_exceptions_list=[]
        
            res = getParamExceptionsList(meter_soft, workmode)
            if res[0]!="1":

                a_err_txt="Ошибка при получении списка исключений для " \
                    f"версии ПО {meter_soft} для ПУ № {meter_tech_number} " \
                    f"({meter_serial_space})."
                printFAIL(a_err_txt)
                return ["0", a_err_txt]

            param_exceptions_list=res[2]

            log_meter_path=os.path.join(log_file_folder, 
                f"log_{i}.txt")
            if not os.path.exists(log_meter_path):
                a_err_txt=f"Отсутствует log-файл {log_meter_path} " \
                    f"для ПУ № {meter_tech_number} " \
                    f"({meter_serial_space})."
                printFAIL(a_err_txt)
                return ["0", a_err_txt]
            
             
            res=meterLogAnalysis(log_meter_path, control_val_dic, 
                param_exceptions_list, "1", mass_log_print_analysis)
            
            err_in_log_list=res[2]

            except_in_log_list=res[3]

            log_line_file_list=res[4]

            no_substrings_found_list=res[5]
            
            if res[0]=="1":
                printGREEN(f"ПУ № {meter_tech_number} ({meter_serial_space}) " \
                    f"успешно прошел этап '{stage_name}'.")
                
                count_meter_stage_ok+=1

                a_mass_dic=mass_log_line_dic.get(meter_tech_number,{})
                a_mass_dic["err_in_log_list"]=err_in_log_list
                a_mass_dic["except_in_log_list"]=except_in_log_list
                a_mass_dic["no_substrings_found_list"]=no_substrings_found_list
                a_mass_dic[f"analisys_res_{str(stage)}"]="ok"
                a_mass_dic["log_line_file_list"]=log_line_file_list
                mass_log_line_dic[meter_tech_number]=a_mass_dic.copy()
                
                continue

            elif res[0]=="0":
                a_txt=f"При прохождении этапа '{stage_name}' " \
                    f"для ПУ № {meter_tech_number} ({meter_serial_space}) " \
                    f"была выявлена ошибка в работе программы."
                printWARNING(a_txt)
                a_mass_dic=mass_log_line_dic.get(meter_tech_number,{})
                a_mass_dic["err_in_log_list"]=err_in_log_list
                a_mass_dic["except_in_log_list"]=except_in_log_list
                a_mass_dic["no_substrings_found_list"]=no_substrings_found_list
                a_mass_dic[f"analisys_res_{str(stage)}"]="error"
                a_mass_dic["log_line_file_list"]=log_line_file_list
                mass_log_line_dic[meter_tech_number]=a_mass_dic.copy()
                
                meter_number_analisys_err_list.append(meter_serial_number)

            elif res[0]=="2":
                meter_number_analisys_err_list.append(meter_serial_number)

                a_defects="\n".join(err_in_log_list)
                a_txt=f"При прохождении этапа '{stage_name}' " \
                    f"для ПУ № {meter_tech_number} ({meter_serial_space}) " \
                    f"были выявлены замечания:\n{a_defects}"
                if len(no_substrings_found_list)>0:
                    a_sub="\n".join(no_substrings_found_list)
                    if len(err_in_log_list)>0:
                        a_txt=f"{a_txt}\nТакже в log-файле не найдены " \
                            f"следующие подстроки:\n{a_sub}"
                    
                    else:
                        a_txt=f"В log-файле не найдены " \
                            f"следующие подстроки:\n{a_sub}"
                
                printFAIL(f"\n{a_txt}")
                if stage==1:
                    a_mass_dic=mass_log_line_dic.get(meter_tech_number,{})
                    a_mass_dic["err_in_log_list"]=err_in_log_list
                    a_mass_dic["except_in_log_list"]=except_in_log_list
                    a_mass_dic["no_substrings_found_list"]=no_substrings_found_list
                    a_mass_dic[f"analisys_res_{str(stage)}"]="bad"
                    a_mass_dic["log_line_file_list"]=log_line_file_list
                    mass_log_line_dic[meter_tech_number]=a_mass_dic.copy()
                        
        
        if stage==0 and count_meter_stage_ok!=mass_number_of_meter:
            a_str=", ".join(meter_number_analisys_err_list)
            a_txt=f"При прохождении этапа '{stage_name}' " \
                f"у ПУ № {a_str} были выявлены замечания."            
            printWARNING(a_txt)

            closeProgram('MassProdAutoConfig.exe')


            a_mode="5"

            menu_item_add_list=["Пропустить проверку конфигурации ПУ"]
            menu_id_add_list=["пропустить тест всех ПУ"]

            if mass_number_of_meter>1 and count_meter_stage_ok==0:
                menu_item_add_list=["Пропустить проверку конфигурации всех ПУ"]
                menu_id_add_list=["пропустить тест всех ПУ"]

            elif mass_number_of_meter>1 and count_meter_stage_ok!=0:
                menu_item_add_list=["Пропустить проверку конфигурации у ПУ с ошибкой", 
                    "Пропустить проверку конфигурации всех ПУ"]
                menu_id_add_list=["пропустить тест ПУ","пропустить тест всех ПУ"]
            res = massSelectActions(a_mode, menu_item_add_list, menu_id_add_list)
            if res[0]=="9":
                return ["9", res[1]]
            
            if res[0]=="1":
                if res[2] == "6":
                    meter_tech_number_list=meter_tech_number_list_old.copy()
                    meter_status_test_list=meter_status_test_list_old.copy()
                    meter_serial_number_list=meter_serial_number_list_old.copy()
                    stage=0
                    continue

                if res[2] in ["3", "4"]:
                    return [res[2], res[1]]

            elif res[0]=="2":
                if res[2]=="пропустить тест всех ПУ":
                    return ["8", "Пропустить тест."]
                
                elif res[2]=="пропустить тест ПУ":
                    a_str=", ".join(meter_number_analisys_err_list)
                    a_txt=f"Для ПУ № {a_str} будет пропущена дальнейшая " \
                        "проверка конфигурации."
                    printWARNING(a_txt)
                    for i in range (0, len(meter_number_analisys_err_list)):
                        ind=meter_serial_number_list.index(meter_number_analisys_err_list[i])
                        meter_tech_number_list[ind]=""
                        meter_status_test_list[ind]="пропущен"
                        meter_serial_number_list[ind]=""

        stage+=1
    

    return ["1", "Анализ log-файла проведен.", mass_log_line_dic]




def readLogFile(file_path: str, print_msg_err="1"):

    log_line_list=[]
    try:
        with open(file_path, 'r', errors="ignore", encoding='utf-8') as fp:
            for line in fp:
                line = line.rstrip('\n')
                log_line_list.append(line)
    except Exception as e:
        a_err_txt = f"При открытии файла '{file_path}' " \
            f"возникла ошибка: {e.args[0]}."
        printMsgWait(a_err_txt, bcolors.WARNING,print_msg_err)
            
        return ["0", a_err_txt, log_line_list]
    
    return ["1", "Файл успешно обработан.", log_line_list]



def prepareFileMask(meter_serial_number_list_in: list, print_msg_err="1",
    workmode="эксплуатация"):

   
    def innerPrepFolderMassProd():

        nonlocal mass_prod_vers

        _, ans2, mass_dir_path = getUserFilePath('MassProdAutoConfig.exe', 
            "1", workmode, "0")
        if mass_dir_path == "":
            a_err_txt="Ошибка при получении пути к папке с " \
                "ф.'MassProdAutoConfig.exe'"
            printFAIL(a_err_txt)
            return ["0", a_err_txt] 
    
        mass_dir_path_head=os.path.split(mass_dir_path)[0]
        
        file_ver_path=f"{mass_dir_path}\\mass_vers.json"

        if os.path.exists(file_ver_path):
            res=readGonfigValue("mass_vers.json")
            if res[0]=="0":
                a_err_txt="Ошибка при получении номера версии " \
                    f"программы 'MassProdAutoConfig.exe' из ф.'{file_ver_path}'."
                printFAIL(a_err_txt)
                return ["0", a_err_txt]

            a_ver=res[2].get("mass_vers",None)

            if a_ver==None:
                a_err_txt="Ошибка при получении номера версии " \
                    f"программы 'MassProdAutoConfig.exe' из ф.'{file_ver_path}'."
                printFAIL(a_err_txt)
                return ["0", a_err_txt] 
            
            if a_ver==mass_prod_vers:
                return ["1", "Папка 'MassConfiguration' подготовлена."]
            
            mass_dir_path_new=f"{mass_dir_path}_{a_ver}"
            try:
                os.rename(mass_dir_path, mass_dir_path_new)
            
            except Exception:
                a_err_txt=f"Ошибка при попытке перименовать папку {mass_dir_path} в " \
                    f"{mass_dir_path_new}'."
                printFAIL(a_err_txt)
                return ["0", a_err_txt]
            
        folder_name_dic={"13.2":"MassConfiguration_13.2",
            "14.1":"MassConfiguration_14.1"}
        
        folder_name_1=folder_name_dic.get(mass_prod_vers, None)

        if folder_name_1 == None:
            a_err_txt="Ошибка при формировании имени папки с " \
                f"ф.'MassProdAutoConfig.exe' с версией {mass_prod_vers}."
            printFAIL(a_err_txt)
            return ["0", a_err_txt]
        
        mass_dir_path_new=f"{mass_dir_path_head}\\{folder_name_1}"
        
        if not os.path.exists(mass_dir_path_new):
            a_err_txt=f"Отсутствует папка {mass_dir_path_new}."
            printFAIL(a_err_txt)
            return ["0", a_err_txt]
        
        try:
            os.rename(mass_dir_path_new, mass_dir_path)
        
        except Exception:
            a_err_txt=f"Ошибка при попытке перименовать папку {mass_dir_path_new} в " \
                f"{mass_dir_path}'."
            printFAIL(a_err_txt)
            return ["0", a_err_txt]
        
        return ["1", "Папка 'MassConfiguration' подготовлена."]
         
    
    res=readGonfigValue("mass_config.json",[],{}, workmode, "1")
    if res[0]!="1":
        a_err_txt="Не удалось прочитать конфигурационные данные " \
            "из файла 'mass_config.json'."
        return ["0",a_err_txt]
        
    mass_config_dic=res[2]
    
    mass_prod_vers=mass_config_dic.get("mass_prod_vers", None)
    if mass_prod_vers==None:
        a_err_txt="Не удалось получить версию программы MassProdAutoConfig.exe" \
            "из файла 'mass_config.json'."
        return ["0",a_err_txt]

    res=innerPrepFolderMassProd()
    if res[0]=="0":
        return ["0", res[1]]
    
    meter_serial_number_list=listCopy(meter_serial_number_list_in)

    for meter_serial_number in meter_serial_number_list:
        if meter_serial_number!=None and meter_serial_number!="":
            break

    else:
        return ["0", "Список серийных номеров ПУ пуст."]

    sheet_name = "Product1"
    res=toGetProductInfo2(meter_serial_number, sheet_name)
    if res[0]=="0":
        a_err_txt = "Ошибка при получении информации о ПУ из " \
            "ф.ProductNumber.xlsx."
        if print_msg_err == "1":
            print(f"{bcolors.WARNING}{a_err_txt}{bcolors.ENDC}")
        return ["0", a_err_txt]
 
    a_txt=res[28]
    if a_txt==None or a_txt=="":
        a_err_txt="В ф.ProductNumber.xlsx отсутствует информация о " \
            "файле для настройки программы проверки конфигурации ПУ" \
            '"MassProdAutoConfig.exe".'
        if print_msg_err == "1":
            printWARNING (a_err_txt)
        return ["0", a_err_txt]

    mask_standard_file_name=a_txt

    a_name="CnfObjsXLSxGXDLMSMasks.omask_dir"
    _, ans2, mask_file_dir_path = getUserFilePath(a_name, "1", 
        workmode=workmode)

    if mask_file_dir_path == "":
        a_err_txt = f"Ошибка в ПП getUserFilePath(): {ans2}"
        if print_msg_err == "1":
            print(f"{bcolors.WARNING}{a_err_txt}{bcolors.ENDC}")
        return ["0", a_err_txt]

    mask_file_path = os.path.join(
        mask_file_dir_path, "CnfObjsXLSxGXDLMSMasks.omask")

    mask_sha256=""

    res=checksumFile(mask_file_path, None, "0")
    if res[0]=="0":
        a_err_txt='При подсчете контрольной суммы ' \
            f'ф."CnfObjsXLSxGXDLMSMasks.omask" возникла ошибка: {res[1]}'
        if print_msg_err == "1":
            printWARNING (a_err_txt) 
        return ["0", res[1]]

    elif res[0]=="3":
        mask_sha256=res[2]

    mask_standard_file_path= f"{mask_file_dir_path}\\{mask_standard_file_name}"

    res=checksumFile(mask_standard_file_path, None, "0")
    if res[0] in ["0", "4"]:
        a_err_txt='При подсчете контрольной суммы эталонного mask-файла' \
            f'"{mask_standard_file_name}" возникла ошибка: {res[1]}'
        if print_msg_err == "1":
            printWARNING (a_err_txt)
        return ["0", res[1]]

    mask_standart_sha256=res[2]

    if mask_sha256!=mask_standart_sha256:
        try:
            shutil.copy2(mask_standard_file_path, mask_file_path)
        except Exception as e:
            a_err_txt = f"Ошибка при копировании файла " \
                f'{mask_standard_file_path}: "{e.strerror}".'
            if print_msg_err == "1":
                print(f"{bcolors.FAIL}{a_err_txt}{bcolors.ENDC}")
            return ["0", a_err_txt]

    printWARNING("Для проверки используется mask-файл " 
        f"'{mask_standard_file_name}'.")
    
    a_dic = {"mask_file_name": mask_standard_file_name}
    saveConfigValue("mass_log_line_multi.json", a_dic, workmode) 
    
    return ["1", "Файл успешно подготовлен."]



def prepareFileLog(employee_id: str, mass_number_of_meter: int,
    print_msg_err="1", workmode="эксплуатация"):
   

    
    res = toformatNow()
    dt=res[1]
    dtt=res[4]
    file_log_name_base = f"Log_1_{employee_id}_{dt[0:2]}-{dt[3:5]}-{dt[8:10]}"


    file_log_name = f"{file_log_name_base}.txt"
    file_log_sum_name=f"{file_log_name_base}_свод.txt"

    dir_name = "mass_config_log_dir"
    _, ans2, dir_path = getUserFilePath(dir_name, "1", workmode=workmode)
    if dir_path == "":
        a_err_txt = f"Ошибка в ПП getUserFilePath(): {ans2}"
        if print_msg_err == "1":
            print(f"{bcolors.WARNING}{a_err_txt}{bcolors.ENDC}")
        return ["0", a_err_txt, "", ""]


    file_log_dir_path=os.path.join(dir_path, employee_id)
    if not os.path.isdir(file_log_dir_path):
        res=createFolder(file_log_dir_path, "1")
        if res[0]=="0":
            return ["0", res[1]]
        

    file_log_sum_path = os.path.join(file_log_dir_path, file_log_sum_name)
    if not os.path.exists(file_log_sum_path):
        with open(file_log_sum_path, "w", errors="ignore") as file:
            pass 

    
    file_log_path = os.path.join(file_log_dir_path, file_log_name)
    if os.path.exists(file_log_path):
        with open(file_log_path, "r", errors="ignore",) as first, \
            open(file_log_sum_path, "a", errors="ignore") as second:
                data = first.read()
                second.write(data)
                

    with open(file_log_path, "w", errors="ignore") as file:
        pass
    
    
    for i in range(0, mass_number_of_meter):
        log_meter_path=os.path.join(file_log_dir_path, 
            f"log_{i}.txt")
        if os.path.exists(log_meter_path):
            os.remove(log_meter_path)

    return ["1", "Файл успешно подготовлен.", file_log_name, 
            file_log_path]



def getParamExceptionsList(meter_soft: str, workmode="эксплуатация"):

    param_exceptions_list = []

    res = readGonfigValue("mass_exception_list.json", [], {}, workmode, "1")
    if res[0] != "1":
        return ["0", "Ошибка при получении списка исключений из "
                "ф.mass_exception_list.json"]


    a_list = res[2]["param_exceptions"]
    for a_dic in a_list:
        if cmpVers(meter_soft, a_dic["meter_soft"]) == "<":
            param_exceptions_list.extend(a_dic["param_exceptions_list"])
    
    return["1", "Список с исключениями сформирован.", 
            param_exceptions_list]
        
    

def setProgramMassSetting(meter_tech_number_list_in: list,
    mass_number_of_meter: int, com_config_list_in: list, 
    mass_control_com_port="2", 
    print_msg_err="1", workmode="эксплуатация"):

    meter_tech_number_list=listCopy(meter_tech_number_list_in)
    com_config_list=listCopy(com_config_list_in)

    def innerRemoveExtraKey():

        reg_path_1="MassProdAutoConfigGUI"

        res = subKeysWinReg(reg_path_1)
        if res[0]=="0":
            a_err_txt="В реестре Windows не найдены ключи " \
                "программы 'MassProdAutoConfig.exe'."
            printMsgWait(a_err_txt, bcolors.WARNING,
                print_msg_err)
            return ["0", a_err_txt]

        subkeys_list=res[2]
        if len(subkeys_list)==0:
            a_err_txt="В реестре Windows не найдены ключи " \
                "программы 'MassProdAutoConfig.exe'."
            printMsgWait(a_err_txt, bcolors.WARNING,
                print_msg_err)
            return ["0", a_err_txt]

        for i in range(1, len(subkeys_list)):
            a_reg_name=""
            reg_name = subkeys_list[i]
            try:
                a_reg_name=str(int(reg_name))
            
            except Exception:
                pass

            if a_reg_name==reg_name:
                res = delKeyWinReg(reg_path_1, reg_name)
                if not res:
                    a_err_txt=f"Не удалось удалить ключ '{reg_name}' " \
                        "в реестре Windows у программы " \
                        "'MassProdAutoConfig.exe'."
                    printMsgWait(a_err_txt, bcolors.WARNING,
                        print_msg_err)
                    return ["0", a_err_txt]
            


                    
                

        return ["1", "Операция выполнена успешно.", subkeys_list[0]]

    
    def innerSettingComPortWinReg():
        
        nonlocal mass_number_of_meter
        
        a_com_limit=mass_number_of_meter
        
        if len(com_config_list)<mass_number_of_meter:
            a_com_limit=len(com_config_list)

        com_str=", ".join(com_config_list[0:a_com_limit])

        to_replace_com_port=True

        if mass_control_com_port=="1":
            com_config_winreg_list=[]

            reg_name="AssignedCOMPortsWithPlacements"
            res=getWinReg(REG_PATH,reg_name)
            if res!=None and res!="":
                a_list=json.loads(res)
                for i in range(0, len(a_list)):
                    com_config_winreg_list.append(a_list[i]["SerialPort"])

            a_com_win_str=", ".join(com_config_winreg_list[0:a_com_limit])

            txt1=""

            if len(com_config_winreg_list)==0:
                txt1=f"{bcolors.WARNING}В настройках программы " \
                    f"MassProdAutoConfig.exe не найдена информация об " \
                    f"используемых COM-портах.\n"
                
            else:
                if a_com_win_str!=com_str:
                    txt1=f"{bcolors.WARNING}В настройках программы " \
                        "MassProdAutoConfig.exe указано, что " \
                        f"используется {a_com_win_str}.\n"
                    
                    if len(com_config_winreg_list)>1:
                        txt1=f"{bcolors.WARNING}В настройках программы " \
                            "MassProdAutoConfig.exe указано, что " \
                            f"используются: {a_com_win_str}.\n"
                
            if txt1!="":
                if len(com_config_list)==1:
                    txt1=f"{txt1}{bcolors.WARNING}При этом значением по " \
                        f"умолчанию является - {com_str}.\n"
                    
                else:
                    txt1=f"{txt1}{bcolors.WARNING}При этом значениями по " \
                        f"умолчанию являются: {com_str}.\n"
                    
                txt1=f"{txt1}{bcolors.OKBLUE}Выберите дальнейшее действие:"
                item_list=["Заменить значение в настройках программы " 
                        "MassProdAutoConfig.exe",
                        "Оставить значение без изменений"]
                item_id=["заменить", "оставить"]
                spec_list=["Прервать проверку"]
                spec_keys=["/"]
                spec_id=["прервать"]
                oo=questionFromList(bcolors.OKBLUE, txt1, item_list, item_id, 
                "", spec_list, spec_keys, spec_id, 1, 1, 1, [], "")
                print()
                if oo=="прервать":
                    return ["9", "Операция прервана пользователем."]
                
                elif oo=="оставить":
                    to_replace_com_port=False

                printGREEN("Будут внесены изменения в настройки программы "
                    "MassProdAutoConfig.exe.")


        if to_replace_com_port:
            a_reg_list=[]
            mass_index=0
            for i in range(0, len(meter_tech_number_list)):
                if meter_tech_number_list[i]!=None and \
                    meter_tech_number_list[i]!="" and \
                    com_config_list[i]!=None and com_config_list[i]!="":
                    a_reg_list.append({"Placement":mass_index,"SerialPort":com_config_list[i]})
                    com_last=com_config_list[i]
                    mass_index+=1
            
            for i in range(len(a_reg_list), 11):
                a_reg_list.append({"Placement":i,"SerialPort":com_last})

            reg_value=json.dumps(a_reg_list, ensure_ascii=False)
            reg_name="AssignedCOMPortsWithPlacements"
            setWinreg(REG_PATH, reg_name, reg_value)

        return ["1", "Настрока COM-порта выполнена успешно."]

    
    res=innerRemoveExtraKey()
    if res[0]=="0":
        return ["0", res[1]]
    
    
    REG_PATH = f"MassProdAutoConfigGUI\\{res[2]}"

        
 
    
    reg_name="NumberPlacementsSPhase"
    setWinreg(REG_PATH, reg_name, "11")


    res=innerSettingComPortWinReg()
    if res[0]!="1":
        return [res[0], res[1]]  
     
    return ["1", "Настройка работы программы выполнена успешно."]



def massSelectActions(select_mode="1", menu_item_add_list=[], 
                      menu_id_add_list=[]):

    window_title="Проверка конфигурации ПУ с помощью программы " \
     "'MassProdAutoConfig'"
        
    res=actionsSelectedtWindow([window_title], None,
        "активировать","1")
    
    action_dic={"1": ["Повторить попытку", 
        "Провести проверку конфигурации ПУ вручную"],
                "2": ["Продолжить анализ записей в журнале",
        "Перезапустить процесс проверки без выхода из "
         "программы", "Перезапустить программу "
         "'MassProdAutoConfig.exe'", "Провести проверку "
         "конфигурации ПУ вручную"],
         "3": ["Перезапустить процесс проверки без выхода из "
         "программы", "Перезапустить программу "
         "'MassProdAutoConfig.exe'", "Провести проверку "
         "конфигурации ПУ вручную"],
         "4": ["Продолжить анализ записей в журнале",
            "Провести проверку конфигурации ПУ вручную"],
         "5": ["Перезапустить программу 'MassProdAutoConfig.exe'", 
               "Провести проверку конфигурации ПУ вручную"]}
    action_id_dic = {"1": ["6", "3"], 
                     "2": ["5", "6", "4", "3"],
                     "3": ["6", "4", "3"],
                     "4": ["5", "3"],
                     "5": ["4", "3"],}
    menu_item_list=action_dic[select_mode]
    menu_id_list=action_id_dic[select_mode]
    header="Выберите дальнейшее действие:"
    res=menuSelectActions(menu_item_list, menu_id_list, header,
        menu_item_add_list, menu_id_add_list,)
    if res[0]=="9":
        printWARNING ("Проверка прервана.")
        return ["9", "Проверка прервана пользователем."]
    
    else:
        return [res[0], res[1], res[2]]



def onPress(key):
    global programm_status 
    
    foregroundWindow = GetWindowText(GetForegroundWindow())
  
    if ("Проверка конфигурации ПУ" in foregroundWindow) or \
        ("MassProdAutoConfigGUI" in foregroundWindow) or \
        ("otk_menu" in foregroundWindow):
        if key.name=="/":
            programm_status = "команда 'Прервать'"


    return
    


def onPressAnyKey(key):
    global programm_status 
    
    foregroundWindow = GetWindowText(GetForegroundWindow())
  
    if ("Проверка конфигурации ПУ" in foregroundWindow) or \
        ("MassProdAutoConfigGUI" in foregroundWindow) or \
        ("otk_menu" in foregroundWindow):

        a_dic={"0":"0", "1":"1"}
        if key.name in a_dic:
            programm_status = f"нажали клавишу "+a_dic[key.name]
        


    return



def onlineReadPrintOLD(file_log_path: str, control_val_dic_in: dict,
    log_analiz_mode="2", time_interval=1, 
    time_interval_limit=10, param_exceptions_list=[],
    print_msg_err="1",):

    ret_pp_descript_dic={"1": "Этап успешно пройден.",
        "2": "Этап пройден и выявлены ошибки.",
        "3": "Провести проверку конфигурации ПУ вручную",
        "4": "Перезапустить программу 'MassProdAutoConfig.exe'",
        "5": "Перезапустить процесс проверки без выхода "
             "из программы",
        "8": "Пропустить тест конфигурации ПУ.",
        "9": "Пользователь прервал проверку ПУ"}

    global programm_status      #имя статуса, в котором 
    
    control_val_dic=control_val_dic_in.copy()

    bloc_start=control_val_dic["blok_marks"]["start"]
    bloc_end=control_val_dic["blok_marks"]["end"]

    warning_list=control_val_dic["colors_marks"]["warning"]
    green_list=control_val_dic["colors_marks"]["green"]

    status_true_list=control_val_dic["status"]["true"]
    status_false_list=control_val_dic["status"]["false"]
    status_cancel_list=control_val_dic["status"]["cancel"]

    err_in_log_list=[]
    
    except_in_log_list = []

    log_file_size=0
    log_line_file_list=[]
    log_line_cur_list=[]
    
    line_num_prev=0

    stage_ok=True
    notification_on=True
    
    programm_status="цикл чтения данных"

    time_sec_start=toformatNow()[3]

    txt="\nЗапускаю анализ содержимого log-файла программы " \
        "'MassProdAutoConfig.exe'.\n" \
        f"{bcolors.WARNING}Если программа 'зависнет', " \
        "то нажмите '/'."
    printGREEN(txt)
    
    while programm_status=="цикл чтения данных" or \
            programm_status == "команда 'Прервать'":

        if programm_status=="цикл чтения данных":
            time_sec_cur=toformatNow()[3]
            if time_sec_cur-time_sec_start>time_interval:
                res=checkResizeFile(file_log_path, log_file_size)
                if res[0]=="0":
                    return ["0", res[1]]
                
                log_file_size_cur=res[2]

                if res[0]=="1" or log_analiz_mode=="1":
                    notification_on=True

                    log_file_size=log_file_size_cur

                    res=readLogFile(file_log_path, print_msg_err)
                    if res[0]=="0":
                        return ["0", res[1]]
                    
                    log_line_cur_list=res[2]
                    log_line_file_list=log_line_cur_list.copy()
                    line_num_cur=len(log_line_cur_list)
                    a_list=log_line_cur_list[line_num_prev:line_num_cur]

                    line_num_prev=line_num_cur

                    for log_line in a_list:
                        color_line=None
                        if findListInStr(log_line, status_false_list, 
                            "0", "0")[0]=="1":
                            if len(param_exceptions_list)>0 and \
                                findListInStr(log_line, param_exceptions_list, 
                                    "0", "0")[0]=="1":
                                except_in_log_list.append(log_line)
                                log_line=log_line+bcolors.WARNING+ \
                                    " (в списке исключений)"
                                color_line = bcolors.OKGREEN

                            else:
                                err_in_log_list.append(log_line)
                                color_line=bcolors.FAIL
                                stage_ok=False

                        res=findListInStr(log_line, warning_list, "0", "0")
                        if res[0]=="1" and color_line==None:
                            color_line=bcolors.WARNING

                        res=findListInStr(log_line, status_cancel_list, "0", "0")
                        if res[0]=="1":
                            color_line=bcolors.WARNING

                        
                        res=findListInStr(log_line, status_true_list, "0", "0")
                        if res[0]=="1" and color_line==None:
                            a_index=res[3]
                            del status_true_list[a_index]
                            color_line=bcolors.OKGREEN

                        res=findListInStr(log_line, green_list, "0", "0")
                        if res[0]=="1" and color_line==None:
                            color_line=bcolors.OKGREEN                        
                        
                        if color_line!=None:
                            log_line=color_line+log_line
                        printColor(log_line)

                        if log_line.find(bloc_end) != -1:
                            if stage_ok and len(status_true_list)>0:
                                stage_ok=False
        
                            programm_status="плановый выход"

                    time_sec_start=time_sec_cur

                else:
                    if time_sec_cur-time_sec_start>time_interval_limit and \
                            notification_on==True:
                        txt = "Размер log-файла без изменений " \
                            f"более {time_interval_limit} сек.\n" \
                            f"Чтобы остановить ожидание новых записей - " \
                            f"нажмите '/'."
                        printWARNING(txt)
                        notification_on=False
 
        else:
            select_mode="2"
            menu_item_add_list=["Пропустить проверку конфигурации ПУ"]
            menu_id_add_list=["пропустить тест"]
            if log_analiz_mode=="1":
                select_mode="4"
                menu_item_add_list=[]
                menu_id_add_list=[]
            res = massSelectActions(select_mode, menu_item_add_list, 
                menu_id_add_list)
            if res[0]=="9":
                return ["9", ret_pp_descript_dic["9"]]
            
            if res[0]=="1":
                if res[2] == "5":
                    programm_status="цикл чтения данных"
                    notification_on=True
                    time_sec_start=time_sec_cur
                    printGREEN("Продолжаю анализ log-файла...")
                    continue

                if res[2] in ["3", "4", "6"]:

                    a_dic={"3": "3", "4": "4", "6": "5"}
                    return [a_dic[res[2]], 
                            ret_pp_descript_dic[a_dic[res[2]]]]            
            
            elif res[0]=="2":
                if res[0]=="пропустить тест":
                    return ["8", "Пропустить тест."]
            
            a_txt="Вернулся неизвестный код.\n" \
                f"{bcolors.OKBLUE}Нажмите '/'."
            questionSpecifiedKey(bcolors.WARNING, a_txt,["/"],
                "", 1) 


    
    if programm_status=="плановый выход":
        if stage_ok:
            return ["1", ret_pp_descript_dic["1"], 
                    err_in_log_list, except_in_log_list,
                    log_line_file_list, status_true_list]
        
        else:
            return ["2", ret_pp_descript_dic["2"],
                err_in_log_list, except_in_log_list,
                log_line_file_list, status_true_list]
    
                

def meterLogAnalysis(file_log_path: str, control_val_dic_in: dict,
    param_exceptions_list=[], print_msg_err="1",
    mass_log_print_analysis="индикатор"):

    ret_pp_descript_dic={"1": "Этап успешно пройден.",
        "2": "Этап пройден и выявлены ошибки."}

    global programm_status      #имя статуса, в котором 
    
    a_control_val_dic=dictCopy(control_val_dic_in)


    bloc_start=a_control_val_dic.get("blok_marks", {}).get("start","" )
    bloc_end=a_control_val_dic.get("blok_marks", {}).get("end", "")

    warning_list=a_control_val_dic.get("colors_marks", {}).get("warning", [])
    green_list=a_control_val_dic.get("colors_marks", {}).get("green", [])

    status_true_list=a_control_val_dic.get("status", {}).get("true", [])

    status_false_list=a_control_val_dic.get("status", {}).get("false", [])
    status_cancel_list=a_control_val_dic.get("status", {}).get("cancel", [])

    log_file_name=os.path.split(file_log_path)[1]

    a_str=log_file_name[4:-4]
    meter_position_str=str(int(a_str)+1)

    for i in range(0,len(status_true_list)):
        a_status=status_true_list[i]
        if "#pos#" in a_status:
            a_status=a_status.replace("#pos#", meter_position_str)
            status_true_list[i]=a_status
    
 
    for i in range(0,len(param_exceptions_list)):
        a_param=param_exceptions_list[i]
        if "#pos#" in a_param:
            a_param=a_param.replace("#pos#", meter_position_str)
            param_exceptions_list[i]=a_param
    
    err_in_log_list=[]
    
    except_in_log_list = []

    log_line_file_list=[]

    stage_ok=True
 
    txt="\nАнализирую содержимое log-файла для ПУ, " \
        f"установленного на позиции {meter_position_str}."
    print(txt)
    
    res=readLogFile(file_log_path, print_msg_err)
    if res[0]=="0":
        return ["0", res[1]]
    
    log_line_file_list=res[2]

    if mass_log_print_analysis=="индикатор":
        tqdm_txt="Читаю информацию из log-файла ПУ"
        bar=tqdm(total=len(log_line_file_list), desc=tqdm_txt)
    
    
    for i_log_line in range(0, len(log_line_file_list)):
        log_line=log_line_file_list[i_log_line]
        color_line=None

        if mass_log_print_analysis=="индикатор":
            bar.update(i_log_line+1)

        if findListInStr(log_line, status_false_list, 
            "0", "0")[0]=="1":
            if len(param_exceptions_list)>0 and \
                findListInStr(log_line, param_exceptions_list, 
                    "0", "0")[0]=="1":
                except_in_log_list.append(log_line)
                log_line=log_line+bcolors.WARNING+ \
                    " (в списке исключений)"
                color_line = bcolors.OKGREEN

            else:
                err_in_log_list.append(log_line)
                color_line=bcolors.FAIL
                stage_ok=False

        res=findListInStr(log_line, warning_list, "0", "0")
        if res[0]=="1" and color_line==None:
            color_line=bcolors.WARNING

        res=findListInStr(log_line, status_cancel_list, "0", "0")
        if res[0]=="1":
            color_line=bcolors.WARNING

        res=findListInStr(log_line, status_true_list, "0", "0")
        if res[0]=="1" and color_line==None:
            a_index=res[3]
            del status_true_list[a_index]
            color_line=bcolors.OKGREEN

        res=findListInStr(log_line, green_list, "0", "0")
        if res[0]=="1" and color_line==None:
            color_line=bcolors.OKGREEN                        
        
        if color_line!=None:
            log_line=color_line+log_line
        
        if mass_log_print_analysis=="показать журнал":
            printColor(log_line)

        if log_line.find(bloc_end) != -1:
            if stage_ok and len(status_true_list)>0:
                stage_ok=False

            break

 
    if mass_log_print_analysis=="индикатор" :
        bar.close()

    if stage_ok:
        return ["1", ret_pp_descript_dic["1"], 
                err_in_log_list, except_in_log_list,
                log_line_file_list, status_true_list]
    
    else:
        return ["2", ret_pp_descript_dic["2"],
            err_in_log_list, except_in_log_list,
            log_line_file_list, status_true_list]
 



def onlineReadMulti(file_log_path: str, control_val_dic_in: dict,
    mass_number_of_meter: int, meter_tech_number_list_in: list,
    stage: int, time_interval=1, time_interval_limit=10, 
    mass_log_split_print="индикатор",
    print_msg_err="1",):
    
    meter_tech_number_list=listCopy(meter_tech_number_list_in)
    control_val_dic=dictCopy(control_val_dic_in)

    ret_pp_descript_dic={"1": "Этап успешно пройден.",
        "3": "Провести проверку конфигурации ПУ вручную",
        "4": "Перезапустить программу 'MassProdAutoConfig.exe'",
        "5": "Перезапустить процесс проверки без выхода "
             "из программы",
        "8": "Пропустить тест конфигурации ПУ.",
        "9": "Пользователь прервал проверку ПУ"}

    global programm_status      #имя статуса, в котором 
    
    window_title="MassProdAutoConfigGUI"

    bloc_start=control_val_dic["blok_marks"]["start"]
    bloc_end=control_val_dic["blok_marks"]["end"]

    status_cancel_list=control_val_dic["status"]["cancel"]

    log_file_size=0
    log_line_file_list=[]
    log_line_cur_list=[]
    
    log_file_folder=os.path.split(file_log_path)[0]

    line_num_prev=0

    log_meter_file_all_size=34816*mass_number_of_meter
    if stage==0:
        log_meter_file_all_size=3072*mass_number_of_meter

    notification_on=True
    
    programm_status="цикл чтения данных"

    time_sec_start=toformatNow()[3]

    txt="\nЗапускаю чтение журнала (log-файла) программы " \
        "'MassProdAutoConfig.exe'.\n" \
        f"{bcolors.WARNING}Если программа 'зависнет', " \
        "то нажмите '/'."
    printGREEN(txt)
    
    ctrl_str_list=["placeIndex - ", "позиция "]
    
    if mass_log_split_print=="индикатор" :
        tqdm_txt="Читаю информацию из log-файла программы " \
            "MassProdAutoConfig.exe"
        bar=tqdm(total=log_meter_file_all_size, desc=tqdm_txt)

    while programm_status=="цикл чтения данных" or \
            programm_status == "команда 'Прервать'":
        if programm_status=="цикл чтения данных":
            time_sec_cur=toformatNow()[3]
            if time_sec_cur-time_sec_start>time_interval:
                res=checkResizeFile(file_log_path, log_file_size)
                if res[0]=="0":
                    return ["0", res[1]]
                
                log_file_size_cur=res[2]

                if res[0]=="1":
                    notification_on=True

                    log_file_size=log_file_size_cur

                    if mass_log_split_print=="индикатор":
                        bar.update(log_file_size)

                    res=readLogFile(file_log_path, print_msg_err)
                    if res[0]=="0":
                        return ["0", res[1]]
                    
                    log_line_cur_list=res[2]
                    line_num_cur=len(log_line_cur_list)
                    a_list=log_line_cur_list[line_num_prev:line_num_cur]

                    line_num_prev=line_num_cur

                    for i_line in range(0, len(a_list)):
                        log_line=a_list[i_line]

                        meter_index=""
                        for ctrl_word in ctrl_str_list:
                            ctrl_pos=log_line.find(ctrl_word)
                            if ctrl_pos!=-1:
                                a_pos=ctrl_pos+len(ctrl_word)
                                for i in range(a_pos, len(log_line)):
                                    symbol=log_line[i]
                                    if not symbol.isnumeric():
                                        break

                                    meter_index=meter_index+symbol
                                
                                if meter_index!="":
                                    if meter_tech_number_list[int(meter_index)-1]!=None and \
                                        meter_tech_number_list[int(meter_index)-1]!="":
                                        log_meter_path=os.path.join(log_file_folder, 
                                            f"log_{str(int(meter_index)-1)}.txt")

                                        with open(log_meter_path, "a", errors="ignore", encoding='utf-8') as file:
                                            file.write(f"{log_line}\n")
                                    
                                    break

                        else:
                            for i in range(0, mass_number_of_meter):
                                if meter_tech_number_list[i]!=None and meter_tech_number_list[i]!="":
                                    log_meter_path=os.path.join(log_file_folder, 
                                        f"log_{i}.txt")
                                    with open(log_meter_path, "a", errors="ignore", encoding='utf-8') as file:
                                        file.write(f"{log_line}\n")
                                    
                            
                        if mass_log_split_print=="показать журнал":
                            printColor(log_line)

                        res=findListInStr(log_line, status_cancel_list, "0", "0")
                        if (res[0]=="1" and i_line==line_num_cur-1) or \
                            log_line.find(bloc_end) != -1:
                            
                            programm_status="плановый выход"

                    time_sec_start=time_sec_cur

                else:
                    if time_sec_cur-time_sec_start>time_interval_limit and \
                            notification_on==True:
                        txt = "Размер log-файла без изменений " \
                            f"более {time_interval_limit} сек.\n" \
                            f"Чтобы остановить ожидание новых записей - " \
                            f"нажмите '/'."
                        printWARNING(txt)
                        notification_on=False

        else:
            
            select_mode="2"
            res = searchTitleWindow(window_title)
            if res[0] in ["0", "2"]:
                select_mode="5"

            
            menu_item_add_list=["Пропустить проверку конфигурации ПУ"]
            menu_id_add_list=["пропустить тест"]
            if mass_log_split_print=="индикатор" :
                bar.close()

            
            res = massSelectActions(select_mode, menu_item_add_list, 
                menu_id_add_list)
            if res[0]=="9":
                return ["9", ret_pp_descript_dic["9"]]
        

            if res[0]=="1":
                if res[2] == "5":
                    if mass_log_split_print=="индикатор" :
                        tqdm_txt="Читаю информацию из log-файла программы " \
                            "MassProdAutoConfig.exe"
                        bar=tqdm(total=log_meter_file_all_size, desc=tqdm_txt)
                    
                    programm_status="цикл чтения данных"
                    notification_on=True
                    time_sec_start=time_sec_cur
                    printGREEN("Продолжаю анализ log-файла...")
                    continue

                if res[2] in ["3", "4", "6"]:

                    a_dic={"3": "3", "4": "4", "6": "5"}
                    return [a_dic[res[2]], 
                            ret_pp_descript_dic[a_dic[res[2]]]]            
            
            elif res[0]=="2":
                if res[2]=="пропустить тест":
                    return ["8", "Пропустить тест."]
            
            a_txt="Вернулся неизвестный код.\n" \
                f"{bcolors.OKBLUE}Нажмите '/'."
            questionSpecifiedKey(bcolors.WARNING, a_txt,["/"],
                "", 1)
    
    if programm_status=="плановый выход":
        if mass_log_split_print=="индикатор" :
            bar.close()

        return ["1", ret_pp_descript_dic["1"]]



def sendCommandShortKey(window_title: str, send_command_list_in: list,
    file_log_path: str, stage_name: str, send_mode: str,
    print_msg_err="1",  workmode="эксплуатация"):
    
    send_command_list=listCopy(send_command_list_in)
    
    res=readGonfigValue("mass_config.json",[],{}, workmode, "1")
    if res[0]!="1":
        a_err_txt="Не удалось прочитать конфигурационные данные."
        return ["0",a_err_txt]
        
    mass_config_dic=res[2]

    time_interval=mass_config_dic.get("time_interval",1)
    
    control_size=0

    if file_log_path!=None and file_log_path!="":
        res=checkResizeFile(file_log_path,0)
        if res[0]=="0":
            return ["0", res[1]]
        
        control_size=res[2]
    
    for attempt in range(0,2):
        res=actionsSelectedtWindow([window_title], None,
            "показать+активировать","1")
        if res[0]!="1":
            a_err_txt=f"Не удалось найти окно '{window_title}'."
            printMsgWait(a_err_txt, bcolors.WARNING, print_msg_err)
            return ["0", a_err_txt]
        
        res = searchImageWindow("", window_title, print_msg_err, 
            "1")
        if res[0]!="1":
            a_err_txt=f"Не удалось получить информацию об окне " \
                f"'{window_title}'."
            printMsgWait(a_err_txt, bcolors.WARNING, print_msg_err)
            return ["0", a_err_txt]

        window_x=res[2][0]
        window_y = res[2][1]
        window_w = res[5][0]
        window_h = res[5][1]

        mouse_x = window_x+window_w/2
        mouse_y = window_y+window_h/4

        for i in range(0,2):
            pause_ui(1)
            print ("Кликаю мышкой в окне...")
            res = mouseClickGraf(mouse_x, mouse_y, "1")
            if res[0] == "0":
                a_err_txt = f"Не удалось кликнуть мышкой в окне '{window_title}'."
                printMsgWait(a_err_txt, bcolors.WARNING, print_msg_err)
                return ["0", a_err_txt]
            
        if send_mode=="1":
            print(f"Отправляю команду запуска этапа '{stage_name}'.")
        
        else:
            print(f"Отправляю команду для процесса этапа '{stage_name}'.")

        for a_key in send_command_list:
            keyboard.send(a_key)
            time.sleep(0.2)
        
        if send_mode=="0":
            return ["1", "Команда успешно отправлена."]

        pause_ui(1.5)
        

        res=cycleCheckResizeFile(file_log_path, control_size, 
            time_interval, 5)
        if res[0]=="0":
            return ["0", res[1]]

        if (send_mode=="1" and res[0]=="1") or \
            (send_mode=="2" and res[0]=="2"):
            return ["1", "Команда успешно выполнена."]
    
    return ["2", "Команда не выполнена."]



def saveToMassConfigJSON(mass_responce: str, responce_dic={},
                         workmode="эксплуатация"):
    a_dic = {"mass_responce": mass_responce}
    a_dic.update(responce_dic)
    saveConfigValue("mass_config.json",a_dic, workmode)
    return



def exitProgram():

    filename = 'MassProdAutoConfig.exe'
    closeProgram(filename)
    sys.exit()



def actionBeforeClosingMass(analisys_res_ok: bool, 
    mass_config_action_result="1"):
    
    global programm_status

    time_sec_start=toformatNow()[3]
    
    if mass_config_action_result=="0":
        programm_status = "ожидаем нажатие клавиши"

        keyboard.on_press(onPressAnyKey)

        a_txt=f"Для закрытия окна программы " \
            "'MassProdAutoConfig.exe' и продолжения проверки ПУ " \
            "нажмите 0."
        
        time_sec_start=0

    if analisys_res_ok:
        return

    elif not analisys_res_ok:
        programm_status = "ожидаем нажатие клавиши"

        keyboard.on_press(onPressAnyKey)

        a_txt=f"{bcolors.WARNING}Выявлены замечания.\n{bcolors.ENDC}"

        if mass_config_action_result=="1":
            a_txt=a_txt+"Через 7 сек. окно с программой 'MassProdAutoConfig.exe' " \
            "будет закрыто.\n" \
            f"Чтобы отменить закрытие окна - нажмите '1'.\n" \
            f"Чтобы закрыть его немедленно и продолжить проверку - нажмите '0'."
        
        if mass_config_action_result=="2":
            a_txt=a_txt+f"Для закрытия окна программы " \
                "'MassProdAutoConfig.exe' и продолжения проверки ПУ " \
                "нажмите 0."
            
            time_sec_start=0
  
    
    printColor(a_txt, bcolors.OKBLUE) 
    while programm_status in ["ожидаем нажатие клавиши",
        "нажали клавишу 0", "нажали клавишу 1"]:
        if time_sec_start>0:
            time_sec_cur=toformatNow()[3]
            if time_sec_cur-time_sec_start>7:
                break

            elif programm_status=="нажали клавишу 1":
                window_title="getLogMassConfig"
                res=actionsSelectedtWindow([window_title], None,
                    "показать+активировать","1")
                
                printGREEN("Закрытие окна программы " 
                    "'MassProdAutoConfig.exe' отменено.")
                printBLUE("Для продолжения проверки ПУ " 
                    "вернитесь в текущее окно и нажмите '/'.")
                questionSpecifiedKey("", "", ["/"], "", 1)
                break

            elif programm_status=="нажали клавишу 0":
                break


        elif programm_status=="нажали клавишу 0":
            break


    keyboard.unhook_all()

    return



def replaceMyTitleWindows():
    title_new="Проверка конфигурации ПУ с помощью программы " \
        "'MassProdAutoConfig'"
    replaceTitleWindow("", title_new)

    return



def mainMassConfig():
    global programm_status


    
    print_result="1"

    n = len(sys.argv)
    if n>1:
        print_result=sys.argv[1]


    title_new="Проверка конфигурации ПУ с помощью программы " \
        "'MassProdAutoConfig'"
    res = replaceTitleWindow("", title_new)
    
    workmode="эксплуатация"
    res=readGonfigValue("opto_run.json",[],{}, workmode, "1")
    if res[0]!="1":
        saveToMassConfigJSON("0", {}, workmode)
        sys.exit()

    opto_config_dic=res[2]

    meter_serial_number_list=opto_config_dic["meter_serial_number_list"]

    meter_tech_number_list=opto_config_dic["meter_tech_number_list"]


    meter_status_test_list=opto_config_dic["meter_status_test_list"]

    workmode=opto_config_dic["workmode"]

    employee_id=opto_config_dic["employee_id"]

    employee_pw_encrypt=opto_config_dic["employee_pw_encrypt"]
    pw=cryptStringSec("расшифровать", employee_pw_encrypt)[2]

    meter_soft_list=opto_config_dic["meter_soft_list"]  

    com_config_current_select =opto_config_dic["com_config_current_select"]

    com_config_user=opto_config_dic["com_config_user"]

    com_config_current=opto_config_dic["com_config_current"]
        
    mass_number_of_meter=0

    for meter_tech_number in meter_tech_number_list:
        if meter_tech_number!=None and meter_tech_number!="":
            mass_number_of_meter+=1
    
    a_txt="оптопорт, подключенный к COM-порту"
    if com_config_current=="com_config_opto":
        if opto_config_dic["com_config_eqv_com"]=="1" :
            com_config_list=opto_config_dic["multi_com_opto_dic"]["com_name"]
        
        else:
            com_config_list=opto_config_dic["multi_com_config_opto_dic"]["com_name"]

    else:
        a_txt=a_txt="интерфейс RS-485, подключенный к COM-порту"
        if opto_config_dic["com_config_eqv_com"]=="1" :
            com_config_list=opto_config_dic["multi_com_rs485_dic"]["com_name"]

        else:
            com_config_list=opto_config_dic["multi_com_config_rs485_dic"]["com_name"]

    if com_config_current_select=="0":
        if com_config_user!=None and com_config_user!="":
            if com_config_user=="com_config_opto":
                a_txt="оптопорт, подключенный к COM-порту"
                if opto_config_dic["com_config_eqv_com"]=="1" :
                    com_config_list=opto_config_dic["multi_com_opto_dic"]["com_name"]
                
                else:
                    com_config_list=opto_config_dic["multi_com_config_opto_dic"]["com_name"]

            else:
                a_txt="интерфейс RS-485, подключенный к COM-порту"
                if opto_config_dic["com_config_eqv_com"]=="1" :
                    com_config_list=opto_config_dic["multi_com_rs485_dic"]["com_name"]

                else:
                    com_config_list=opto_config_dic["multi_com_config_rs485_dic"]["com_name"]
        
    if len(com_config_list)==0:
        a_err_txt="Не указан COM-порт для проведения " \
            f"проверки конфигурации ПУ."
        a_err_txt_1=f"{a_err_txt}\n" \
            f"{bcolors.OKBLUE}Нажмите Enter."
        printMsgWait(a_err_txt_1, bcolors.FAIL,"1", ["\r"])
        saveToMassConfigJSON("0", {}, workmode)
        sys.exit()
    
    if len(com_config_list)>1:
        if "оптопорт" in a_txt:
            a_txt="оптопорты, подключенные к COM-портам"
        
        else:
            a_txt="интерфейс RS-485, подключенные к COM-портам"

    a_com_str=", ".join(com_config_list)
    txt=f"Для проверки конфигурации ПУ используем {a_txt} {a_com_str}."
    printGREEN(txt)

    

    
    
    
    
    
    


   
    a_dic={"mask_file_name": ""}
    saveConfigValue("mass_log_line_multi.json",a_dic, workmode,
        "заменить файл")

    saveToMassConfigJSON("5", {}, workmode)

    res=readGonfigValue("mass_config.json",[],{}, workmode, "1")
    if res[0]!="1":
        saveToMassConfigJSON("0", {}, workmode)
        sys.exit()

    mass_config_dic=res[2]

    
    time_wait_exec=mass_config_dic["time_wait_exec"]

    time_wait_open_window=mass_config_dic["time_wait_open_window"]    


    mass_control_com_port=mass_config_dic["mass_control_com_port"]
    
    res = prepareFileMask(meter_serial_number_list, "1", workmode)
    if res[0] == "0":
        saveToMassConfigJSON(res[0], {}, workmode)
        sys.exit()

    
    while True:

        err_in_log_list=[]

        except_in_log_list=[]

        log_line_file_list=[]

        no_substrings_found_list=[]
        
        res = prepareFileLog(employee_id, mass_number_of_meter, "1",
            workmode)
        if res[0] == "0":
            saveToMassConfigJSON(res[0], {}, workmode)
            sys.exit()

        file_log_path = res[3]

        
        res = setProgramMassSetting(meter_tech_number_list,
            mass_number_of_meter, com_config_list, 
            mass_control_com_port, "1", workmode)
        if res[0]!="1":
            saveToMassConfigJSON(res[0], {}, workmode)
            sys.exit()
        

        replaceMyTitleWindows()
        
        
        res=executeMassProdAutoConfig(meter_tech_number_list,
            employee_id, pw, "1", time_wait_exec)
        if res[0]=="0":
            saveToMassConfigJSON(res[0], {}, workmode)
            exitProgram()



                

             

        res=cycleOnlineReadLog(file_log_path, meter_tech_number_list,
            meter_status_test_list, meter_serial_number_list, mass_number_of_meter, 
            meter_soft_list, workmode, "1")

        if res[0] in ["0", "3", "8", "9"]:
            saveToMassConfigJSON(res[0], {}, workmode)
            exitProgram()

        elif res[0]=="1":
            mass_log_line_dic=res[2]

            break

        elif res[0]=="2":
            saveToMassConfigJSON("4", {}, workmode)
            a_txt="Проверка конфигурации ПУ была проведена в " \
                "ручном режиме."
            a_mass_dic={}
            for meter_tech_number in meter_tech_number_list:
                a_mass_dic["err_in_log_list"]=[]
                a_mass_dic["except_in_log_list"]=[]
                a_mass_dic["no_substrings_found_list"]=[]
                a_mass_dic["analisys_res_0"]="вручную",
                a_mass_dic["analisys_res_1"]="вручную",
                a_mass_dic["log_line_file_list"]=[]
                a_dic={meter_tech_number:a_mass_dic}
                saveConfigValue("mass_log_line_multi.json",a_dic, workmode)

            exitProgram()
        

    
    saveToMassConfigJSON("1", {}, workmode)


    saveConfigValue("mass_log_line_multi.json", mass_log_line_dic, workmode)

    
    analisys_res_ok=True

    if print_result=="1":
        printGREEN("\nВывод:")
        for meter_tech_number in meter_tech_number_list:
            a_dic=mass_log_line_dic[meter_tech_number]
            err_in_log_list=a_dic["err_in_log_list"]
            except_in_log_list=a_dic["except_in_log_list"]
            no_substrings_found_list=a_dic["no_substrings_found_list"]
            log_line_file_list=a_dic["log_line_file_list"]
            analisys_res_0=a_dic["analisys_res_0"]
            analisys_res_1=a_dic["analisys_res_1"]

            if analisys_res_0=="" or analisys_res_1=="":
                printWARNING(f"Проверка конфигурации ПУ № {meter_tech_number} "
                    "была пропущена.")
                continue
            
            a_conclusion_txt=f"{bcolors.OKGREEN}По результатам " \
                "проверки установлено, что конфигурация ПУ №" \
                f"{meter_tech_number} соответствует требованиям."
            
            if analisys_res_0=="bad" or analisys_res_1=="bad":
                analisys_res_ok=False
                a_conclusion_txt=f"{bcolors.FAIL}В ходе проведения проверки " \
                    f"конфигурации ПУ №{meter_tech_number} выявлены " \
                    f"замечания."
                
            if len(err_in_log_list)>0:
                a_list=listCopy(err_in_log_list)
                a_list=insHiphenColor(a_list,"- ", bcolors.FAIL)[2]
                a_list_txt='\n'.join(a_list)
                a_conclusion_txt=f"{bcolors.FAIL}В ходе проведения проверки " \
                    f"конфигурации ПУ №{meter_tech_number} выявлены " \
                    f"следующие замечания:\n{a_list_txt}"
                
            printColor(a_conclusion_txt)

            if len(except_in_log_list)>0:
                a_list=listCopy(except_in_log_list)
                a_list=insHiphenColor(a_list,"- ", "")[2]
                a_list_txt='\n'.join(a_list)
                printGREEN("Нижеперечисленные параметры конфигурации " 
                    "соответствуют требованиям списка исключения:\n"
                    f"{a_list_txt}")
                
            if len(no_substrings_found_list)>0:
                a_list=listCopy(no_substrings_found_list)
                a_list=insHiphenColor(a_list,"- ", "")[2]
                a_list_txt='\n'.join(a_list)
                printWARNING("Нижеперечисленные контрольные подстроки из " 
                    "списка status_true_list не были найдены в log-файле:\n"
                    f"{a_list_txt}")
                
            printGREEN(f"Проверка конфигурации ПУ № {meter_tech_number} закончена.")

    
    exitProgram()



if  __name__ ==  '__main__' : 

    mainMassConfig()


