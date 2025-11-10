# обновляем программы otk в автоматическом режиме 220725

import os
from libs.otkLib import *
from otk_opto import optoRunVarRead


# 170425
# если данная программа запустилась как основная программа, 
# а не подгружаемый модуль
if  __name__ ==  '__main__' : 
    # прочитаем сохраненные в файле "opto_run.json" 
    # конифг. значения
    default_value_dict = optoRunVarRead()
    # присвоим значения переменным
    # режим работы программ
    workmode=default_value_dict['workmode']

    # получим номер последнего обновления локальных файлов,
    # указанного в конфиг.файле "update_progr.json" и номер
    # актуального обновления в общей папке "\\\\PE-WS-12-166\\Update\\autoUpdate",
    # указанного в файле "progr_num.json"
    res=getVersUpdProgrFiles(workmode=workmode)
    # если вернулась ошибка
    if res[0]=="0":
        # выйдем из программы
        sys.exit()
    # прочитаем номер последнего обновления локальных файлов
    prog_last_upd_num=res[2]
    # прочитаем номер актуального обновления файлов в общей папке
    prog_actual_upd_num=res[3]
    # прочитаем путь к общей папке autoUpdate
    auto_update_dir=res[4]
    # прочитаем дату/время пакета последнего обновления локальных файлов
    prog_last_upd_moment=res[5]
    # рочитаем дату/время пакета актуального обновления
    prog_actual_upd_moment=res[6]

    # если номер последнего обновления больше или равен номеру
    # актуального обновления
    if prog_last_upd_num>=prog_actual_upd_num:
        # выведем сообщение
        printGREEN("Обновление программ не требуется. Установлена версия № " \
              f"{prog_last_upd_num} от {prog_last_upd_moment}.\n")
        # ожидаем нажатия Enter
        keystrokeEnter()
        # выйдем из программы
        sys.exit()

    # получим ссылку на рабочий каталог
    res=getWorkDirLink()
    # если вернулась ошибка
    if res[0]=="0":
        # ожидаем нажатия Enter
        keystrokeEnter()
        # выйдем из программы
        sys.exit()
    #ссылка на рабочий каталог
    work_dir=res[2]
    
    # подготовим список для всех описаний изменений в версиях
    description_of_changes_all_list=[]

    # Будем выполнять цикл, пока номер тек.обновления будет меньше
    #  или равен номеру актуального обновления программ
    # подготовим переменную для сводного списка с именами файлов и путями для обновления
    file_path_list=[]
    while prog_last_upd_num<prog_actual_upd_num:
        # перейдем на сл. номер обновления
        prog_last_upd_num+=1
        # сформируем полный путь до папки с очередным обновлением
        upd_dir_cur=f"{auto_update_dir}\\{str(prog_last_upd_num)}"
        # если в папке autoUpdate имеется папка с номером очередной версии
        if os.path.isdir(upd_dir_cur):
            # сформируем полный путь до ф.file_path.json
            file_name_1=os.path.join(upd_dir_cur, "file_path.json")
            # загрузим список имен файлов и путей к ним из ф.file_path.json
            try:
                with open(file_name_1, "r", errors="ignore",encoding='utf-8') as file:
                    content=json.load(file)
            except Exception as e:
                print(f"{bcolors.FAIL}Не удалось получить данные из ф. {file_name_1}: " 
                      f"{e.args[1]}.{bcolors.ENDC}")
                # ожидаем нажатия Enter
                keystrokeEnter()
                # завершим работу программы
                sys.exit()
            
            # получим список с именами файлов для обновления
            file_upd_list=content["list"]

            # получим описание изменений
            description_of_changes_list=content["description_of_changes_list"]

            # если имеется описание изменений, то добавим его в список описаний
            if len(description_of_changes_list)>0:
                # запишем номер версии обновления
                description_of_changes_all_list.append(f"Версия № {prog_last_upd_num}:")
                # добавим список описаний изменений текущей версии
                description_of_changes_all_list.extend(description_of_changes_list)


            # переберем сформированный список с именами файлов для обновления
            for file_upd_cur_dic in file_upd_list:
                # установим метку нового файла для списка
                file_new="1"
                # переберем сводный список с именами файлов
                i=0
                while i<len(file_path_list):
                    # получим тек. зн-е из сводного списка
                    file_path_cur_dic=file_path_list[i]
                    # если ключ "fileDestPath" содержит аналогичное значение из списка
                    if file_path_cur_dic["fileDestPath"]==f"{work_dir}\\{file_upd_cur_dic['fileNameDest']}":
                        # заменим ссылку в сводном списке на новую ссылку на эталонный файл
                        file_path_list[i]["fileSourcePath"]=f"{upd_dir_cur}\\{file_upd_cur_dic['fileNameSource']}"
                        # сбросим метку нового файла в спиcке
                        file_new="0"
                        # выйдем из цикла while
                        break
                    # перейдем к следующей записи в списке
                    i+=1
                # если такого файла в списке нет (метка нового файла ="1")
                if file_new=="1":
                    # подготовим переменную для сохранения данных
                    dic_1={}
                    # запишем информацию о новом файле в словарь
                    dic_1["fileSourcePath"]=f"{upd_dir_cur}\\{file_upd_cur_dic['fileNameSource']}"
                    # сформируем полный путь до файла в целевой папке с учетом рабочего каталога
                    dic_1["fileDestPath"]=f"{work_dir}\\{file_upd_cur_dic['fileNameDest']}"
                    dic_1["updateMode"]=file_upd_cur_dic["mode"]
                    # добавим информацию о новом файле в сводный список
                    file_path_list.append(dic_1)
    # произведем обновление файлов с выводом сообщений об ошибках
    res=updateFiles(file_list=file_path_list,err_msg_print="1")
    # если обновление прошло успешно
    if res[0]=="1":
        printGREEN(f"\nФайлы программы обновлены.")
        # преобразуем список с описаниями изменений в строку
        description_of_changes="\n".join(description_of_changes_all_list)
        # если описания имеются
        if description_of_changes!="":
            # выведем на экран описание изменений
            printWARNING(f"\nОписание внесенных изменений:\n{description_of_changes}")
        # имя конфиг.файла
        file_name="update_progr_local.json"
        # номер и дату/время пакета актуального обновления файлов запишем в словарь
        # конфиг. данных
        config_dict={"progLastUpdNum":prog_actual_upd_num,
                     "progLastUpdMoment":prog_actual_upd_moment}
        saveConfigValue(file_name_in=file_name,var_config_dict=config_dict,
                        workmode=workmode)
        
    # ожидаем нажатия Enter
    keystrokeEnter()

    # Запустим ф.otk_menu.bat
    # получим полный путь к ф.otk_menu.bat
    _, ans2, a_file_path = getUserFilePath('otk_menu.bat')
    # если указан путь к файлу
    if a_file_path != "":    
        txt1="Пожалуйста подождите, идет загрузка " \
            "стртового меню программы...\n"
        printGREEN(txt1)
        # запустим на исполнение файл otk_menu.bat
        # в виде отдельного окна
        subprocess.Popen(f"start {a_file_path}", shell=True)

    # получим дескриптор нашего окна
    hwnd = win32gui.GetForegroundWindow()
    
    # закроем свое окно программы
    actionsSelectedtWindow([], hwnd, "закрыть", "1")
    
    # выйдем из программы
    sys.exit()
    