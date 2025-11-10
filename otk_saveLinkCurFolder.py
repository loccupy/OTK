# 050524

import os
import json  # для чтения и сохранения значений переменных по умолчанию в файле в формате json

# записывает в ф.\libs\link_current_folder.json полный путь
# к корневой папке, где расположен ф.saveLinkCurFolder.py
def saveLinkCurFolder():
    # директория, где расположен тек.ф saveLinkCurFolder.py
    dirname = os.path.dirname(__file__)
    # сформируем полный путь к ф.link_current_folder.json
    file_name = os.path.join(dirname,
        "libs\\link_current_folder.json")
    # создадим словарь
    file_dic = {"link_current_folder":dirname}
    # сохраним словарь в файле в формате json
    with open(file_name, "w", errors="ignore", encoding='utf-8') as file:
        json.dump(file_dic, file)
