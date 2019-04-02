import os
import platform
import shutil

def creation_date(path_to_file):
    if platform.system() == 'Windows':
        return os.path.getctime(path_to_file)
    else:
        # Проблема с Linux
        # Последний раз, когда содержимое файла было изменено
        stat = os.stat(path_to_file)
        try:
            return stat.st_birthtime
        except AttributeError:
            return stat.st_mtime

path = "./"
files = []
new_file_date = -1
for file in os.listdir(path):
    if file.endswith('xlsx'):  # file[len(file)-file[::-1].find('.'):] == 'xlsx':
        files.append(file)
        new_file_date = creation_date(path+file) if creation_date(path+file) > new_file_date else new_file_date
for file in files:
    if new_file_date > creation_date(path+file):
        os.rename(path + file, path + './old/' + file)
os.system("python ./scripts/parse.py")