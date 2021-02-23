from docx import Document
import pandas as pd
from os import listdir
from os.path import isfile, join

# В папке с файлами выбираем только файлы с расширением .docx
mypath = "/Users/VivTekki/Desktop/Яндекс.Диск/ЕУ ДПО ЯНДЕКС/project_tombs"
onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f)) and '.docx' in f]

# Создаем датафрейм
df_text = pd.DataFrame(columns=['id', 'jewish_text', 'rus_transl', 'folder', 'im_name'])

# Открываем найденные файлы и ищем таблицы в них
for file in onlyfiles:
    doc = Document(mypath + "/" + file)

    table_doc = doc.tables[0]
    folder_id = ""

# В таблицах идем по ячейкам, записываем содержимое в нужные нам переменные
    for row in table_doc.rows:
        cells = row.cells
        id = cells[0].text
        jewish_text = cells[1].text
        rus_transl = cells[2].text
        if id.isdigit() == False:
            continue

# решение с "> 1000" немного костыль, открыта более элегантным рабочим вариантам)
# это нужно, чтобы отделить номер папки (> 10 символов) от номера файла
# номера файлов в отдельных папках Гобермана не превышают 100, но чтобы перестраховаться, я указала 1000
        if int(id) > 1000:
            folder_id = id
            folder_name = file.split('.docx')[0].split('каталог ')[1] + "/" + folder_id
        else:
            df_id = folder_id + "_" + id
            im_name = id.zfill(6)

            df_text = df_text.append({'id': df_id,
                                      'jewish_text': jewish_text,
                                      'rus_transl': rus_transl,
                                      'folder': folder_name,
                                      'im_name': im_name}, ignore_index=True)

df_text.to_csv(mypath + '/Goberman_parsed.csv', index=False)