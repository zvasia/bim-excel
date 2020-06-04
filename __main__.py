from datetime import datetime
from datetime import timedelta
import pandas as pd

# pd.options.display.max_rows = 20

sd_person_file_path = "/Users/vasiliy/Yandex.Disk.localized/rep/python/bim-excel/SourceData_JSON/small_data_persons" \
                      ".json"
sd_contacts_file_path = "/Users/vasiliy/Yandex.Disk.localized/rep/python/bim-excel/SourceData_JSON" \
                         "/small_data_contracts.json"
bd_person_file_path = "/Users/vasiliy/Yandex.Disk.localized/rep/python/bim-excel/SourceData_JSON/big_data_persons.json"
bd_contacts_file_path = "/Users/vasiliy/Yandex.Disk.localized/rep/python/bim-excel/SourceData_JSON" \
                         "/big_data_contracts.json"

writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')

# 1.3. Загрузить списки людей из файлов small_data и big_data в Excel на разные листы.
# 1.4. На листе small_data записи должны быть отсортированы по фамилии, на листе big_data по
# имени.


# Получаем dataframe и извлекаем имена и фамилии для последующей сортировки и фильтрации
sd_persons = pd.read_json(sd_person_file_path)
sd_persons['First name'] = sd_persons['Name'].str.split(' ').str.get(1)
sd_persons['Last name'] = sd_persons['Name'].str.split(' ').str.get(0)

bd_persons = pd.read_json(bd_person_file_path)
bd_persons['First name'] = bd_persons['Name'].str.split(' ').str.get(1)
bd_persons['Last name'] = bd_persons['Name'].str.split(' ').str.get(0)

# сортируем и сохраняем в xlsx
sd_persons.sort_values(by='Last name').to_excel(writer, sheet_name="SD Persons")
bd_persons.sort_values(by='First name').to_excel(writer, sheet_name="BD Persons")


# 1.5. Найти людей в small_data, которых нет в big_data по фамилии и вывести их на отдельный
# лист Excel либо в файл.
uniq_persons = sd_persons.loc[~sd_persons['Last name'].isin(bd_persons['Last name'])]
uniq_persons.to_excel("output.xlsx", sheet_name="Uniq SD Persons")


# 1.6. Найти группу однофамильцев, у которых разница в возрасте составляет 10 лет и вывести
# на отдельный лист Excel либо в файл.

concat_df = bd_persons.merge(sd_persons, how='outer').drop_duplicates(subset='ID').sort_values('ID')

homonyms = pd.DataFrame([])

for row in concat_df.iterrows():
    for row2 in concat_df.iterrows():
        # Проверяем совпадают ли фамилии
        if row[1][4] == row2[1][4]:
            # Проверяем разницу в возрасте
            if abs(row[1][2] - row2[1][2]) == 10:
                homonyms = homonyms.append(row[1])

homonyms.to_excel(writer, "Homonyms 10y")

# 1.7. Найти людей, у которых в фамилии или имени содержатся английские символы и вывести
# их на отдельный лист Excel либо в файл.


symbols = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v',
           'w', 'x', 'y', 'z']
result = pd.DataFrame([])

# Проверяем наличие символов в имени
for s in range(len(symbols)):
    result = result.append(concat_df[concat_df['Name'].str.contains(symbols[s], case=False)])

result.drop_duplicates().to_excel(writer, "Eng symbols")

# 2.4. Вывести список людей, отсортированный в обратном порядке по количеству контактов с
# другими людьми, при условии, что контакт происходил 5ть минут и более. Если контакт
# происходил менее 5ти минут, это контактом не считается.

sd_contacts = pd.read_json(sd_contacts_file_path)
bd_persons = pd.read_json(bd_contacts_file_path)


# Проставляем в список людей длительность контакта
def add_time(pid, pdf, time):
    if pdf.loc[pdf['ID'] == pid, ['Time']].isna().values:
        pdf.loc[pdf['ID'] == pid, ['Time']] = time
    else:
        pdf.loc[pdf['ID'] == pid, ['Time']] = pdf.loc[pdf['ID'] == pid, ['Time']].values + time


datetime_format = "%d.%m.%Y %H:%M:%S"

concat_df['Contacts'] = 0
concat_df['Time'] = None

for row in sd_contacts.iterrows():
    from_date = datetime.strptime(row[1][0], datetime_format)
    to_date = datetime.strptime(row[1][1], datetime_format)
    contact_time = to_date - from_date
    if contact_time > timedelta(minutes=5):
        concat_df.loc[concat_df['ID'] == row[1][2], ['Contacts']]\
            = concat_df.loc[concat_df['ID'] == row[1][2]]['Contacts'].values + 1
        concat_df.loc[concat_df['ID'] == row[1][3], ['Contacts']]\
            = concat_df.loc[concat_df['ID'] == row[1][3], ['Contacts']]['Contacts'].values + 1
        add_time(row[1][2], concat_df, contact_time)
        add_time(row[1][3], concat_df, contact_time)

concat_df.sort_values(by='Contacts', ascending=False).to_excel(writer, "By contacts")


# 2.5. Вывести список людей, отсортированный в обратном порядке по общей длительности
# контакта с другими людьми.

concat_df.sort_values(by='Time', ascending=False).to_excel(writer, "By time")


# 2.6. Найти возрастную группу людей, которая имеет наиболее частый контакт с другими
# людьми. Так же, как и в задании 2.1 контактом считается контакт длительностью 5 и более
# минут.

ages = pd.DataFrame(concat_df['Age'].drop_duplicates())
ages['Contacts'] = None

for row in concat_df.iterrows():
    for age in ages.iterrows():
        if row[1][2] == age[1][0]:
            if ages.loc[ages['Age'] == row[1][2], ['Contacts']].isna().values:
                ages.loc[ages['Age'] == row[1][2], ['Contacts']] = row[1][5]
            else:
                ages.loc[ages['Age'] == row[1][2], ['Contacts']] = ages.loc[ages['Age'] == row[1][2], ['Contacts']] \
                                                                   + row[1][5]

ages.sort_values(by='Contacts', ascending=False).head(1).to_excel(writer, "Most contacts per age")
writer.save()
