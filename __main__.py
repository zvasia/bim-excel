from datetime import datetime
from datetime import timedelta
import pandas as pd

# Вряд ли кому-то интересны пути на твоей машине, пишем относительные!
sd_person_file_path = "./SourceData_JSON/small_data_persons.json"
sd_contacts_file_path = "./SourceData_JSON/small_data_contracts.json"
bd_person_file_path = "./SourceData_JSON/big_data_persons.json"
bd_contacts_file_path = "./SourceData_JSON/big_data_contracts.json"

writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')

# 1.3. Загрузить списки людей из файлов small_data и big_data в Excel на разные листы.
# 1.4. На листе small_data записи должны быть отсортированы по фамилии, на листе big_data по
# имени.

# Получаем dataframe и извлекаем имена и фамилии для последующей сортировки и фильтрации

sd_persons = pd.read_json(sd_person_file_path)
sd_persons[['Last name', 'First name']] = sd_persons['Name'].str.split(' ', n=1, expand=True)

bd_persons = pd.read_json(bd_person_file_path)
bd_persons[['Last name', 'First name']] = bd_persons['Name'].str.split(' ', n=1, expand=True)

# сортируем и сохраняем в xlsx
sd_persons.sort_values(by='Last name').to_excel(writer, sheet_name="SD Persons")
bd_persons.sort_values(by='First name').to_excel(writer, sheet_name="BD Persons")

# 1.5. Найти людей в small_data, которых нет в big_data по фамилии и вывести их на отдельный
# лист Excel либо в файл.

uniq_persons = sd_persons.loc[~sd_persons['Last name'].isin(bd_persons['Last name'])]
uniq_persons.to_excel(writer, sheet_name="Uniq SD Persons")

# 1.6. Найти группу однофамильцев, у которых разница в возрасте составляет 10 лет и вывести
# на отдельный лист Excel либо в файл.

# Получаем полный список людей
concat_df = bd_persons.merge(sd_persons, how='outer').drop_duplicates(subset='ID').sort_values('ID')

# Создаем список совпадений по фамилиям

homonyms = pd.merge(concat_df, concat_df, on='Last name', suffixes=('_1', '_2'))

# Отфильтровываем совпадения по разнице в возрасте

homonyms = homonyms[homonyms.Age_1 - homonyms.Age_2 == 10]

homonyms = homonyms[['Last name', 'ID_1', 'First name_1', 'ID_2', 'First name_2']]

homonyms.to_excel(writer, "Homonyms 10y", index=False)

# 1.7. Найти людей, у которых в фамилии или имени содержатся английские символы и вывести
# их на отдельный лист Excel либо в файл.

result = concat_df[
    (concat_df['Last name'].str.match('.*[a-zA-Z]')) |
    (concat_df['First name'].str.match('.*[a-zA-Z]'))
    ]
result.drop_duplicates().to_excel(writer, "Eng symbols", index=False)

# 2.4. Вывести список людей, отсортированный в обратном порядке по количеству контактов с
# другими людьми, при условии, что контакт происходил 5ть минут и более. Если контакт
# происходил менее 5ти минут, это контактом не считается.

sd_contacts = pd.read_json(sd_contacts_file_path)
bd_contacts = pd.read_json(bd_contacts_file_path)
all_contacts = pd.concat([sd_contacts, bd_contacts])

# Парсим даты
datetime_format = "%d.%m.%Y %H:%M:%S"
convert_date = lambda x: datetime.strptime(x, datetime_format)
all_contacts['From'] = all_contacts['From'].map(convert_date)
all_contacts['To'] = all_contacts['To'].map(convert_date)

# Просчитываем время контакта
all_contacts['Contact time'] = all_contacts.To - all_contacts.From

# Важная штука: на каждый контакт в данных есть только одна строка, с указанием участников в Member1_ID и Member2_ID
# Агрегацию возможно провести только по одному из этих полей, но тогда мы потеряем данные о втором участнике.
# Поэтому продублируем таблицу встреч, поменяв в дубле участников местами.
# Теперь о каждой встрече будет две записи, в первой Member1_ID это одни участник, а во второй - другой.
mirror = all_contacts.copy()
mirror[['Member1_ID', 'Member2_ID']] = mirror[['Member2_ID', 'Member1_ID']]
all_contacts = pd.concat([all_contacts, mirror], ignore_index=1)

# Оставляем только контакты длиннее 5 минут
long_contacts = all_contacts[all_contacts['Contact time'] > timedelta(minutes=5)]
pd.set_option('display.max_columns', None)

# Посчитаем количество встреч на каждого участника
c_count = long_contacts.groupby('Member1_ID')['Member1_ID'].agg(['count'])

# Отсортируем по убыванию
c_count = c_count.sort_values(by=['count'], ascending=False)

# Переименуем столбец Member1_ID в ID, чтобы по нему можно было привязать участников из concat_df
c_count = c_count.reset_index().rename(columns={'Member1_ID': 'ID'})

# Добавим в таблицу имена участников и сохраняем в XLS
contacts_stat = pd.merge(c_count, concat_df, on='ID')[['ID', 'Last name', 'First name', 'count']]
contacts_stat.to_excel(writer, "Contacts by count", index=False)

# 2.5. Вывести список людей, отсортированный в обратном порядке по общей длительности
# контакта с другими людьми.

c_len = all_contacts.groupby('Member1_ID')['Contact time'].agg(['sum'])
c_len = c_len.sort_values(by=['sum'], ascending=False)
c_len = c_len.reset_index().rename(columns={'Member1_ID': 'ID'})

contacts_stat = pd.merge(c_len, concat_df, on='ID')[['ID', 'Last name', 'First name']]
contacts_stat.to_excel(writer, "Contacts by length", index=False)

# 2.6. Найти возрастную группу людей, которая имеет наиболее частый контакт с другими
# людьми. Так же, как и в задании 2.1 контактом считается контакт длительностью 5 и более
# минут.

contacts = pd.merge(concat_df, long_contacts, left_on='ID', right_on='Member1_ID', how='inner')
ages = contacts.groupby('Age')['Contact time'].agg('sum')
ages.sort_values(ascending=False).head(1).to_excel(writer, "Most contactable age")


writer.save()

