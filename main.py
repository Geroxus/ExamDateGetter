import requests
import pandas as pd
from colorama import Fore

data: dict = {}
final: dict = {}

# fetch the website from tuhh and sort everything such that it becomes a searchable dictionary
r = requests.get('https://intranet.tuhh.de/stud/pruefung/index.php')
src_lines = r.text.splitlines()

i = 0
for line in src_lines:
    if 'titleline' in line:
        module: str
        date: str
        # print(line)
        if 'strong' in line:
            module = line.split('strong>')[1][:-2]
        else:
            module = line.split('>')[2][:-4]
        if "(Modul)" in module:
            module = module[:-8]
        # print(lines[i+1])
        date = src_lines[i + 1].split('15%\">')[1][:-5]
        data[module] = date
        # print(module + ': ' + date)
    if i == 0:
        print('start ...')
        # print(lines)
    i += 1

all_dates = pd.DataFrame.from_dict(data, orient='index', columns=['Date'])

# uncomment these 2 lines to make the script create an excel sheet over all dates
with pd.ExcelWriter('./all_dates.xlsx') as writer:
    all_dates.to_excel(writer)

# all_dates.info()  # print out info about the data frame so the user knows whether it was successful or not

# excel_sheet: str = 'K:/Uni/stundenplan.xlsx'  # path to excel sheet to compare all_dates against
# print('start reading \"' + excel_sheet + '\"')
#
# complete_df = pd.read_excel(io=excel_sheet)  # we need to manipulate the excel sheet in here and rewrite completely
# after we are done
# filter_df = pd.read_excel(io=excel_sheet, usecols='H:J', nrows=6, header=0)
filter_df = pd.read_clipboard(sep=None, index_col=None, header=None, names=['Klausur', 'a'])
# print(filter_df.to_string())
filter_list: list = filter_df.get('Klausur').to_list()
for index, entry in enumerate(filter_list):
    if ' /' in entry:
        filter_list[index] = entry.split(' /')[0]
        # print(filter_list[index])
# print(filter_list)
print('Anzahl zu findender Termine:', len(filter_list))
# this is probably a really good example to use list comprehension at but it's dicts and I can't be bothered |future me:
#                                                                                                            |     lmao
for key_value in filter_list:
    found: bool = False
    for key in data.keys():
        if key_value == key:
            print(Fore.BLUE + key + ': ' + data[key] + Fore.RESET)
            final[key] = data[key]
            found = True
            break
    if not found:
        print(Fore.RED + 'Klausur \"' + key_value + '\" konnte nicht gefunden werden.' + Fore.RESET)
print('Termine gefunden')
out = pd.DataFrame.from_dict(final, orient='index', columns=['Date'])
# print(out.to_string())
with pd.ExcelWriter('./exam_dates.xlsx', mode='w') as writer:
    out.to_excel(writer, header=False, index=False)

# print(complete_df.to_string())
# complete_df.to_excel('./test2.xlsx')

out.to_clipboard(index=None, header=None)

# goodnight

print('\nfinished.')
