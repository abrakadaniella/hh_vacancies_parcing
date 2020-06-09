import telebot
import requests
import pandas as pd
import json
from bs4 import BeautifulSoup as bs
from xlsxwriter.utility import xl_rowcol_to_cell
from config import token

bot = telebot.TeleBot(token)

parameters = {}

markup = telebot.types.ReplyKeyboardMarkup(one_time_keyboard=True, row_width=1)
itembtn1 = telebot.types.KeyboardButton('Vacancy')
itembtn3 = telebot.types.KeyboardButton('Experience')
itembtn4 = telebot.types.KeyboardButton('Employment')
itembtn5 = telebot.types.KeyboardButton('Schedule')
itembtn6 = telebot.types.KeyboardButton('Region')
itembtn8 = telebot.types.KeyboardButton('Specialization')
itembtn9 = telebot.types.KeyboardButton('Industry')
itembtn11 = telebot.types.KeyboardButton('Only with Salary')
itembtn12 = telebot.types.KeyboardButton('Currency')
itembtn14 = telebot.types.KeyboardButton('Period')
itembtn15 = telebot.types.KeyboardButton('Starting Date')
itembtn16 = telebot.types.KeyboardButton('Ending Date')
itembtn17 = telebot.types.KeyboardButton('Premium')
itembtn2 = telebot.types.KeyboardButton('Done')
itembtn7 = telebot.types.KeyboardButton('Number of Pages')
itembtn10 = telebot.types.KeyboardButton('Per page')
markup.add(itembtn1, itembtn3, itembtn4, itembtn5, itembtn6, itembtn8, itembtn9, itembtn11, itembtn12,
           itembtn14, itembtn15, itembtn16, itembtn17, itembtn7, itembtn10, itembtn2)


@bot.message_handler(commands=['start'])
def start_msg(message):
    bot.send_message(message.chat.id, "Из предложенных вариантов выберите параметр,\
    который хотите указать для поиска и следуйте инструкциям, которые пришлёт бот.", reply_markup=markup)


# VACANCY PARAMETER
@bot.message_handler(regexp='Vacancy')
def vac_btn(message):
    bot.send_message(message.chat.id, """Пожалуйста, введите название вакансии в формате: "Vac: ______"
Например, Vac: Data Engineer AND DWH. 
Для ознакомления с поисковым языком запросов hh, пройдите по этой ссылке: https://hh.ru/article/1175""")


@bot.message_handler(func=lambda message: 'Vac:' in message.text, content_types=['text'])
def recieved_vac(message):
    parameters['text'] = message.text.split(':')[1].strip()
    bot.send_message(message.chat.id, "Выберите другой параметр, или нажмите кнопку \"Done\", если Вы закончили.",
                     reply_markup=markup)


# EXPERIENCE PARAMETER
@bot.message_handler(regexp='Experience')
def exp_btn(message):
    bot.send_message(message.chat.id, """Пожалуйста введите один из вариантов, указанных ниже:
  
* noExp   
* between1And3
* between3And6
* moreThan6

Формат сообщения: \"Exp: _______\"""")


@bot.message_handler(func=lambda message: 'Exp:' in message.text, content_types=['text'])
def recieved_exp(message):
    if message.text == 'Exp: noExp':
        parameters['experience'] = 'noExperience'
    else:
        parameters['experience'] = message.text.split(':')[1].strip()
    bot.send_message(message.chat.id, "Выберите другой параметр, или нажмите кнопку \"Done\", если Вы закончили.",
                     reply_markup=markup)


# EMPLOYMENT PARAMETER
@bot.message_handler(regexp='Employment')
def emp_btn(message):
    bot.send_message(message.chat.id, """Пожалуйста введите один из вариантов, указанных ниже:
    
* full
* part
* project
* volunteer
* probation

Формат сообщения: \"Emp: _______\". 
Можно указать несколько вариантов, для этого нужно разделить их запятой.""")


@bot.message_handler(func=lambda message: 'Emp:' in message.text, content_types=['text'])
def recieved_emp(message):
    params = message.text.split(':')[1].strip().split()
    new_params = []
    for value in params:
        val = value.strip().replace(',', '')
        new_params.append(val)
    parameters['employment'] = new_params
    bot.send_message(message.chat.id, "Выберите другой параметр, или нажмите кнопку \"Done\", если Вы закончили..",
                     reply_markup=markup)


# SCHEDULE PARAMETER
@bot.message_handler(regexp='Schedule')
def sch_btn(message):
    bot.send_message(message.chat.id, """Пожалуйста введите один из вариантов, указанных ниже:
    
* fullDay
* shift
* flexible
* remote
* flyInFlyOut

Формат сообщения: \"Sch: _______\". 
Можно указать несколько вариантов, для этого нужно разделить их запятой.""")


@bot.message_handler(func=lambda message: 'Sch:' in message.text, content_types=['text'])
def received_sch(message):
    params = message.text.split(':')[1].strip().split()
    new_params = []
    for value in params:
        val = value.strip().replace(',', '')
        new_params.append(val)
    parameters['schedule'] = new_params
    bot.send_message(message.chat.id, "Выберите другой параметр, или нажмите кнопку \"Done\", если Вы закончили.",
                     reply_markup=markup)


# AREA PARAMETER
@bot.message_handler(regexp='Region')
def reg_btn(message):
    bot.send_message(message.chat.id, """Пожалуйста, введите название страны/города в формате: \"Area: _______\".
Можно указать несколько вариантов, для этого нужно разделить их запятой.""")


@bot.message_handler(func=lambda message: 'Area:' in message.text, content_types=['text'])
def received_area(message):
    regions = message.text.split(':')[1].strip().split()
    regions_separate = []
    for value in regions:
        val = value.strip().replace(',', '')
        regions_separate.append(val)
    area_url = 'https://api.hh.ru/areas'
    url_cont = requests.get(area_url).content
    area_json = json.loads(url_cont)
    area_ids = []
    for region in regions_separate:
        for i in range(len(area_json)):
            if region == area_json[i]['name']:
                area_id = area_json[i]['id']
                if area_id not in area_ids:
                    area_ids.append(area_id)
            for j in range(len(area_json[i]['areas'])):
                if region == area_json[i]['areas'][j]['name']:
                    area_id = area_json[i]['areas'][j]['id']
                    if area_id not in area_ids:
                        area_ids.append(area_id)
                for k in range(len(area_json[i]['areas'][j]['areas'])):
                    if region == area_json[i]['areas'][j]['areas'][k]['name']:
                        area_id = area_json[i]['areas'][j]['areas'][k]['id']
                        if area_id not in area_ids:
                            area_ids.append(area_id)

    parameters['area'] = area_ids
    bot.send_message(message.chat.id, "Выберите другой параметр, или нажмите кнопку \"Done\", если Вы закончили.",
                     reply_markup=markup)


# SPECIALIZATION PARAMETER
@bot.message_handler(regexp='Specialization')
def spec_btn(message):
    bot.send_message(message.chat.id, """Пожалуйста, введите название специализации в формате: \"Spec: _______\".
Можно указать несколько вариантов, для этого нужно разделить их точкой.""")


@bot.message_handler(func=lambda message: 'Spec:' in message.text, content_types=['text'])
def received_spec(message):
    params = message.text.split(':')[1].strip().split('.')
    sep_spec = []
    for value in params:
        val = value.strip()
        sep_spec.append(val)
    spec_url = 'https://api.hh.ru/specializations'
    url_cont = requests.get(spec_url).content
    spec_json = json.loads(url_cont)
    spec_ids = []
    for spec in sep_spec:
        for i in range(len(spec_json)):
            if spec == spec_json[i]['name']:
                spec_id = spec_json[i]['id']
                if spec_id not in spec_ids:
                    spec_ids.append(spec_id)
            for j in range(len(spec_json[i]['specializations'])):
                if spec == spec_json[i]['specializations'][j]['name']:
                    spec_id = spec_json[i]['specializations'][j]['id']
                    if spec_id not in spec_ids:
                        spec_ids.append(spec_id)

    parameters['specialization'] = spec_ids
    bot.send_message(message.chat.id, "Выберите другой параметр, или нажмите кнопку \"Done\", если Вы закончили.",
                     reply_markup=markup)


# INDUSTRY PARAMETER
@bot.message_handler(regexp='Industry')
def ind_btn(message):
    bot.send_message(message.chat.id, """Пожалуйста, введите название индустрии в формате: \"Ind: _______\".
Можно указать несколько вариантов, для этого нужно разделить их точкой.""")


@bot.message_handler(func=lambda message: 'Ind:' in message.text, content_types=['text'])
def received_ind(message):
    params = message.text.split(':')[1].strip().split('.')
    ind_sep = []
    for value in params:
        val = value.strip()
        ind_sep.append(val)
    ind_url = 'https://api.hh.ru/industries'
    url_cont = requests.get(ind_url).content
    ind_json = json.loads(url_cont)
    ind_ids = []
    for ind in ind_sep:
        for i in range(len(ind_json)):
            if ind == ind_json[i]['name']:
                ind_id = ind_json[i]['id']
                if ind_id not in ind_ids:
                    ind_ids.append(ind_id)
            for j in range(len(ind_json[i]['industries'])):
                if ind == ind_json[i]['industries'][j]['name']:
                    ind_id = ind_json[i]['industries'][j]['id']
                    if ind_id not in ind_ids:
                        ind_ids.append(ind_id)

    parameters['industry'] = ind_ids
    bot.send_message(message.chat.id, "Выберите другой параметр, или нажмите кнопку \"Done\", если Вы закончили.",
                     reply_markup=markup)


# CURRENCY PARAMETER
@bot.message_handler(regexp='Currency')
def cur_btn(message):
    bot.send_message(message.chat.id, "Пожалуйста, введите валютную аббреваитуру в формате: \"Cur: _______\". \
Например, \"Cur: KZT\"")


@bot.message_handler(func=lambda message: 'Cur:' in message.text, content_types=['text'])
def recieved_cur(message):
    parameters['currency'] = message.text.split(':')[1].strip()
    bot.send_message(message.chat.id, "Выберите другой параметр, или нажмите кнопку \"Done\", если Вы закончили.",
                     reply_markup=markup)


# ONLY WITH SALARY PARAMETER
@bot.message_handler(regexp='Only with Salary')
def ows_btn(message):
    bot.send_message(message.chat.id, "Пожалуйста, введите 'true' или 'false' в формате: \"OWS: _______\". \
Например, \"OWS: true\"")


@bot.message_handler(func=lambda message: 'OWS:' in message.text, content_types=['text'])
def recieved_ows(message):
    parameters['only_with_salary'] = message.text.split(':')[1].strip()
    bot.send_message(message.chat.id, "Выберите другой параметр, или нажмите кнопку \"Done\", если Вы закончили.",
                     reply_markup=markup)


# PERIOD PARAMETER
@bot.message_handler(regexp='Period')
def per_btn(message):
    bot.send_message(message.chat.id, """Пожалуйста, введите количество дней, за которое нужно собрать вакансии
в формате: \"Per: _______\". 
Максимальное количество дней: 30.""")


@bot.message_handler(func=lambda message: 'Per:' in message.text, content_types=['text'])
def received_per(message):
    parameters['period'] = message.text.split(':')[1].strip()
    bot.send_message(message.chat.id, "Выберите другой параметр, или нажмите кнопку \"Done\", если Вы закончили.",
                     reply_markup=markup)


# DATE FROM PARAMETER
@bot.message_handler(regexp='Starting Date')
def stdt_btn(message):
    bot.send_message(message.chat.id,
                     "Пожалуйста, введите дату с которой нужно начать собирать вакансии в формате: \
\"Stdt: YYYY-MM-DD\". Например, \"Stdt: 2020-05-05\"")


@bot.message_handler(func=lambda message: 'Stdt:' in message.text, content_types=['text'])
def received_stdt(message):
    parameters['date_from'] = message.text.split(':')[1].strip()
    bot.send_message(message.chat.id, "Выберите другой параметр, или нажмите кнопку \"Done\", если Вы закончили.",
                     reply_markup=markup)


# DATE TO PARAMETER
@bot.message_handler(regexp='Ending Date')
def end_btn(message):
    bot.send_message(message.chat.id, "Пожалуйста, введите дату, до которой нужно собрать вакансии в формате: \"End: \
    \"YYYY-MM-DD\". Например, \"End: 2020-05-05\"")


@bot.message_handler(func=lambda message: 'End:' in message.text, content_types=['text'])
def received_end(message):
    parameters['date_to'] = message.text.split(':')[1].strip()
    bot.send_message(message.chat.id, "Выберите другой параметр, или нажмите кнопку \"Done\", если Вы закончили.",
                     reply_markup=markup)


# PREMIUM PARAMETER
@bot.message_handler(regexp='Premium')
def pr_btn(message):
    bot.send_message(message.chat.id, "Пожалуйста, введите 'true' или 'false' в формате: \"Pr: _______\". Например, \
\"Pr: true\"")


@bot.message_handler(func=lambda message: 'Pr:' in message.text, content_types=['text'])
def received_end(message):
    parameters['premium'] = message.text.split(':')[1].strip()
    bot.send_message(message.chat.id, "Выберите другой параметр, или нажмите кнопку \"Done\", если Вы закончили.",
                     reply_markup=markup)


# NUMBER OF PAGES PARAMETER
@bot.message_handler(regexp='Number of Pages')
def num_pgs_btn(message):
    bot.send_message(message.chat.id, "Пожалуйста, введите количество страниц, на скольких искать вакансии в формате: \
\"NP: _______\".")


@bot.message_handler(func=lambda message: 'NP:' in message.text, content_types=['text'])
def received_num_pgs(message):
    parameters['page'] = message.text.split(':')[1].strip()
    bot.send_message(message.chat.id, "Выберите другой параметр, или нажмите кнопку \"Done\", если Вы закончили.",
                     reply_markup=markup)


# PER PAGE PARAMETER
@bot.message_handler(regexp='Per Page')
def per_page_btn(message):
    bot.send_message(message.chat.id, "Пожалуйста, введите количество вакансий на одной странице в формате: \
\"PPg: _______\". Максимальное значение: 100.")


@bot.message_handler(func=lambda message: 'PPg:' in message.text, content_types=['text'])
def received_per_page(message):
    parameters['per_page'] = message.text.split(':')[1].strip()
    bot.send_message(message.chat.id, "Выберите другой параметр, или нажмите кнопку \"Done\", если Вы закончили.",
                     reply_markup=markup)


# DONE
@bot.message_handler(regexp='Done')
def file(message):
    bot.send_message(message.chat.id, """...Получаю информацию от HH...""")
    url = 'https://api.hh.ru/vacancies?'
    for para in list(parameters.keys()):
        if (para == 'text') and (parameters.get('text', None) is not None):
            to_insert = parameters.get('text').replace(' ', '+')
            url += 'text={}'.format(to_insert)
        if (parameters.get(para, None) is not None) and (para != 'text') and (para != 'page'):
            if type(parameters.get(para)) == list:
                for el in parameters.get(para):
                    url += '&{}={}'.format(para, el)
            else:
                url += '&{}={}'.format(para, parameters.get(para))
    data = []
    for i in range(int(parameters.get('page'))):
        par = {'page': i}
        response = requests.get(url, params=par).json()
        data.append(response)

    vac_urls = []
    for i in range(len(data)):
        for j in range(len(data[i]['items'])):
            vac_urls.append(data[i]['items'][j]['url'])
    urls_cont = []
    for i in range(len(vac_urls)):
        response = requests.get(vac_urls[i]).json()
        urls_cont.append(response)

    vacancy_details = urls_cont[0].keys()
    df = pd.DataFrame(columns=vacancy_details)
    for i in range(len(urls_cont)):
        df = df.append(urls_cont[i], ignore_index=True)

    parameters.clear()

    df['employer'] = df['employer'].apply(lambda x: x.get('name', None) if x is not None else None)
    df['department'] = df['department'].apply(lambda x: x.get('name', None) if x is not None else None)
    df['area'] = df['area'].apply(lambda x: x.get('name', None) if x is not None else None)
    df['address'] = df['address'].apply(lambda x: x.get('raw', None) if x is not None else None)

    def get_name(x):
        attrs = []
        for index in range(len(x)):
            attrs.append(x[index].get('name', None))
        return '; '.join(attrs)

    df['specializations'] = df['specializations'].apply(lambda x: get_name(x) if x is not None else None)
    df['experience'] = df['experience'].apply(lambda x: x.get('name', None) if x is not None else None)
    df['schedule'] = df['schedule'].apply(lambda x: x.get('name', None) if x is not None else None)
    df['employment'] = df['employment'].apply(lambda x: x.get('name', None) if x is not None else None)

    def rem_tags(x):
        soup = bs(x, features="html.parser")
        return soup.get_text()

    df['description'] = df['description'].apply(lambda x: rem_tags(x) if x is not None else None)
    df['key_skills'] = df['key_skills'].apply(lambda x: get_name(x) if x is not None else None)
    df['salary_from'] = df['salary'].apply(lambda x: x.get('from') if x is not None else None)
    df['salary_to'] = df['salary'].apply(lambda x: x.get('to') if x is not None else None)

    df.insert(len(df.columns), 'result', None)

    dic_res = {'name': ['Minimum salary', 'Maximum salary', 'Mean', 'Median', 'Mode']}

    df_res = pd.DataFrame(dic_res)

    final_df = pd.concat([df, df_res], sort=False).reset_index()

    final_df = final_df[
        ['id', 'name', 'employer', 'department', 'salary_from', 'salary_to', 'area', 'address', 'specializations',
         'experience', 'schedule', 'employment', 'description', 'key_skills', 'published_at', 'premium', 'result']]

    final_df = final_df.rename(columns={
        'name': 'vacancy',
        'salary_from': 'starting salary',
        'salary_to': 'salary up to',
        'area': 'region',
        'key_skills': 'key skills',
        'published_at': 'date published'

    })

    writer = pd.ExcelWriter(r'data.xlsx', engine='xlsxwriter')
    final_df.to_excel(writer, index=False, header=True, sheet_name='vacancies')

    workbook = writer.book
    worksheet = writer.sheets['vacancies']

    num_rows = len(df.index)
    column = len(final_df.columns) - 1

    # MIN FORMULA
    min_cell = xl_rowcol_to_cell(num_rows + 1, column)
    min_formula = '=MIN(E2:F{})'.format(num_rows)
    worksheet.write_formula(min_cell, min_formula)

    # MAX FORMULA
    max_cell = xl_rowcol_to_cell(num_rows + 2, column)
    max_formula = '=MAX(E2:F{})'.format(num_rows)
    worksheet.write_formula(max_cell, max_formula)

    # MEAN FORMULA
    mean_cell = xl_rowcol_to_cell(num_rows + 3, column)
    mean_formula = '=AVERAGE(E2:F{})'.format(num_rows)
    worksheet.write_formula(mean_cell, mean_formula)

    # MEDIAN FORMULA
    median_cell = xl_rowcol_to_cell(num_rows + 4, column)
    median_formula = '=MEDIAN(E2:F{})'.format(num_rows)
    worksheet.write_formula(median_cell, median_formula)

    # MODE FORMULA
    mode_cell = xl_rowcol_to_cell(num_rows + 5, column)
    mode_formula = '=MODE(E2:F{})'.format(num_rows)
    worksheet.write_formula(mode_cell, mode_formula)

    writer.save()

    f = open('data.xlsx', 'rb')
    bot.send_document(message.chat.id, f)


bot.polling()
