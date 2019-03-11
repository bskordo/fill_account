import openpyxl
import argparse
import os
import sys
import json


filepath = os.path.join(sys.path[0], "purch_work.json")


def load_json_data(filepath):
    if not os.path.exists(filepath):
        return None
    with open(filepath, 'r') as file_handler:
        return json.load(file_handler)

def get_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('file_which_need_to_fill', help='file_which_need_to_fill')
    arg = parser.parse_args()
    return arg.file_with_address




def write_infromation_into_file(my_file, information_from_json):
    work_book = openpyxl.load_workbook(filename=my_file)
    work_sheet = work_book.active
    work_sheet['B4'] = information_from_json['principal']['full_name']
    work_sheet['D18'] = information_from_json['principal']['short_name']
    work_sheet['D19'] = information_from_json['principal']['inn']
    work_sheet['D20'] = information_from_json['principal']['kpp']
    work_sheet['D21'] = information_from_json['principal']['ogrn']
    work_sheet['F23'] = information_from_json['principal']['registration_certificate_date']
    work_sheet['F24'] = information_from_json['principal']['registration_place']
    #Адрес местонахождения указанный при
    work_sheet['F26'] = information_from_json['principal']['legal_address']['data']['postal_code']
    work_sheet['F27'] = information_from_json['principal']['legal_address']['data']['region']
    work_sheet['F28'] = information_from_json['principal']['legal_address']['data']['city_district']
    work_sheet['F29'] = information_from_json['principal']['legal_address']['data']['city_with_type']
    work_sheet['F30'] = information_from_json['principal']['legal_address']['data']['settlement_with_type']
    work_sheet['F31'] = information_from_json['principal']['legal_address']['data']['street']
    work_sheet['F32'] = information_from_json['principal']['legal_address']['data']['house']
    work_sheet['F33'] = information_from_json['principal']['legal_address']['data']['block']
    if information_from_json['principal']['legal_address']['data']['block'] =='оф':
        work_sheet['F35'] = information_from_json['principal']['legal_address']['data']['flat']


    #End
    #Контакты Компании
    work_sheet['F37'] = information_from_json['principal']['email']
    work_sheet['F38'] = information_from_json['principal']['contact_phone']
    #СВЕДЕНИЯ ОБ ОТВЕТСТВЕННЫХ ЛИЦАХ
    work_sheet['F40'] = information_from_json['principal']['CEO']['name']
    if information_from_json['principal']['CEO']['is_name_changed'] is True:
        work_sheet['F41'] = 'ДА'
    else:
        work_sheet['F41'] = 'НЕТ'
    work_sheet['F42'] = information_from_json['principal']['CEO']['citizenship']
    work_sheet['F43'] = information_from_json['principal']['CEO']['birth_place']
    work_sheet['F44'] = information_from_json['principal']['CEO']['birth_date']
    work_sheet['F45'] = information_from_json['principal']['CEO']['snils']
    work_sheet['F46'] = information_from_json['principal']['CEO']['inn']
    work_sheet['D48'] = information_from_json['principal']['okved_main']
    work_sheet['D51'] = information_from_json['principal']['employees_number']
    #3.0. Сведения о заявке 
    work_sheet['D64'] = 'ПОТ Экспресс - гарантия'
    work_sheet['D65'] = information_from_json['guarantee_amount']
    work_sheet['D66'] = 'БГ по '+information_from_json['purchase_law']+' Ф3'
    if information_from_json['has_prepayment'] is True:
        work_sheet['D67'] = 'ДА'
    else:
        work_sheet['D67'] = 'НЕТ'
    if information_from_json['is_big_deal'] is True:
        work_sheet['D68'] = 'ДА'
    else:
        work_sheet['D68'] = 'НЕТ'
    work_sheet['D64'] = 'ПОТ Экспресс - гарантия'

    work_sheet['D69'] = 'Исполнение контракта'
    work_sheet['G64'] = information_from_json['guarantee_type_label']
    work_sheet['G65'] = information_from_json['guarantee_start_date']
    work_sheet['G66'] = information_from_json['guarantee_end_date']
    guarantee_amount = float(information_from_json['guarantee_amount'])
    purchase_price = float(information_from_json['purchase_starting_price'])
    percent = (guarantee_amount/purchase_price)*100
    fin_percent = str(round(percent,1))+'%'
    work_sheet['G67'] = fin_percent
    work_sheet['G69'] = 'Типовая'
    work_sheet['D73'] = information_from_json['beneficiary']['full_name']
    work_sheet['D74'] = information_from_json['beneficiary']['short_name']
    work_sheet['D75'] = information_from_json['beneficiary']['inn']
    work_sheet['D76'] = information_from_json['beneficiary']['kpp']
    work_sheet['D77'] = information_from_json['beneficiary']['ogrn']
    work_sheet['D79'] = information_from_json['beneficiary']['legal_address']['value']
    work_sheet['D81'] = information_from_json['purchase_number']
    work_sheet['D82'] = information_from_json['purchase_subject']
    work_sheet['D83'] = information_from_json['purchase_url']
    work_sheet['D85'] = information_from_json['purchase_starting_price']
    work_sheet['D86'] = '-'
    work_sheet['D87'] = '-'
    work_sheet['E88'] = '-'
    work_sheet['E89'] = '-'
    work_sheet['G88'] = '-'
    work_sheet['G89'] = '-'
    #БУХ ОТЧЕТ
    count_sheet = work_book.worksheets[1]
    count_sheet['B2'] = information_from_json['principal']['full_name']
    #Нематериальные активы_1110
    count_sheet['D8'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][0]['fields'][0]['values']['31.12.2017']
    count_sheet['E8'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][0]['fields'][0]['values']['31.03.2018']
    count_sheet['F8'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][0]['fields'][0]['values']['30.06.2018']
    count_sheet['G8'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][0]['fields'][0]['values']['30.09.2018']

    #Результаты исследований и разработок_1120
    count_sheet['D9'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][0]['fields'][1]['values']['31.12.2017']
    count_sheet['E9'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][0]['fields'][1]['values']['31.03.2018']
    count_sheet['F9'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][0]['fields'][1]['values']['30.06.2018']
    count_sheet['G9'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][0]['fields'][1]['values']['30.09.2018']

    #Нематериальные поисковые активы_1130
    count_sheet['D10'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][0]['fields'][2]['values']['31.12.2017']
    count_sheet['E10'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][0]['fields'][2]['values']['31.03.2018']
    count_sheet['F10'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][0]['fields'][2]['values']['30.06.2018']
    count_sheet['G10'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][0]['fields'][2]['values']['30.09.2018']

    #Запасы_1210
    count_sheet['D13'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][1]['fields'][0]['values']['31.12.2017']
    count_sheet['E13'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][1]['fields'][0]['values']['31.03.2018']
    count_sheet['F13'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][1]['fields'][0]['values']['30.06.2018']
    count_sheet['G13'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][1]['fields'][0]['values']['30.09.2018']
    
    #НДС по приобретенным ценностям 1220
    count_sheet['D14'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][1]['fields'][1]['values']['31.12.2017']
    count_sheet['E14'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][1]['fields'][1]['values']['31.03.2018']
    count_sheet['F14'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][1]['fields'][1]['values']['30.06.2018']
    count_sheet['G14'] = information_from_json['principal']['accountability']['data']['fieldsets'][0]['groups'][1]['fields'][1]['values']['30.09.2018']

    #Резервный капитал_1360
    count_sheet['D17'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][0]['fields'][4]['values']['31.12.2017']
    count_sheet['E17'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][0]['fields'][4]['values']['31.03.2018']
    count_sheet['F17'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][0]['fields'][4]['values']['30.06.2018']
    count_sheet['G17'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][0]['fields'][4]['values']['30.09.2018']

    #Нераспределенная прибыль (непокрытый убыток)_1370
    count_sheet['D18'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][0]['fields'][5]['values']['31.12.2017']
    count_sheet['E18'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][0]['fields'][5]['values']['31.03.2018']
    count_sheet['F18'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][0]['fields'][5]['values']['30.06.2018']
    count_sheet['G18'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][0]['fields'][5]['values']['30.09.2018']

    #Заёмные средства   1410
    count_sheet['D21'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][1]['fields'][0]['values']['31.12.2017']
    count_sheet['E21'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][1]['fields'][0]['values']['31.03.2018']
    count_sheet['F21'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][1]['fields'][0]['values']['30.06.2018']
    count_sheet['G21'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][1]['fields'][0]['values']['30.09.2018']
    
    #Отложенные налоговые обязательства_1420
    count_sheet['D22'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][1]['fields'][1]['values']['31.12.2017']
    count_sheet['E22'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][1]['fields'][1]['values']['31.03.2018']
    count_sheet['F22'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][1]['fields'][1]['values']['30.06.2018']
    count_sheet['G22'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][1]['fields'][1]['values']['30.09.2018']

    #Оценочные обязательства_1430
    count_sheet['D23'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][1]['fields'][2]['values']['31.12.2017']
    count_sheet['E23'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][1]['fields'][2]['values']['31.03.2018']
    count_sheet['F23'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][1]['fields'][2]['values']['30.06.2018']
    count_sheet['G23'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][1]['fields'][2]['values']['30.09.2018']

    #Прочие обязательства_1450
    count_sheet['D26'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][1]['fields'][3]['values']['31.12.2017']
    count_sheet['E26'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][1]['fields'][3]['values']['31.03.2018']
    count_sheet['F26'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][1]['fields'][3]['values']['30.06.2018']
    count_sheet['G26'] = information_from_json['principal']['accountability']['data']['fieldsets'][1]['groups'][1]['fields'][3]['values']['30.09.2018']

    #Выручка
    count_sheet['D32'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][0]['fields'][0]['values']['31.12.2017']
    count_sheet['E32'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][0]['fields'][0]['values']['31.03.2018']
    count_sheet['F32'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][0]['fields'][0]['values']['30.06.2018']
    count_sheet['G32'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][0]['fields'][0]['values']['30.09.2018']

    #Себестоимость продаж
    count_sheet['D33'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][0]['fields'][1]['values']['31.12.2017']
    count_sheet['E33'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][0]['fields'][1]['values']['31.03.2018']
    count_sheet['F33'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][0]['fields'][1]['values']['30.06.2018']
    count_sheet['G33'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][0]['fields'][1]['values']['30.09.2018']

     #Доходы от участия в других организациях_2310
    count_sheet['D42'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][0]['values']['31.12.2017']
    count_sheet['E42'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][0]['values']['31.03.2018']
    count_sheet['F42'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][0]['values']['30.06.2018']
    count_sheet['G42'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][0]['values']['30.09.2018']
    
    #Проценты к получению_2320
    count_sheet['D43'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][1]['values']['31.12.2017']
    count_sheet['E43'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][1]['values']['31.03.2018']
    count_sheet['F43'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][1]['values']['30.06.2018']
    count_sheet['G43'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][1]['values']['30.09.2018']
    
    #Проценты к уплате_2330
    count_sheet['D44'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][2]['values']['31.12.2017']
    count_sheet['E44'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][2]['values']['31.03.2018']
    count_sheet['F44'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][2]['values']['30.06.2018']
    count_sheet['G44'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][2]['values']['30.09.2018']

    #Прочие доходы_2340
    count_sheet['D46'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][3]['values']['31.12.2017']
    count_sheet['E46'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][3]['values']['31.03.2018']
    count_sheet['F46'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][3]['values']['30.06.2018']
    count_sheet['G46'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][3]['values']['30.09.2018']

    #Прочие доходы_2350
    count_sheet['D47'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][4]['values']['31.12.2017']
    count_sheet['E47'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][4]['values']['31.03.2018']
    count_sheet['F47'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][4]['values']['30.06.2018']
    count_sheet['G47'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][4]['values']['30.09.2018']

    #Текущий налог на прибыль_2410
    count_sheet['D50'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][6]['values']['31.12.2017']
    count_sheet['E50'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][6]['values']['31.03.2018']
    count_sheet['F50'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][6]['values']['30.06.2018']
    count_sheet['G50'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][6]['values']['30.09.2018']

    #Изменение отложенных налоговых обязательств_2430
    count_sheet['D52'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][8]['values']['31.12.2017']
    count_sheet['E52'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][8]['values']['31.03.2018']
    count_sheet['F52'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][8]['values']['30.06.2018']
    count_sheet['G52'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][8]['values']['30.09.2018']

    #Прочее_2460
    count_sheet['D55'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][10]['values']['31.12.2017']
    count_sheet['E55'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][10]['values']['31.03.2018']
    count_sheet['F55'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][10]['values']['30.06.2018']
    count_sheet['G55'] = information_from_json['principal']['accountability']['data']['fieldsets'][2]['groups'][1]['fields'][10]['values']['30.09.2018']


    #КОНЕЦ БУХ ОТЧЕТА



    work_book.save(my_file)


if __name__ == '__main__':
    file_name = get_args()
    information_from_json = load_json_data(filepath)
    write_infromation_into_file(file_name, information_from_json)