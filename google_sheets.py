import gspread
import pandas as pd

sa = gspread.service_account(filename='excel-391118-14e98763ea19.json')

def get_sheet_values(sheet_name):
    sh = sa.open('Артикулы')
    wks = sh.worksheet(sheet_name)
    return [item[0] for item in wks.get_all_values()]

def get_wb_articles():
    return get_sheet_values('Артикулы Вб')

def get_ozon_articles():
    return get_sheet_values('Артикулы ОЗОН')


def get_guide():
    # Чтения данных со справочника
    sh = sa.open('Справочник')
        
    wks = sh.worksheet('Лист1')
    data = wks.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    
    return df
       