from openpyxl.worksheet.worksheet import Worksheet
from osnov_vid_def import vivodi_v_otchet


def format_int_value(cell_value: int | float, nuli: str='0') -> str:
    #Добавляет nuli, если cell целое число
    if isinstance(cell_value, int):                      
        return f'{cell_value},{nuli}'
    else:
        return str(round(cell_value, 3)).replace('.', ',')


def format_float_value(number: int | float, decimal_places: int) -> str:
    '''
    Добавляет хвостовые нули в протяженности 
    в зависимости от количества нулей после запятой,
    чтобы было 0,000 формат км
    '''
    return f"{number:,.{decimal_places}f}".replace('.', ',')


def convert_str_to_float(value: str | float | int) -> float | int:
    '''Если в ячейке сохранена строка, переводит ее в число'''
    if isinstance(value, str):
        return float(value.replace(',', '.'))
    return value


def format_str_or_digit_value(cell_value: int | float | str, decimal_places: int) -> str:
     if isinstance(cell_value, str):
          return cell_value.replace('.', ',')
     return format_float_value(cell_value, decimal_places)


def format_km_with_plus_values(start_km: int | float, end_km: int | float) -> str:
     format_start_km = format_float_value(start_km, 3).replace(',', '+')
     format_end_km = format_float_value(end_km, 3).replace(',', '+')
     return f'км {format_start_km} - км {format_end_km}'

        
def return_base_context(sheet: Worksheet) -> dict:
        context = {
        'number': sheet['B1'].value, 
        'name': sheet['C1'].value,   
        'opisanie': sheet['AM6'].value,
        'shirina': format_str_or_digit_value(sheet['B6'].value, 1),
        'categoria': sheet['E3'].value,
        'protyazhennost': format_int_value(sheet['B4'].value, '0'),
        'prinadlezhnost': sheet['B7'].value,
        'tip_pokr': sheet['B5'].value,
        'osn_vid_def': vivodi_v_otchet(sheet)[1],
        }
        return context


def change_table_2(table_2: list):
    ''' Редактирует таблицу 2 (замена . на , и добавление хвостовых нулей)'''
    # {'km_nach': 45, 'km_kon': 46, 'pokr_i': 47, 'shir_i': 48, 'ball_i': 49,}
    for row in table_2:
        row['km_nach'] = format_float_value(row['km_nach'], 3)
        row['km_kon'] = format_float_value(row['km_kon'], 3)
        row['shir_i'] = format_str_or_digit_value(row['shir_i'], 1)
        row['ball_i'] = format_str_or_digit_value(row['ball_i'], 1)
    

def change_table_3(table_3: list):
    ''' Редактирует таблицу 3 (замена . на , и добавление хвостовых нулей'''
    # {'km': 51, 'ball_i': 52, 'kpr_i': 53, }
    for row in table_3:
        row['ball_i'] = format_str_or_digit_value(row['ball_i'], 1)
        if row['kpr_i'] == 0.5:
            row['kpr_i'] == '-'
        else:
            row['kpr_i'] = format_float_value(row['kpr_i'], 2)


def change_table_4(table_4: list):
    ''' Редактирует таблицу 4 (замена . на , и добавление хвостовых нулей'''
    # {'km': 56, 'kpr_i': 57, 'E_i': 58, }
    for row in table_4:
        row['kpr_i'] = format_float_value(row['kpr_i'], 2)
        row['E_i'] = format_str_or_digit_value(row['E_i'], 0)

            
def return_if_error_value(cell_value: int | float | str, decimal_places: int) -> str:
     if isinstance(cell_value, str):
          return cell_value
     return f"{cell_value:,.{decimal_places}f}".replace('.', ',')

        
# workbook = xl.load_workbook('Ведомость тест.xlsx', data_only=True)
# sheet_names = [i for i in workbook.sheetnames if i not in ['Лист1', 'ИД', 'В обсл', 'аб1']]
# sheet_1: Worksheet = workbook['У 1']
# print(type(sheet_1['B5']))

# print(return_base_context(sheet_1))
