import pytest
import openpyxl as xl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell
from osnov_vid_def import vivodi_v_otchet


def format_int_value(cell_value: int | float, nuli: str) -> str:
    #Добавляет nuli, если cell целое число
    if isinstance(cell_value, int):                      
        return f'{cell_value},{nuli}'
    else:
        return str(round(cell_value, 3)).replace('.', ',')
    
def format_shirina(cell_value: int | float | str) -> str:
     if isinstance(cell_value, str):
          return cell_value.replace('.', ',')
     return format_int_value(cell_value, '0')


def format_float_value(number: int | float, decimal_places: int) -> str:
    '''
    Добавляет хвостовые нули в протяженности 
    в зависимости от количества нулей после запятой,
    чтобы было 0,000 формат км
    '''
    return f"{number:,.{decimal_places}f}".replace('.', ',')


def format_km_with_plus_values(start_km: int | float, end_km: int | float) -> str:
     format_start_km = format_float_value(start_km, 3).replace(',', '+')
     format_end_km = format_float_value(end_km, 3).replace(',', '+')
     return f'км {format_start_km} - км {format_end_km}'

        

def return_base_context(sheet: Worksheet) -> dict:
        context = {
        'number': sheet['B1'].value,
        'name': sheet['C1'].value,
        'opisanie': sheet['AM6'].value,
        'shirina': sheet['B6'].value,
        'categoria': sheet['E3'].value,
        'protyazhennost': format_int_value(sheet['B4'].value, '0'),
        'prinadlezhnost': sheet['B7'].value,
        'tip_pokr': sheet['B5'].value,
        }
        return context



workbook = xl.load_workbook('Ведомость тест.xlsx', data_only=True)
sheet_names = [i for i in workbook.sheetnames if i not in ['Лист1', 'ИД', 'В обсл', 'аб1']]
sheet_1: Worksheet = workbook['У 1']
print(type(sheet_1['B5']))

print(return_base_context(sheet_1))
