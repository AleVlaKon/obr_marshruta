import openpyxl as xl


def format_int_value(cell_value: int | float, nuli: str) -> str:
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
        

# workbook = xl.load_workbook('Ведомость.xlsx', data_only=True)
# sheet_names = [i for i in workbook.sheetnames if i not in ['Лист1', 'ИД', 'В обсл', 'аб1']]
# print(type(workbook['ИД']['A1']))

def test_format_int_value():
    assert format_int_value(1, '00') == '1,00'
    assert format_int_value(1.256, '00') == '1,256'
    assert format_int_value(1, '0') == '1,0'
    assert format_int_value(1.25678532, '00') == '1,257'


def test_dob_n():
    assert format_float_value(1, 3) == '1,000'
    assert format_float_value(1.2, 3) == '1,200'
    assert format_float_value(1.25, 3) == '1,250'
    assert format_float_value(1.256, 3) == '1,256'
    assert format_float_value(1, 2) == '1,00'
    assert format_float_value(1.2, 2) == '1,20'
    assert format_float_value(1.25, 2) == '1,25'
    assert format_float_value(1.256, 2) == '1,26'


test_format_int_value()
test_dob_n()