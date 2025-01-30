
import openpyxl as xl




def format_cell_value(cell_value: int | float, nuli: str) -> str:
    #Добавляет nuli, если cell целое число
    if isinstance(cell_value, int):                      
        return f'{cell_value},{nuli}'
    else:
        return str(cell_value).replace('.', ',')
    

def dob_n(number: int | float) -> str:
    '''
    Добавляет хвостовые нули в протяженности 
    в зависимости от количества нулей после запятой,
    чтобы было 0,000 формат км
    '''
    return f"{number:,.3f}".replace('.', ',')
        

# workbook = xl.load_workbook('Ведомость.xlsx', data_only=True)
# sheet_names = [i for i in workbook.sheetnames if i not in ['Лист1', 'ИД', 'В обсл', 'аб1']]
# print(type(workbook['ИД']['A1']))

def test_format_cell_value():
    pass


def test_dob_n():
    assert dob_n(1) == '1,000'
    assert dob_n(1.2) == '1,200'
    assert dob_n(1.25) == '1,250'
    assert dob_n(1.256) == '1,256'

test_dob_n()