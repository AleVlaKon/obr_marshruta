import pytest
from refactor import format_float_value, format_int_value, format_km_with_plus_values
from refactor import return_base_context, format_shirina
import openpyxl as xl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell
from refactor import change_table_2, change_table_3, change_table_4



def test_format_int_value():
    assert format_int_value(1, '00') == '1,00'
    assert format_int_value(1.256, '00') == '1,256'
    assert format_int_value(1, '0') == '1,0'
    assert format_int_value(1.25678532, '00') == '1,257'


def test_format_float_value():
    assert format_float_value(1, 3) == '1,000'
    assert format_float_value(1.2, 3) == '1,200'
    assert format_float_value(1.25, 3) == '1,250'
    assert format_float_value(1.256, 3) == '1,256'
    assert format_float_value(1, 2) == '1,00'
    assert format_float_value(1.2, 2) == '1,20'
    assert format_float_value(1.25, 2) == '1,25'
    assert format_float_value(1.256, 2) == '1,26'
    assert format_float_value(1.25678934, 2) == '1,26'
    assert format_float_value(1.25678934, 3) == '1,257'
    assert format_float_value(1.25678934, 0) == '1'


def test_format_km_with_plus_values():
    assert format_km_with_plus_values(1, 3) == 'км 1+000 - км 3+000'
    assert format_km_with_plus_values(1.25, 3.05) == 'км 1+250 - км 3+050'
    assert format_km_with_plus_values(1.256, 3.053) == 'км 1+256 - км 3+053'
    assert format_km_with_plus_values(0, 0) == 'км 0+000 - км 0+000'


def test_format_shirina():
    assert format_shirina(1) == '1,0'
    assert format_shirina(7.54) == '7,54'
    assert format_shirina(7.5) == '7,5'
    assert format_shirina('7.5-15.0') == '7,5-15,0'


@pytest.fixture
def workbook():
    return xl.load_workbook('Ведомость тест.xlsx', data_only=True)


@pytest.fixture
def sheet(workbook):
    return workbook['У 1']


def test_base_context(sheet):
    context_from_file = return_base_context(sheet)
    request_context = {
        'number': 'У 1',
        'name': 'А-380 «Гузор-Бухоро-Нукус-Бейнеу»   км 565+929 - км 690+658',
        'opisanie': 'Маршрут № У 1 представляет собой участок федеральной автомобильной дороги «А-380 «Гузор-Бухоро-Нукус-Бейнеу»   км 565+929 - км 690+658», который начинается от автомобильной дороги "Подъезд к ст. Мискин" и следует до автомобильной дороги "Подъезд к ст. Караузяк".',
        'shirina': '7,0-15,0',
        'categoria': 'III',
        'protyazhennost': '122,2',
        'prinadlezhnost': 'федеральная',
        'tip_pokr': 'асф. бет., цементобетон',
        'osn_vid_def': 'Основные виды дефектов – сетка трещин.',
        }
    assert context_from_file == request_context

@pytest.fixture
def table_2():
    context_table_2 = [
        {'km_nach': 0.020, 'km_kon': 0.121, 'pokr_i': 'ЩПС', 'shir_i': 6.5, 'ball_i': 2,},
        {'km_nach': 1.000, 'km_kon': 2.001, 'pokr_i': 'ЩПС', 'shir_i': "7,5х2", 'ball_i': 2.5,},
        {'km_nach': 123.020, 'km_kon': 123.000, 'pokr_i': 'ЩПС', 'shir_i': 6, 'ball_i': 2.56,},
    ]
    return context_table_2

def test_change_table_2(table_2):
    result_table = [
        {'km_nach': '0,020', 'km_kon': "0,121", 'pokr_i': 'ЩПС', 'shir_i': '6,5', 'ball_i': '2,0',},
        {'km_nach': "1,000", 'km_kon': "2,001", 'pokr_i': 'ЩПС', 'shir_i': "7,5х2", 'ball_i': '2,5',},
        {'km_nach': '123,020', 'km_kon': '123,000', 'pokr_i': 'ЩПС', 'shir_i': '6,0', 'ball_i': '2,6',},
    ]
    change_table_2(table_2)

    assert table_2 == result_table


    #test table_3
@pytest.fixture
def table_3():
    context_table_3 = [
        {'km': 'км 1+000 км 2+000', 'ball_i': 3.121, 'kpr_i': 1},
        {'km': 'км 2+000 км 2+350', 'ball_i': 3, 'kpr_i': 0.8},
        {'km': 'км 2+350 км 3+000', 'ball_i': 3.2, 'kpr_i': 0.8333333333},
    ]   
    return context_table_3

def test_change_table_3(table_3):
    result_table = [
        {'km': 'км 1+000 км 2+000', 'ball_i': '3,1', 'kpr_i': '1,00'},
        {'km': 'км 2+000 км 2+350', 'ball_i': '3,0', 'kpr_i': '0,80'},
        {'km': 'км 2+350 км 3+000', 'ball_i': '3,2', 'kpr_i': '0,83'},
    ]
    change_table_3(table_3)

    assert table_3 == result_table


#test table_4
@pytest.fixture
def table_4():
    context_table_4 = [
        {'km': 'км 1+000 км 2+000', 'kpr_i': 1, 'E_i': 132},
        {'km': 'км 2+000 км 2+350', 'kpr_i': 0.8, 'E_i': 132.4},
        {'km': 'км 2+350 км 3+000', 'kpr_i': 0.8333333333, 'E_i': '#ЗНАЧ'},
    ]   
    return context_table_4

def test_change_table_4(table_4):
    result_table = [
        {'km': 'км 1+000 км 2+000', 'kpr_i': '1,00', 'E_i': '132'},
        {'km': 'км 2+000 км 2+350', 'kpr_i': '0,80', 'E_i': '132'},
        {'km': 'км 2+350 км 3+000', 'kpr_i': '0,83', 'E_i': '#ЗНАЧ'},
    ]
    change_table_4(table_4)

    assert table_4 == result_table