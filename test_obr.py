import pytest
from refactor import format_float_value, format_int_value


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


test_format_int_value()
test_format_float_value()

@pytest.fixture
def return_workbook():
    return 2


def test_data(return_workbook):
    assert return_workbook + 1 == 3