import requests
from bs4 import BeautifulSoup
from Product import Product
from openpyxl import Workbook
from datetime import datetime

#주소 만들기
def make_address_moma(productcode, color, size):
    return "https://store.moma.org/on/demandware.store/Sites-moma-Site/en_US/Product-Variation?pid="+productcode+"&dwvar_"+productcode+"_size="+size+"&dwvar_"+productcode+"_color=" + color

#재고파악
def view_status(address, classvalue):
    result = 'X'
    webpage = requests.get(address)
    soup = BeautifulSoup(webpage.content, "html.parser")
    status = soup.find_all(attrs={'class': classvalue})
    for unit in status:
        if unit.get_text().find('Add to Cart') != -1:
            result = 'O'
            break
    return result

#엑셀에 등록
def view_product(product, excel):
    for _color in product.colorlist:
        for _size in product.sizelist:
            _result = view_status(make_address_moma(
                product.code, _color, _size), _moma_class_value)
            product.resultlist.append(_result)
            excel.append([product.productname, _color, _size, _result])
            
            print('상품: ' + product.productname + '\t색상: ' +
                  _color + '\t사이즈: ' + _size + '\t재고유무: ' + _result)

#찾을 class 이름
_moma_class_value = 'product-content'

# MoMA 후드
_product_code_hoodie = '8694'
_size_list_hoodie = ['Small', 'Medium', 'Large', 'X-Large']
_color_list_hoodie = ['Navy', 'Gray', 'Black']
_result_list_hoodie = []

# MoMA 맨투맨
_product_code_shirts = '400613'
_size_list_shirts = ['Small', 'Medium', 'Large', 'X-Large']
_color_list_shirts = ['Gray']
_result_list_shirts = []

#후드 객체 생성
_product_hoodie = Product('Champion Hoodie - MoMA Edition', _product_code_hoodie,
                          _color_list_hoodie, _size_list_hoodie, _result_list_hoodie)

#맨투맨 객체 생성
_product_shirts = Product('Champion Crewneck Sweatshirt - MoMA Edition', _product_code_shirts,
                          _color_list_shirts, _size_list_shirts, _result_list_shirts)

#엑셀 생성
write_wb = Workbook()
write_ws = write_wb.create_sheet('재고관리')

#내용 추가
write_ws.append(['product', 'color', 'size', 'instock'])
view_product(_product_hoodie, write_ws)
view_product(_product_shirts, write_ws)

#엑셀 저장
write_wb.save(datetime.today().strftime("%Y-%m-%d") + ' 재고.xlsx')
print("엑셀 저장 완료")
