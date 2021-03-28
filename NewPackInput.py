#! python3
# NewPackInput.py - input Label, product and date and will add to google sheets
from products import NewProduct
new_product = NewProduct()
new_product.data_input()
new_product.open_workbook()
while True:
    if new_product.date == '':
        new_product.add_product_to_workbook_date_not_known()
    else:
        new_product.add_product_to_workbook_date_known()
    if new_product.are_you_finished() == 'y':
        break
    else:
        new_product.data_input()
new_product.close_workbook()



