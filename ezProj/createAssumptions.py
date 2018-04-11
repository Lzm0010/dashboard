import xlwings as xw
import win32api
import opAssumptions
import helpers as h


def main():
    # workbook and variables
    wb = xw.Book.caller()
    sht = wb.sheets['BasicAssumptions']
    project_name = sht.range('B2').value
    projection_type = str(sht.range('B3').value).lower()

    # check for which projection
    if projection_type != 'o' and projection_type != 'b' and projection_type != 'd':
        win32api.MessageBox(wb.app.hwnd, "B3 is not valid, use o, d, or b")

    if projection_type == 'o' or projection_type == 'b':
        h.create_new_sheet(wb, 'OperatingAssumptions', 'BasicAssumptions')
        opAssumptions.create_ops(wb)

    if projection_type == 'd' or projection_type == 'b':
        h.create_new_sheet(wb, 'DevelopmentAssumptions', 'BasicAssumptions')

    # save workbook
    wb.save("C:/Users/leem/Desktop/dealflow/ezprojections/ezProj/" + project_name + ".xlsm")
