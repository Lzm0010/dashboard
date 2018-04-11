import xlwings as xw
import helpers as h


# def makeDataFrames(wb, sht):
#     # make data frames for each section
#     def make_range(col):
#         letters = 'ABCDEFGHIJKLMNOPQRST'
#         lrow = h.lastRow('OperatingAssumptions', wb, col)
#         sec = sht.range(letters[(col-1):col] + '3:' + letters[col:(col+1)] + str(lrow)).options(pd.DataFrame, index=False, header=False).value
#         sec.columns = [['names', 'values']]
#         return sec
#
#     return (make_range(1), make_range(4), make_range(7),
#             make_range(10), make_range(13), make_range(16),
#             make_range(19))


def create_inv_for_products(wb, sht, products, product):

    for i in range(1, int(products + 1)):
        next_row = h.lastRow('Inventory', wb, 5) + 2
        sht.range('B' + str(next_row)).value = '=OperatingAssumptions!$E' + str(i * (len(product.rows) + 1) - (len(product.rows) - 3))
        sht.range('D' + str(next_row)).value = [['Starting Units'], ['Units Sold'], ['Units Added'], ['$ Inv Added'], ['Remaining']]
        # sht.range('E' + str(next_row)).value = '=OperatingAssumptions!$E' + str(i * (len(product.rows) + 1) - (len(product.rows) - 22))
        sht.range('E' + str(next_row)).value = 0
        sht.range('F' + str(next_row) + ':P' + str(next_row)).value = '=E' + str(next_row + 4)
        sht.range('E' + str(next_row + 1) + ':P' + str(next_row + 1)).value = '=units_sold_channel_method(E2, All_Channels, Product' + str(i) + ', All_Products)'
        sht.range('E' + str(next_row + 2) + ':P' + str(next_row + 2)).value = '=units_added(E' + str(next_row) + ', E' + str(next_row + 1) + ', OperatingAssumptions!$E' + str(i * (len(product.rows) + 1) - (len(product.rows) - 23)) + ', OperatingAssumptions!$E' + str(i * (len(product.rows) + 1) - (len(product.rows) - 22)) + ')'
        sht.range('E' + str(next_row + 2) + ':P' + str(next_row + 2)).name = 'InvAdded'
        sht.range('E' + str(next_row + 3) + ':P' + str(next_row + 3)).value = '=inv_cost_added(E2, InvAdded, Product' + str(i) + ')'
        sht.range('E' + str(next_row + 4) + ':P' + str(next_row + 4)).value = '=remaining_units(E' + str(next_row) + ', E' + str(next_row + 1) + ', E' + str(next_row + 2) + ')'

    sht.range('E4:P' + str(next_row + 4)).name = 'AllInv'
    sht.range('D' + str(next_row + 6)).value = "Total Inventory Cost Added"
    sht.range('E' + str(next_row + 6) + ':P' + str(next_row + 6)).value = '=total_inventory_cost_added(E2, AllInv, ' + str(products) + ')'
    sht.range('E' + str(next_row + 6) + ':P' + str(next_row + 6)).name = 'TotalInvAddedRow'


def create_sales_for_products(wb, sht, products, product):

    for i in range(1, int(products + 1)):
        next_row = h.lastRow('SalesDetail', wb, 5) + 2
        sht.range('B' + str(next_row)).value = '=OperatingAssumptions!$E' + str(i * (len(product.rows) + 1) - (len(product.rows) - 3))
        sht.range('D' + str(next_row)).value = [['Retail Price'], ['WS Price'], ['Total Sales'], ['Retail Units'], ['WS Units']]
        sht.range('E' + str(next_row) + ':P' + str(next_row)).value = '=OperatingAssumptions!$E' + str(i * (len(product.rows) + 1) - (len(product.rows) - 5))
        sht.range('E' + str(next_row + 1) + ':P' + str(next_row + 1)).value = '=get_ws_price(E' + str(next_row) + ', OperatingAssumptions!$E' + str(i * (len(product.rows) + 1) - (len(product.rows) - 7)) + ')'
        sht.range('E' + str(next_row + 2) + ':P' + str(next_row + 2)).value = '=total_sales(E' + str(next_row) + ', E' + str(next_row + 1) + ', E' + str(next_row + 3) + ', E' + str(next_row + 4) + ')'
        sht.range('E' + str(next_row + 4) + ':P' + str(next_row + 4)).value = '=wholesale_units(E2, Wholesale, Product' + str(i) + ', All_Products)'
        sht.range('E' + str(next_row + 3) + ':P' + str(next_row + 3)).value = '=retail_units(Inventory!E' + str(next_row + 1) + ', E' + str(next_row + 4) + ')'

    last_row = h.lastRow('SalesDetail', wb, 5)
    sht.range('E4:P' + str(last_row)).name = 'AllSales'
    sht.range('D' + str(last_row + 2)).value = 'Total Retail Sales'
    sht.range('D' + str(last_row + 3)).value = 'Total WS Sales'
    sht.range('D' + str(last_row + 4)).value = 'Total All Sales'
    sht.range('E' + str(last_row + 2) + ':P' + str(last_row + 2)).value = '=all_products_retail_sales(E2, AllSales,' + str(products) + ')'
    sht.range('E' + str(last_row + 3) + ':P' + str(last_row + 3)).value = '=all_products_ws_sales(E2, AllSales,' + str(products) + ')'
    sht.range('E' + str(last_row + 4) + ':P' + str(last_row + 4)).value = '=SUM(E'+ str(last_row + 2) + ':E' + str(last_row + 3) + ')'
    sht.range('E' + str(last_row + 4) + ':P' + str(last_row + 4)).name = 'AllProductSales'


def create_cost_for_products(wb, sht, products, product):

    for i in range(1, int(products + 1)):
        next_row = h.lastRow('CostDetail', wb, 5) + 2
        sht.range('B' + str(next_row)).value = '=OperatingAssumptions!$E' + str(i * (len(product.rows) + 1) - (len(product.rows) - 3))
        sht.range('D' + str(next_row)).value = [['Units Ordered'], ['Discount'], ['Monthly Units'], ['Monthly Cost']]
        sht.range('E' + str(next_row) + ':P' + str(next_row)).value = '=Inventory!E' + str(next_row + 2 + (i - 1))
        sht.range('E' + str(next_row + 1) + ':P' + str(next_row + 1)).value = 0
        sht.range('E' + str(next_row + 2) + ':P' + str(next_row + 2)).value = '=Inventory!E' + str(next_row + 1 + (i - 1))
        sht.range('E' + str(next_row + 3) + ':P' + str(next_row + 3)).value = '=monthly_cost_per_product(Product' + str(i) + ', E' + str(next_row + 2) + ')'

    last_row = h.lastRow('CostDetail', wb, 5)
    sht.range('E4:P' + str(last_row)).name = 'AllCost'
    sht.range('D' + str(last_row + 2)).value = 'Total Cost'
    sht.range('E' + str(last_row + 2) + ':P' + str(last_row + 2)).value = '=all_products_cost(E2, AllCost,' + str(products) + ')'
    sht.range('E' + str(last_row + 2) + ':P' + str(last_row + 2)).name = 'AllProductCost'


def create_employees(wb, sht, sales, rnd, oh):

    for i in range(1, int(sales + 1)):
        next_row = h.lastRow('Employees', wb, 5) + 1
        sht.range('C' + str(next_row)).value = '=OperatingAssumptions!$H' + str((i * 6) - 2)
        sht.range('E' + str(next_row) + ':P' + str(next_row)).value = '=employee_cost(E2, General_Sales_and_Marketing' + str(i) + ')'

    for i in range(int(sales + 1), int(sales + rnd + 1)):
        next_row = h.lastRow('Employees', wb, 5) + 1
        sht.range('C' + str(next_row)).value = '=OperatingAssumptions!$H' + str((i * 6) - 1)
        sht.range('E' + str(next_row) + ':P' + str(next_row)).value = '=employee_cost(E2, Research_and_Development' + str(i - int(sales)) + ')'

    for i in range(int(sales + rnd + 1), int(sales + rnd + oh + 1)):
        next_row = h.lastRow('Employees', wb, 5) + 1
        sht.range('C' + str(next_row)).value = '=OperatingAssumptions!$H' + str(i * 6)
        sht.range('E' + str(next_row) + ':P' + str(next_row)).value = '=employee_cost(E2, Overhead' + str(i - int(sales) - int(rnd)) + ')'

    sht.range('C' + str(next_row + 2)).value = [['Full Time']]
    sht.range('E' + str(next_row + 2) + ':P' + str(next_row + 2)).value = '=SUM(E' + str(next_row - int(sales) - int(rnd) - int(oh) + 1) + ':E' + str(next_row) + ')'
    sht.range('C' + str(next_row + 3)).value = [['Taxes & Benefits']]
    sht.range('E' + str(next_row + 3) + ':P' + str(next_row + 3)).value = '=tax_and_benefits(E' + str(next_row + 2) + ', EmployeeBonus)'
    sht.range('C' + str(next_row + 4)).value = [['Total Employee Cost']]
    sht.range('E' + str(next_row + 4) + ':P' + str(next_row + 4)).value = '=SUM(E' + str(next_row + 2) + ':E' + str(next_row + 3) + ')'
    sht.range('C' + str(next_row + 2) + ':P' + str(next_row + 4)).name = 'TotalEmployeeCost'


def create_assets(wb, sht, tang_assets, intang_assets):

    sht.range('A3').value = [['Depreciation']]
    for i in range(1, int(tang_assets + 1)):
        if i == 1:
            next_row = h.lastRow('AssetSchedule', wb, 5) + 2
        else:
            next_row = h.lastRow('AssetSchedule', wb, 5) + 1
        sht.range('D' + str(next_row)).value = '=OperatingAssumptions!$Q' + str(i * 4)
        sht.range('E' + str(next_row) + ':P' + str(next_row)).value = '=asset_monthly_cost(E2, OperatingAssumptions!$Q' + str((i * 4) + 1) + ', OperatingAssumptions!$Q' + str((i * 4) + 2) + ', OperatingAssumptions!$Q' + str((i * 4) + 3) + ')'

    next_row = h.lastRow('AssetSchedule', wb, 5) + 1
    sht.range('D' + str(next_row)).value = [['Total']]
    sht.range('E' + str(next_row) + ':P' + str(next_row)).value = '=SUM(E' + str(next_row - int(tang_assets)) + ':E' + str(next_row - 1) + ')'
    sht.range('E' + str(next_row) + ':P' + str(next_row)).name = 'Depreciation'


    sht.range('A' + str(next_row + 2)).value = [['Amortization']]
    for i in range(int(tang_assets + 1), int(intang_assets + tang_assets + 1)):
        if i == int(tang_assets + 1):
            next_row = h.lastRow('AssetSchedule', wb, 5) + 3
        else:
            next_row = h.lastRow('AssetSchedule', wb, 5) + 1
        sht.range('D' + str(next_row)).value = '=OperatingAssumptions!$Q' + str((i * 4) + 1)
        sht.range('E' + str(next_row) + ':P' + str(next_row)).value = '=asset_monthly_cost(E2, OperatingAssumptions!$Q' + str((i * 4) + 2) + ', OperatingAssumptions!$Q' + str((i * 4) + 3) + ', OperatingAssumptions!$Q' + str((i * 4) + 4) + ')'

    next_row = h.lastRow('AssetSchedule', wb, 5) + 1
    sht.range('D' + str(next_row)).value = [['Total']]
    sht.range('E' + str(next_row) + ':P' + str(next_row)).value = '=SUM(E' + str(next_row - int(intang_assets)) + ':E' + str(next_row - 1) + ')'
    sht.range('E' + str(next_row) + ':P' + str(next_row)).name = 'Amortization'


def create_loan(wb, sht, loan, seed_rds):
    seed_rds = int(seed_rds)
    for i in range(0, (int(loan[3][1]) * 12) + 1):
        if i == 0:
            sht.range('A5').value = i
            if seed_rds == 0:
                sht.range('E5').value = '=OperatingAssumptions!$T4 - (OperatingAssumptions!$T4 * OperatingAssumptions!$T6)'
            else:
                sht.range('E5').value = '=OperatingAssumptions!$T' + str((seed_rds * 4) + 5) + ' - (OperatingAssumptions!$T' + str((seed_rds * 4) + 5) + ' * OperatingAssumptions!$T' + str((seed_rds * 4) + 7) + ')'
        else:
            sht.range('A' + str(i + 5)).value = i
            if seed_rds == 0:
                sht.range('F' + str(i + 5)).value = '=OperatingAssumptions!$T5'
                sht.range('B' + str(i + 5)).value = '=-PMT(F' + str(i + 5) + '/12, (OperatingAssumptions!$T7*12), E' + str(i + 4) + ')'
            else:
                sht.range('F' + str(i + 5)).value = '=OperatingAssumptions!$T' + str((seed_rds * 4) + 6)
                sht.range('B' + str(i + 5)).value = '=-PMT(F' + str(i + 5) + '/12, (OperatingAssumptions!$T' + str((seed_rds * 4) + 8) + '*12) - A' + str(i + 4) + ', E' + str(i + 4) + ')'

            sht.range('C' + str(i + 5)).value = '=ROUND(E' + str(i + 4) + '*F' + str(i + 5) + '/12, 2)'
            sht.range('D' + str(i + 5)).value = '=B' + str(i + 5) + '-C' + str(i + 5)
            sht.range('E' + str(i + 5)).value = '=E' + str(i + 4) + '-D' + str(i + 5)

    last_row = h.lastRow('LoanSchedule', wb, 3)
    sht.range('C6:C' + str(last_row)).name = 'Interest'
    sht.range('E5:E' + str(last_row)).name = 'LoanBalance'


def create_pl(wb, sht, expenses):
    shtemp = wb.sheets['Employees']

    sht.range('B4').value = 'Total Income'
    sht.range('E4:P4').value = '=total_income(E2, AllProductSales)'
    sht.range('B6').value = 'Total Cost of Sales'
    sht.range('E6:P6').value = '=total_cost(E2, AllProductCost)'
    sht.range('B8').value = 'Gross Margin'
    sht.range('E8:P8').value = '=E4 - E6'
    sht.range('B10').value = 'Fixed Business Expenses'
    sht.range('C11:P13').value = shtemp.range('TotalEmployeeCost').value

    for i in range(0, len(expenses)):
        sht.range('C'+ str(i + 14)).value = expenses[i][0]
        sht.range('E' + str(i + 14) + ':P' + str(i + 14)).value = expenses[i][1]

    next_row = h.lastRow('P&L', wb, 5) + 1
    sht.range('E14:P' + str(next_row - 1)).name = 'FixedExpPL'
    sht.range('B' + str(next_row)).value = 'Total Fixed Business Expenses'
    sht.range('E' + str(next_row) + ':P' + str(next_row)).value = '=SUM(E13:E' + str(next_row - 1) + ')'
    sht.range('B' + str(next_row + 2)).value = 'EBITDA'
    sht.range('E' + str(next_row + 2) + ':P' + str(next_row + 2)).value = '=E8 - E' + str(next_row)
    sht.range('E' + str(next_row + 2) + ':P' + str(next_row + 2)).name = 'EBITDA'
    sht.range('B' + str(next_row + 4)).value = 'Other Expenses'
    sht.range('C' + str(next_row + 5)).value = 'Amortization'
    sht.range('E' + str(next_row + 5) + ':P' + str(next_row + 5)).value = '=amortization_amt(E2, Amortization)'
    sht.range('C' + str(next_row + 6)).value = 'Depreciation'
    sht.range('E' + str(next_row + 6) + ':P' + str(next_row + 6)).value = '=depr_amt(E2, Depreciation)'
    sht.range('C' + str(next_row + 7)).value = 'Prefered Return'
    # sht.range('E' + str(next_row + 7) + ':P' + str(next_row + 7)).value = '='
    sht.range('C' + str(next_row + 8)).value = 'Interest'
    sht.range('E' + str(next_row + 8) + ':P' + str(next_row + 8)).value = '=int_amt(E2, Interest)'
    sht.range('C' + str(next_row + 9)).value = 'Tax'
    # sht.range('E' + str(next_row + 9) + ':P' + str(next_row + 9)).value = '='
    sht.range('B' + str(next_row + 10)).value = 'Total Other Expenses'
    sht.range('E' + str(next_row + 10) + ':P' + str(next_row + 10)).value = '=SUM(E' + str(next_row + 5) + ':E' + str(next_row + 9) + ')'
    sht.range('B' + str(next_row + 12)).value = 'Net Income'
    sht.range('E' + str(next_row + 12) + ':P' + str(next_row + 12)).value = '=E' + str(next_row + 2) + ' - E' + str(next_row + 10)


def create_bs(wb, sht, intang_assets, tang_assets, seed_rds):
    sht.range('A3').value = 'Assets'
    sht.range('B4').value = 'Current Assets'
    sht.range('C5').value = 'Cash'
    sht.range('C6').value = 'Accounts Receivable'
    sht.range('E6:P6').value = "=OperatingAssumptions!$N$10/30 * 'P&L'!E4"
    sht.range('C7').value = 'Raw Materials'
    sht.range('C8').value = 'Inventory'
    sht.range('E8:P8').value = "=inventory(E2, 'P&L'!E6, TotalInvAddedRow, D8)"
    sht.range('C9').value = 'Prepaid Expenses'
    sht.range('C10').value = 'Other Current'
    sht.range('B11').value = 'Total Current Assets'
    sht.range('E11:P11').value = '=SUM(E5:E10)'
    sht.range('B13').value = 'Intangible Assets'
    for i in range(int(tang_assets + 1), int(intang_assets + tang_assets + 1)):
        if i == int(tang_assets + 1):
            next_row = h.lastRow('BS', wb, 5) + 3
        else:
            next_row = h.lastRow('BS', wb, 5) + 1
        sht.range('C' + str(next_row)).value = '=OperatingAssumptions!$Q' + str((i * 4) + 1)
        sht.range('E' + str(next_row) + ':P' + str(next_row)).value = '=OperatingAssumptions!$Q' + str((i * 4) + 3)

    sht.range('B' + str(next_row + 1)).value = 'Total Intangible Assets'
    sht.range('E' + str(next_row + 1) + ':P' + str(next_row + 1)).value = '=SUM(E13:E' + str(next_row) + ')'
    sht.range('E' + str(next_row + 1) + ':P' + str(next_row + 1)).name = 'IntangibleAssets'
    sht.range('B' + str(next_row + 3)).value = 'Less: Accumulated Amortization'
    sht.range('E' + str(next_row + 3)).value = '=-amortization_amt(E2, Amortization)'
    sht.range('F' + str(next_row + 3) + ':P' + str(next_row + 3)).value = '= E' + str(next_row + 3) + ' - amortization_amt(F2, Amortization)'
    sht.range('E' + str(next_row + 3) + ':P' + str(next_row + 3)).name = 'AccumAmortization'

    sht.range('B' + str(next_row + 5)).value = 'Fixed Assets'
    for i in range(1, int(tang_assets + 1)):
        if i == 1:
            next_row = h.lastRow('BS', wb, 5) + 3
        else:
            next_row = h.lastRow('BS', wb, 5) + 1
        sht.range('C' + str(next_row)).value = '=OperatingAssumptions!$Q' + str(i * 4)
        sht.range('E' + str(next_row) + ':P' + str(next_row)).value = '=OperatingAssumptions!$Q' + str((i * 4) + 2)

    sht.range('B' + str(next_row + 1)).value = 'Total Fixed Assets'
    sht.range('E' + str(next_row + 1) + ':P' + str(next_row + 1)).value = '=SUM(E' + str(next_row - int(tang_assets)) + ':E' + str(next_row) + ')'
    sht.range('E' + str(next_row + 1) + ':P' + str(next_row + 1)).name = 'FixedAssetRow'
    sht.range('B' + str(next_row + 3)).value = 'Less: Accumulated Depreciation'
    sht.range('E' + str(next_row + 3)).value = '=-depr_amt(E2, Depreciation)'
    sht.range('F' + str(next_row + 3) + ':P' + str(next_row + 3)).value = '= E' + str(next_row + 3) + ' - depr_amt(F2, Depreciation)'
    sht.range('E' + str(next_row + 3) + ':P' + str(next_row + 3)).name = 'AccumDepreciation'
    sht.range('A' + str(next_row + 5)).value = 'Total Assets'
    sht.range('E' + str(next_row + 5) + ':P' + str(next_row + 5)).value = '=E11+E' +  str(next_row - int(tang_assets) - 4)  + '+E' + str(next_row - int(tang_assets) - 2)  + '+E' + str(next_row + 1) + '+E' + str(next_row + 3)
    sht.range('A' + str(next_row + 8)).value = "Liabilities and Owner's Equity"
    sht.range('B' + str(next_row + 9)).value = 'Liabilities'
    sht.range('C' + str(next_row + 10)).value = 'Accounts Payable'
    sht.range('E' + str(next_row + 10) + ':P' + str(next_row + 10)).value = "=OperatingAssumptions!$N$11/30 * acct_payable(E2, FixedExpPL, 'P&L'!E6)"
    sht.range('E' + str(next_row + 10) + ':P' + str(next_row + 10)).name = 'AccountsPayableRow'
    sht.range('C' + str(next_row + 11)).value = 'Loan Payable'
    sht.range('E' + str(next_row + 11) + ':P' + str(next_row + 11)).value = '=loan_payable(E2, LoanBalance)'
    sht.range('E' + str(next_row + 11) + ':P' + str(next_row + 11)).name = 'LoanPayable'
    sht.range('C' + str(next_row + 12)).value = 'Mortgage Payable'
    sht.range('C' + str(next_row + 13)).value = 'Credit Card Debt'
    sht.range('C' + str(next_row + 14)).value = 'Line of Credit Balance'
    sht.range('B' + str(next_row + 15)).value = 'Total Liabilities'
    sht.range('E' + str(next_row + 15) + ':P' + str(next_row + 15)).value = "=SUM(E" + str(next_row + 9) + ":E" + str(next_row + 14)  + ")"
    sht.range('B' + str(next_row + 17)).value = "Owner's Equity"
    sht.range('C' + str(next_row + 18)).value = "Capital Stock"
    sht.range('E' + str(next_row + 18) + ':P' + str(next_row + 18)).value = '=capital_stock(E2,' + str(seed_rds) + ', Capitalization)'
    sht.range('E' + str(next_row + 18) + ':P' + str(next_row + 18)).name = 'CapitalStock'
    sht.range('C' + str(next_row + 19)).value = "Retained Earnings"
    sht.range('E' + str(next_row + 19) + ':P' + str(next_row + 19)).value = '=retained_earnings(E2, EBITDA, Interest, D' + str(next_row + 19) + ')'
    sht.range('B' + str(next_row + 20)).value = "Total Owner's Equity"
    sht.range('E' + str(next_row + 20) + ':P' + str(next_row + 20)).value = "=SUM(E" + str(next_row + 17) + ":E" + str(next_row + 19)  + ")"
    sht.range('A' + str(next_row + 22)).value = "Total Liabilities and Owner's Equity"
    sht.range('E' + str(next_row + 22) + ':P' + str(next_row + 22)).value = "=E" + str(next_row + 15) + "+E" + str(next_row + 20)


def create_cf(wb, sht):
    sht.range('A4').value = 'Beginning Cash Balance'
    sht.range('E4').value = 0
    sht.range('F4:P4').value = '=E36'
    sht.range('A6').value = 'Cash Inflows'
    sht.range('B7').value = 'Income from Sales'
    sht.range('E7:P7').value = "='P&L'!E4"
    sht.range('B8').value = 'Change in A/R'
    sht.range('E8').value = '=0-BS!E6'
    sht.range('F8:P8').value = '=BS!E6-BS!F6'
    sht.range('A9').value = 'Total Cash Inflows'
    sht.range('E9:P9').value = '=SUM(E7:E8)'
    sht.range('A11').value = 'Cash Outflows'
    sht.range('B12').value = 'New Fixed Asset Purchases'
    sht.range('E12:P12').value = '=find_change_in_bs(E2, FixedAssetRow)'
    sht.range('B13').value = 'Change in Raw Materials'
    sht.range('B14').value = 'Change in Inventory'
    sht.range('E14').value = '=0-BS!E8'
    sht.range('F14:P14').value = '=BS!E8-BS!F8'
    sht.range('B15').value = 'Change in A/P'
    sht.range('E15:P15').value = '=-find_change_in_bs(E2, AccountsPayableRow)'
    sht.range('B16').value = 'Cost of Sales'
    sht.range('E16:P16').value = "=-'P&L'!E6"
    sht.range('B17').value = 'Intangible Assets'
    sht.range('E17:P17').value = '=find_change_in_bs(E2, IntangibleAssets)'
    sht.range('B18').value = 'Total Salary and Related'
    sht.range('E18:P18').value = "=-'P&L'!E13"
    sht.range('B19').value = 'Fixed Business Expenses'
    sht.range('E19:P19').value = '=fixed_expenses(E2, FixedExpPL)'
    sht.range('B20').value = 'Preferred Return'
    sht.range('B21').value = 'Loan Interest'
    sht.range('E21:P21').value = '=-int_amt(E2, Interest)'
    sht.range('B22').value = 'Taxes'
    sht.range('A23').value = 'Total Cash Outflows'
    sht.range('E23:P23').value = '=SUM(E11:E22)'
    sht.range('B25').value = 'Add: Amortization'
    sht.range('E25:P25').value = '=find_change_in_bs(E2, AccumAmortization)'
    sht.range('B26').value = 'Add: Depreciation'
    sht.range('E26:P26').value = '=find_change_in_bs(E2, AccumDepreciation)'
    sht.range('A28').value = 'Net Cash Flow'
    sht.range('E28:P28').value = '=E9 + E23 + E25 + E26'
    sht.range('B30').value = 'Equity Financing'
    sht.range('E30:P30').value = '=-find_change_in_bs(E2, CapitalStock)'
    sht.range('B31').value = 'Loan'
    sht.range('E31:P31').value = '=-find_change_in_bs(E2, LoanPayable)'
    sht.range('B32').value = 'Interest Income'
    sht.range('B33').value = 'Short Term Borrowing'
    sht.range('B34').value = 'Short Term Repayments'
    sht.range('A36').value = 'Ending Cash Balance'
    sht.range('E36:P36').value = '=E4 + SUM(E28:E34)'
    sht.range('A38').value = 'Line of Credit Balance'


def main():
    # workbook and sheets
    wb = xw.Book.caller()
    shtb = wb.sheets['BasicAssumptions']
    shtop = wb.sheets['OperatingAssumptions']
    num_of_products = shtb.range('B13').value
    product = xw.Range('Product1')
    sales = int(shtb.range('B16').value)
    rnd = int(shtb.range('B17').value)
    overhead = int(shtb.range('B18').value)
    tang_assets = shtb.range('B21').value
    intang_assets = shtb.range('B22').value
    loan = shtop.range('Loan').value
    seed_rds = shtb.range('B25').value

    last_row_fixed_exp = h.lastRow('OperatingAssumptions', wb, 10)
    shtop.range('J4:K' + str(last_row_fixed_exp)).name = 'FixedExpenses'
    f_expenses = shtop.range('FixedExpenses').value

    # Dataframes
    # income, products, staff, other, fin, asset, cap = makeDataFrames(wb, shtop)

    # Inventory
    h.create_new_sheet(wb, 'Inventory', 'OperatingAssumptions')
    shtinv = wb.sheets['Inventory']
    shtinv.range('A1').value = 'Inventory'
    shtinv.range('E1').value = 'Month'
    shtinv.range('E2').value = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
    create_inv_for_products(wb, shtinv, num_of_products, product)

    # Sales Detail by Product
    h.create_new_sheet(wb, 'SalesDetail', 'Inventory')
    shtsales = wb.sheets['SalesDetail']
    shtsales.range('A1').value = 'Sales Detail'
    shtsales.range('E1').value = 'Month'
    shtsales.range('E2').value = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
    create_sales_for_products(wb, shtsales, num_of_products, product)

    # Cost Detail by Product
    h.create_new_sheet(wb, 'CostDetail', 'SalesDetail')
    shtcos = wb.sheets['CostDetail']
    shtcos.range('A1').value = 'Cost Detail'
    shtcos.range('E1').value = 'Month'
    shtcos.range('E2').value = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
    create_cost_for_products(wb, shtcos, num_of_products, product)

    # Employees
    h.create_new_sheet(wb, 'Employees', 'CostDetail')
    shtemp = wb.sheets['Employees']
    shtemp.range('A1').value = 'Employees'
    shtemp.range('E1').value = 'Month'
    shtemp.range('E2').value = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
    create_employees(wb, shtemp, sales, rnd, overhead)

    # Asset Schedule
    h.create_new_sheet(wb, 'AssetSchedule', 'Employees')
    shtass = wb.sheets['AssetSchedule']
    shtass.range('A1').value = 'Asset Schedule'
    shtass.range('E1').value = 'Month'
    shtass.range('E2').value = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
    create_assets(wb, shtass, tang_assets, intang_assets)

    # Loan Schedule
    h.create_new_sheet(wb, 'LoanSchedule', 'AssetSchedule')
    shtloan = wb.sheets['LoanSchedule']
    shtloan.range('A1').value = 'Loan Schedule'
    shtloan.range('A4').value = ['Period', 'PMT', 'Interest Paid', 'Principal', 'Balance', 'Annual Rate', 'Additional Payment']
    create_loan(wb, shtloan, loan, seed_rds)

    # Profit and Loss
    h.create_new_sheet(wb, "P&L", "LoanSchedule")
    shtpl = wb.sheets['P&L']
    shtpl.range('A1').value = 'Projected Income Statement'
    shtpl.range('E1').value = 'Month'
    shtpl.range('E2').value = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
    create_pl(wb, shtpl, f_expenses)

    # Balance sheet
    h.create_new_sheet(wb, "BS", "P&L")
    shtbs = wb.sheets['BS']
    shtbs.range('A1').value = 'Balance Sheet'
    shtbs.range('E1').value = 'Month'
    shtbs.range('E2').value = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
    create_bs(wb, shtbs, intang_assets, tang_assets, seed_rds)

    # Cash Flow
    h.create_new_sheet(wb, "CF", "BS")
    shtcf = wb.sheets['CF']
    shtcf.range('A1').value = 'Cash Flow Statement'
    shtcf.range('E1').value = 'Month'
    shtcf.range('E2').value = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
    create_cf(wb, shtcf)
    shtbs.range('E5:P5').value = '=CF!E36'
