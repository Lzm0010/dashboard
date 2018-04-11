import xlwings as xw
import helpers as h


def create_channel(idx, wb, channel, flex='Base Number', flex2='Sales Rate', flex3='Base Units'):
    next_row = h.lastRow(idx, wb) + 1
    sht = wb.sheets[idx]
    sht.range('A' + str(next_row)).value = [
        [channel], [flex], ['Initial Month'],
        ['Growth Frequency'], ['Growth Amount'], [flex2],
        [flex3], ['Initial Month of Sales'], ['Growth Frequency'],
        ['Growth Amount']
    ]
    sht.range('B' + str(next_row + 1) + ":B" + str(next_row + 4)).color = (255, 242, 204)
    sht.range('B' + str(next_row + 6) + ":B" + str(next_row + 9)).color = (255, 242, 204)
    sht.range('B' + str(next_row + 1) + ":B" + str(next_row + 4)).value = [[1], [1], [1], [1]]
    sht.range('B' + str(next_row + 6) + ":B" + str(next_row + 9)).value = [[10], [1], [1], [10]]
    sht.range('A' + str(next_row) + ":B" + str(next_row + 9)).name = channel


def create_products(idx, wb, prod_num, retail, ws, event, ecom):
    sht = wb.sheets[idx]
    for i in range(int(prod_num)):
        next_row = h.lastRow(idx, wb, 4) + 1
        product_info = [
            ['Type'], ['Product Name'], ['Price and Revenue Share'],
            ['Price'], ['Annual Price Change'], ['Third Party Revenue Share'],
            ['Customer Acquisition Cost'], ['Sales Commission'],
            ['Sales per Month per Salesperson'], ['Sales Base Salary'],
            ['Cost of Good'], ['Cost'], ['Annual Cost Change'],
            ['Direct Labor Hours'], ['Direct Labor Wage per Hour'],
            ['Shipping Expense'], ['Payment Processing Percentage'],
            ['Acct Mgmt & Support by Revenue'],
            ['Acct Mgmt Employee Base Salary'],
            ['Inventory'], ['Beginning Inventory'], ['Reorder Point'],
            ['Growth Frequency'], ['Growth Amount'], ['Channels']
        ]
        num_of_defaults = []
        if retail == "yes":
            product_info += [['Stores']]
            num_of_defaults += [['yes']]
        if ws == "yes":
            product_info += [['Wholesale']]
            num_of_defaults += [['yes']]
        if event == "yes":
            product_info += [['Events']]
            num_of_defaults += [['yes']]
        if ecom == "yes":
            product_info += [['Website']]
            num_of_defaults += [['yes']]
        sht.range('D' + str(next_row)).value = product_info
        sht.range('E' + str(next_row + 1)).color = (255, 242, 204)
        sht.range('E' + str(next_row + 3) + ":E" + str(next_row + 5)).color = (255, 242, 204)
        sht.range('E' + str(next_row + 7) + ":E" + str(next_row + 9)).color = (255, 242, 204)
        sht.range('E' + str(next_row + 11) + ":E" + str(next_row + 18)).color = (255, 242, 204)
        sht.range('E' + str(next_row + 20) + ":E" + str(next_row + 23)).color = (255, 242, 204)
        sht.range('E' + str(next_row + 25) + ":E" + str(next_row + (len(product_info) - 1))).color = (255, 242, 204)
        sht.range('E' + str(next_row + 1)).value = "Product " + str(i+1)
        sht.range('E' + str(next_row + 3) + ":E" + str(next_row + 5)).value = [[20], [0], [0]]
        sht.range('E' + str(next_row + 7) + ":E" + str(next_row + 9)).value = [[0], [0], [0]]
        sht.range('E' + str(next_row + 11) + ":E" + str(next_row + 18)).value = [[5000], [0], [0], [0], [1250], [0.02], [100000], [40000]]
        sht.range('E' + str(next_row + 20) + ":E" + str(next_row + 23)).value = [[600], [100], [1], [0]]
        sht.range('E' + str(next_row + 25) + ":E" + str(next_row + (len(product_info) - 1))).value = num_of_defaults
        sht.range('D' + str(next_row + 1) + ":E" + str(next_row + (len(product_info) - 1))).name = "Product" + str(i + 1)


def create_job(idx, wb, num_of_jobs, job_type):
    sht = wb.sheets[idx]
    next_row = h.lastRow(idx, wb, 7) + 1
    sht.range('G' + str(next_row)).value = job_type
    for i in range(int(num_of_jobs)):
        next_row = h.lastRow(idx, wb, 7) + 1
        sht.range('G' + str(next_row)).value = [
            ['Job Title'], ['Base Salary'], ['Base Number of Employees'],
            ['Initial Month'], ['Growth Frequency'], ['Growth Amount']
        ]
        sht.range('H' + str(next_row) + ":H" + str(next_row + 5)).color = (255, 242, 204)
        sht.range('H' + str(next_row) + ":H" + str(next_row + 5)).value = [[job_type + str(i+1)], [40000], [1], [1], [6], [1]]
        sht.range('G' + str(next_row) + ":H" + str(next_row + 5)).name = job_type + str(i + 1)


def create_bonus(idx, wb):
    sht = wb.sheets[idx]
    next_row = h.lastRow(idx, wb, 7) + 1
    sht.range('G' + str(next_row)).value = [
        ['Bonus and Other Employee Expenses'], ['Recruiting Costs'],
        ['Computer Equipment'], ['Software and Licensing'],
        ['Communication Services'], ['Travel & Entertainment'],
        ['Outside Services'], ['Subscription and Dues'], ['Annual Bonus'],
        ['Benefits'], ['Tax Percentage'], ['Training Costs'], ['Annual Raise'],
        ['Employee Turnover Decimal']
    ]
    sht.range('H' + str(next_row + 1) + ":H" + str(next_row + 13)).color = (255, 242, 204)
    sht.range('H' + str(next_row + 1) + ":H" + str(next_row + 13)).value = [[0], [0], [0], [0], [0], [0], [0], [0], [263], [5.6], [0], [0], [0]]
    sht.range('G' + str(next_row + 1) + ":H" + str(next_row + 13)).name = 'EmployeeBonus'


def create_other(idx, wb):
    sht = wb.sheets[idx]
    next_row = h.lastRow(idx, wb, 10) + 2
    sht.range('J' + str(next_row)).value = [
        ['Additional Ad & Marketing'], ['Bank Fees'],
        ['Insurance'], ['License Fees & Permits'],
        ['Legal & Professional Fees'], ['Office Expenses & Supplies'],
        ['Rent'], ['Miscellaneous'], ['Utilities'], ['Website']
    ]
    sht.range('K' + str(next_row) + ":K" + str(next_row + 9)).color = (255, 242, 204)
    sht.range('K' + str(next_row) + ":K" + str(next_row + 9)).value = [[3000], [300], [274], [0], [500], [100], [3000], [0], [300], [100]]


def create_financials(idx, wb):
    sht = wb.sheets[idx]
    next_row = h.lastRow(idx, wb, 13) + 1
    sht.range('M' + str(next_row)).value = [
        ['Income Statement'], ['Interest Rate'],
        ['Discount Rate'], ['Bad Debt Percentage'],
        ['Existing Loss Carry Forward'], ['Balance Sheet'],
        ['Opening Cash Balance'], ['Acct Receivable Days'],
        ['Acct Payable Days'], ['Days of Inventory on Hand']
    ]
    sht.range('N' + str(next_row + 1) + ":N" + str(next_row + 4)).color = (255, 242, 204)
    sht.range('N' + str(next_row + 1) + ":N" + str(next_row + 4)).value = [[4.25], [0.02], [0.01], [0]]
    sht.range('N' + str(next_row + 6) + ":N" + str(next_row + 9)).color = (255, 242, 204)
    sht.range('N' + str(next_row + 6) + ":N" + str(next_row + 9)).value = [[0], [30], [60], [60]]


def create_assets(idx, wb, num_of_assets, asset_type):
    sht = wb.sheets[idx]
    next_row = h.lastRow(idx, wb, 16) + 1
    sht.range('P' + str(next_row)).value = asset_type
    for i in range(int(num_of_assets)):
        next_row = h.lastRow(idx, wb, 16) + 1
        sht.range('P' + str(next_row)).value = [
            ['Name'], ['Month Acquired'], ['CapEx'], ['Useful Life']
        ]
        sht.range('Q' + str(next_row) + ":Q" + str(next_row + 3)).color = (255, 242, 204)
        sht.range('Q' + str(next_row) + ":Q" + str(next_row + 3)).value = [['Asset ' + str(i + 1)], [1], [10000], [5]]


def create_seed(idx, wb, num_of_rds):
    sht = wb.sheets[idx]
    next_row = h.lastRow(idx, wb, 19) + 1
    sht.range('S' + str(next_row)).value = 'Equity'
    for i in range(int(num_of_rds)):
        next_row = h.lastRow(idx, wb, 19) + 1
        sht.range('S' + str(next_row)).value = [
            ['Seed ' + str(i + 1)], ['Month'], ['Cash Amount'],
            ['Shares Amount'],
        ]
        sht.range('T' + str(next_row + 1) + ':T' + str(next_row + 3)).color = (255, 242, 204)
        sht.range('T' + str(next_row + 1) + ':T' + str(next_row + 3)).value = [[1 + i], [160000], [1000000]]


def create_cap(idx, wb):
    sht = wb.sheets[idx]
    next_row = h.lastRow(idx, wb, 19) + 1
    sht.range('S' + str(next_row)).value = [
        ['Debt'], ['Loan Amount'], ['Interest Rate'],
        ['Down Payment Percentage'], ['Length of Loan'], ['Exit Assumptions'],
        ['Founder Shares'], ['Convertible Note'], ['Liquidation Preference'],
        ['Exit Valuation Method'], ['Exit Valuation Multiple']
    ]
    sht.range('T' + str(next_row + 1) + ':T' + str(next_row + 4)).color = (255, 242, 204)
    sht.range('T' + str(next_row + 1) + ':T' + str(next_row + 4)).value = [[500000], [0.0525], [0.20], [20]]
    sht.range('T' + str(next_row + 6) + ':T' + str(next_row + 10)).color = (255, 242, 204)
    sht.range('T' + str(next_row + 6) + ':T' + str(next_row + 10)).value = [[1000000], [0], [False], ['EBITDA'], [9.73]]
    sht.range('S' + str(next_row + 1) + ':T' + str(next_row + 4)).name = 'Loan'


def create_ops(wb):

    # sheets
    shtb = wb.sheets['BasicAssumptions']
    shtop = wb.sheets['OperatingAssumptions']

    # values
    retail = str(shtb.range('B7').value).lower()  # yes or no
    wholesale = shtb.range('B8').value.lower()  # yes or no
    event = shtb.range('B9').value.lower()  # yes or no
    ecommerce = shtb.range('B10').value.lower()  # yes or no
    num_of_products = shtb.range('B13').value  # int
    sales_marketing_roles = shtb.range('B16').value  # int
    r_and_d_roles = shtb.range('B17').value  # int
    overhead_roles = shtb.range('B18').value  # int
    tang_assets = shtb.range('B21').value  # int
    intang_assets = shtb.range('B22').value  # int
    seed_rounds = shtb.range('B25').value  # int

    # create sheet
    shtop.range('A1').value = [['Operating Assumptions'], ['Income']]
    shtop.range('D2').value = "Products"
    shtop.range('G2').value = "Staff"
    shtop.range('J2').value = "Other Expenses"
    shtop.range('M2').value = "Financials"
    shtop.range('P2').value = "Asset Investment"
    shtop.range('S2').value = "Capitalization"

    if retail == 'yes':
        create_channel('OperatingAssumptions', wb, 'Stores')

    if wholesale == 'yes':
        create_channel('OperatingAssumptions', wb, 'Wholesale')

    if event == 'yes':
        create_channel('OperatingAssumptions', wb, 'Events')

    if ecommerce == 'yes':
        create_channel('OperatingAssumptions', wb, 'Website', 'Base Visitors', 'Conversion Rate', 'Base %')

    chan_end_row = h.lastRow('OperatingAssumptions', wb, 1)
    shtop.range('A3:B' + str(chan_end_row)).name = "All_Channels"

    if num_of_products is not None:
        create_products('OperatingAssumptions', wb, num_of_products, retail, wholesale, event, ecommerce)
        prod_end_row = h.lastRow('OperatingAssumptions', wb, 4)
        shtop.range('D3:E' + str(prod_end_row)).name = "All_Products"

    if sales_marketing_roles is not None:
        create_job('OperatingAssumptions', wb, sales_marketing_roles, "General_Sales_and_Marketing")

    if r_and_d_roles is not None:
        create_job('OperatingAssumptions', wb, r_and_d_roles, "Research_and_Development")

    if overhead_roles is not None:
        create_job('OperatingAssumptions', wb, overhead_roles, "Overhead")

    create_bonus('OperatingAssumptions', wb)
    create_other('OperatingAssumptions', wb)
    create_financials('OperatingAssumptions', wb)

    if tang_assets is not None:
        create_assets('OperatingAssumptions', wb, tang_assets, "Furniture, Fixtures, & Equipment")

    if intang_assets is not None:
        create_assets('OperatingAssumptions', wb, intang_assets, "Intangible Assets")

    if seed_rounds is not None:
        create_seed('OperatingAssumptions', wb, seed_rounds)

    create_cap('OperatingAssumptions', wb)
    cap_end_row = h.lastRow('OperatingAssumptions', wb, 20)
    shtop.range('S4:T' + str(cap_end_row)).name = 'Capitalization'

    shtop.autofit('c')

    wb.macro('CreateButton')()
