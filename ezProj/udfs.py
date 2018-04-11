import xlwings as xw
import math


# inventory functions
@xw.func
def units_sold_channel_method(month, all_channels, product, all_products):
    # get number of channels
    cells = 0
    for row in all_channels:
        cells += 1

    num_of_channels = cells / 10

    # base number of units
    total_units = 0

    # loop thru channels
    for i in range(1, int(num_of_channels + 1)):

        # variables
        channel_name = all_channels[0 + ((i - 1) * 10)][0]
        base_outlets = all_channels[1 + ((i - 1) * 10)][1]
        outlet_initial_month = all_channels[2 + ((i - 1) * 10)][1]
        chan_growth_freq = all_channels[3 + ((i - 1) * 10)][1]
        chan_growth_amt = all_channels[4 + ((i - 1) * 10)][1]
        base_units = all_channels[6 + ((i - 1) * 10)][1]
        sales_initial_month = all_channels[7 + ((i - 1) * 10)][1]
        sales_growth_freq = all_channels[8 + ((i - 1) * 10)][1]
        sales_growth_amt = all_channels[9 + ((i - 1) * 10)][1]

        # get number of products for this channel
        num_of_products = 0
        for row in all_products:
            if row[0] == channel_name and row[1] == 'yes':
                num_of_products += 1

        # get row number of channel to check for if it applies to product
        for i, row in enumerate(product):
            if row[0] == channel_name:
                row_number = i

        # if product is distributed by channel...
        if product[row_number][1] == 'yes':
            # find number of stores by Month. For website this is visitors
            if month - (outlet_initial_month - 1) <= 0:
                stores = 0
            else:
                stores = (math.ceil((month - (outlet_initial_month - 1)) / chan_growth_freq) * chan_growth_amt) - chan_growth_amt + base_outlets

            # get number of units by Month
            if month - (sales_initial_month - 1) <= 0:
                units = 0
            else:
                if channel_name == 'Website':
                    # find growth rate then multiply times the amount and add to base number
                    units = (math.ceil((month - (sales_initial_month - 1)) / sales_growth_freq) * (sales_growth_amt * 0.01)) - (sales_growth_amt * 0.01) + (base_units * 0.01)
                else:
                    units = (math.ceil((month - (sales_initial_month - 1)) / sales_growth_freq) * sales_growth_amt) - sales_growth_amt + base_units

            # return stores times units divided by num of products in channel
            total_units += math.floor((stores * units) / num_of_products)
        else:
            total_units += 0

    return total_units


@xw.func
def units_sold_product_method():
    # take formula based on inventory turnover
    # add inventory turnover to products
    pass


@xw.func
def units_added(units, sold, reorder_point, order_amount):
    if (int(units or 0) - int(sold or 0)) <= int(reorder_point or 0):
        return int(order_amount or 0)
    else:
        return 0


@xw.func
def inv_cost_added(month, units_added_row, product):
    #get total product cost
    cost = product[10][1]
    dl_hours = product[12][1]
    dl_wage_per_hour = product[13][1]
    shipping_expense = product[14][1]
    pmt_processing_pctg = product[15][1]
    reorder_amt = product[19][1]
    total_cost = ((dl_hours * dl_wage_per_hour) + cost + shipping_expense)
    cost_per_unit = (total_cost + (total_cost * pmt_processing_pctg)) / reorder_amt
    return round(cost_per_unit * units_added_row[int(month-1)], 2)


@xw.func
def remaining_units(start, sold, added):
    return int(start or 0) - int(sold or 0) + int(added or 0)


@xw.func
def total_inventory_cost_added(month, all_inv, products):
    col = int(month - 1)
    cost = 0
    for i in range(int(products)):
        cost += all_inv[(i * 6) + 3][col]

    return cost


# sales detail functions
@xw.func
def get_ws_price(price, ws_percentage):
    return price * (ws_percentage * 0.01)


@xw.func
def retail_units(total_units, ws_units):
    return total_units - ws_units


@xw.func
def wholesale_units(month, channel, product, all_products):
    # variables
    channel_name = channel[0][0]
    base_outlets = channel[1][1]
    outlet_initial_month = channel[2][1]
    chan_growth_freq = channel[3][1]
    chan_growth_amt = channel[4][1]
    base_units = channel[6][1]
    sales_initial_month = channel[7][1]
    sales_growth_freq = channel[8][1]
    sales_growth_amt = channel[9][1]

    # get number of products for this channel
    num_of_products = 0
    for row in all_products:
        if row[0] == channel_name and row[1] == 'yes':
            num_of_products += 1

    # get row number of channel to check for if it applies to product
    for i, row in enumerate(product):
        if row[0] == channel_name:
            row_number = i

    # if product is distributed by channel...
    if product[row_number][1] == 'yes':
        # find number of stores by Month. For website this is visitors
        if month - (outlet_initial_month - 1) <= 0:
            stores = 0
        else:
            stores = (math.ceil((month - (outlet_initial_month - 1)) / chan_growth_freq) * chan_growth_amt) - chan_growth_amt + base_outlets

        # get number of units by Month
        if month - (sales_initial_month - 1) <= 0:
            units = 0
        else:
            units = (math.ceil((month - (sales_initial_month - 1)) / sales_growth_freq) * sales_growth_amt) - sales_growth_amt + base_units

        # return stores times units divided by num of products in channel
        total_units = math.floor((stores * units) / num_of_products)
    else:
        total_units = 0

    return total_units


@xw.func
def total_sales(r_price, ws_price, r_units, ws_units):
    return (int(r_price or 0) * int(r_units or 0)) + (int(ws_price or 0) * int(ws_units or 0))


@xw.func
def all_products_retail_sales(month, all_sales, products):
    col = int(month - 1)
    sales = 0
    for i in range(int(products)):
        retail_price = all_sales[i * 3][col]
        retail_units = all_sales[(i * 3) + 3][col]
        sales += round(retail_price * retail_units, 2)

    return sales


@xw.func
def all_products_ws_sales(month, all_sales, products):
    col = int(month - 1)
    sales = 0
    for i in range(int(products)):
        retail_price = all_sales[(i * 3) + 1][col]
        retail_units = all_sales[(i * 3) + 4][col]
        sales += round(retail_price * retail_units, 2)

    return sales


# cost detail functions
@xw.func
def monthly_cost_per_product(product, units):
    cost = product[10][1]
    dl_hours = product[12][1]
    dl_wage_per_hour = product[13][1]
    shipping_expense = product[14][1]
    pmt_processing_pctg = product[15][1]
    reorder_amt = product[19][1]
    total_cost = ((dl_hours * dl_wage_per_hour) + cost + shipping_expense)
    cost_per_unit = (total_cost + (total_cost * pmt_processing_pctg)) / reorder_amt
    return round(cost_per_unit * units, 2)


@xw.func
def all_products_cost(month, all_cost, products):
    col = int(month - 1)
    cost = 0
    for i in range(int(products)):
        cost += all_cost[(i * 5) + 3][col]

    return cost


# employee functions
@xw.func
def employee_cost(month, employee):
    salary = employee[1][1]
    base_emp = employee[2][1]
    emp_init_month = employee[3][1]
    emp_growth_freq = employee[4][1]
    emp_growth_amt = employee[5][1]

    if month - (emp_init_month - 1) <= 0:
        emp_num = 0
    else:
        emp_num = (math.ceil((month - (emp_init_month - 1)) / emp_growth_freq) * emp_growth_amt) - emp_growth_amt + base_emp

    return round(emp_num * (salary / 12), 2)


@xw.func
def tax_and_benefits(total, bonusRange):
    benefits = bonusRange[8][1]
    tax = (bonusRange[9][1] * 0.01)
    return round((total * tax) + benefits, 2)


# asset functions
@xw.func
def asset_monthly_cost(month, init_month, capex, life):
    if month - (init_month - 1) <= 0:
        cost = 0
    else:
        cost = round(capex / (life * 12), 2)

    return cost


# profit and loss functions
@xw.func
def total_income(month, product_row):
    return product_row[int(month - 1)]


@xw.func
def total_cost(month, product_row):
    return product_row[int(month - 1)]


@xw.func
def amortization_amt(month, row):
    return row[int(month-1)]


@xw.func
def depr_amt(month, row):
    return row[int(month-1)]

@xw.func
def int_amt(month, col):
    return col[int(month-1)]


# balance sheet functions
@xw.func
def inventory(month, cogs, inv_cost, prev_month):
    if prev_month is not None:
        total_inv = float(prev_month) - cogs + inv_cost[int(month) - 1]
        return total_inv
    else:
        return inv_cost[int(month) - 1] - cogs


@xw.func
def acct_payable(month, fixed_exp, cogs):
    sum = cogs
    for row in fixed_exp:
        sum += row[int(month - 1)]
    return sum


@xw.func
def loan_payable(month, loan_balance):
    return loan_balance[int(month-1)]


@xw.func
def capital_stock(month, number_of_rds, financing_inputs):
    stock = 0
    if number_of_rds > 0:
        for i in range(1, int(number_of_rds + 1)):
            init_month = financing_inputs[(i*4)-3][1]
            cash = financing_inputs[(i*4)-2][1]
            if month - (init_month - 1) <= 0:
                stock += 0
            else:
                stock += cash

    return stock


@xw.func
def retained_earnings(month, ebitda, interest, prev_month):
    if prev_month is not None:
        retained_earnings = float(prev_month) + ebitda[int(month) - 1] - interest[int(month -1)]
        return retained_earnings
    else:
        return ebitda[int(month) - 1] - interest[int(month -1)]


# cash flow functions
@xw.func
def find_change_in_bs(month, row):
    if month - 1 != 0:
        return row[int(month) - 2] - row[int(month) - 1]
    else:
        return 0 - row[int(month) - 1]


@xw.func
def fixed_expenses(month, fixed_exp):
    sum = 0
    for row in fixed_exp:
        sum += row[int(month - 1)]
    return -sum
