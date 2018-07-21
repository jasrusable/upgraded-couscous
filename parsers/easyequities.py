from datetime import datetime
from openpyxl import load_workbook
from pprint import PrettyPrinter


pp = PrettyPrinter().pprint


def is_recurring_buying_instruction_row(row):
    if 'RELEASE Reserved funds for Buying Instruction:' in row['comment']:
        return True
    return False


def get_buying_instruction_id(row):
    return int(row['comment'].split(': ')[-1])


def is_empty(row):
    return row.get('date') is None


def get_units_delta(row):
    parts = row['comment'].split(' ')
    assert parts[-2] == '@'
    units = parts[-3]
    return float(units)


def get_price(row):
    parts = row['comment'].split(' ')
    assert parts[-2] == '@'
    price = parts[-1]
    return float(price.replace(',', ''))


def get_action(row):
    parts = row['comment'].split(' ')
    assert parts[-2] == '@'
    action_map = {
        'Bought': 'buy',
        'Sold': 'sell',
    }
    return action_map[parts[0]]


def is_buy_or_sell_row(row):
    parts = row['comment'].split(' ')
    has_action_as_first_word = parts[0] in ['Bought', 'Sold']
    has_at_symbol_seperating_units_and_price = len(parts) > 2 and parts[-2] == '@'
    return has_action_as_first_word and has_at_symbol_seperating_units_and_price


def parse_row(excel_sheet_row):
    return {
        'date': excel_sheet_row[0].value,
        'comment': excel_sheet_row[1].value.replace('\xa0', ' '),
        'debit_credit': float(excel_sheet_row[2].value),
    }


def get_fund_name_from_buy_or_sell_row(row):
    assert is_buy_or_sell_row(row)
    parts = row['comment'].split(' ')
    return ' '.join(parts[1: -3])


def is_recurring_investment_fee_row(row):
    if 'Recurring Investment Fee' in row['comment']:
        return True
    return False


def is_investor_protection_levy_row(row):
    if 'Investor protection levy (IPL) and administration' in row['comment']:
        return True
    return False

def is_broker_commission_row(row):
    if 'Broker Commission' in row['comment']:
        return True
    return False


def is_settlement_and_administration_row(row):
    if 'Settlement and administration' in row['comment']:
        return True
    return False


def is_vat_on_costs_row(row):
    if 'Value Added Tax on costs (VAT)' in row['comment']:
        return True
    return False


def get_broker_commission_rate(row):
    assert is_broker_commission_row(row)
    parts = row['comment'].split(' ')
    return float(parts[-1])


def get_fund_name_from_broker_commission_row(row):
    assert is_broker_commission_row(row)
    parts = row['comment'].split(' ')
    return ' '.join(parts[2: -2])


def parse_buy_or_sell_row(buy_or_sell_row):
    assert is_buy_or_sell_row(buy_or_sell_row)
    units_delta = get_units_delta(buy_or_sell_row)
    price = get_price(buy_or_sell_row)
    fund_name = get_fund_name_from_buy_or_sell_row(buy_or_sell_row)
    return {
        'action': get_action(buy_or_sell_row),
        'units_delta': units_delta,
        'price': price,
        'fund_name': fund_name,
        'net_amount': buy_or_sell_row['debit_credit'],
    }


def parse_recurring_investment_fee_row(recurring_investment_fee_row):
    assert is_recurring_investment_fee_row(recurring_investment_fee_row)
    return {
        'name': 'recurring_investment_fee',
        'type': 'fixed',
        'amount': recurring_investment_fee_row['debit_credit'],
    }


def parse_broker_commission_fee_row(broker_commission_row):
    assert is_broker_commission_row(broker_commission_row)
    broker_commission_rate = get_broker_commission_rate(broker_commission_row)
    return {
        'name': 'broker_commission',
        'type': 'percentage',
        'rate': broker_commission_rate,
        'amount': broker_commission_row['debit_credit'],
    }

def parse_settlement_and_administration_fee_row(settlement_and_administration_row):
    assert is_settlement_and_administration_row(settlement_and_administration_row)
    return {
        'name': 'settlement_and_administration',
        'type': 'fixed',
        'amount': settlement_and_administration_row['debit_credit'],
    }


def parse_investor_protection_levy_fee_row(investor_protection_levy_row):
    assert is_investor_protection_levy_row(investor_protection_levy_row)
    return {
        'name': 'investor_protection_levy_and_administration',
        'type': 'fixed',
        'amount': investor_protection_levy_row['debit_credit'],
    }


def parse_vat_on_costs_fee_row(vat_on_costs_row):
    assert is_vat_on_costs_row(vat_on_costs_row)
    return {
        'name': 'vat_on_costs',
        'type': 'fixed',
        'amount': vat_on_costs_row['debit_credit'],
    }


def parse_release_funds_row(release_funds_row):
    return {
        'action_type': 'recurring',
    }


def parse_recurring_investment_purchase_post_2018(release_funds_row, buy_or_sell_row, recurring_investment_fee_row, broker_commission_row, settlement_and_administration_row, investor_protection_levy_row, vat_on_costs_row):
    data = {
        **parse_release_funds_row(release_funds_row),
        **parse_purchase_or_sale(
            buy_or_sell_row,
            broker_commission_row,
            settlement_and_administration_row,
            investor_protection_levy_row,
            vat_on_costs_row,
        )
    }
    data['fees'].append(parse_recurring_investment_fee_row(recurring_investment_fee_row))
    return data


def parse_recurring_investment_purchase_pre_2018(release_funds_row, buy_or_sell_row, broker_commission_row, settlement_and_administration_row, investor_protection_levy_row, vat_on_costs_row):
    return {
        **parse_release_funds_row(release_funds_row),
        **parse_purchase_or_sale(
            buy_or_sell_row,
            broker_commission_row,
            settlement_and_administration_row,
            investor_protection_levy_row,
            vat_on_costs_row,
        )
    }


def parse_purchase_or_sale(buy_or_sell_row, broker_commission_row, settlement_and_administration_row, investor_protection_levy_row, vat_on_costs_row):
    return {
        'date': buy_or_sell_row['date'],
        **parse_buy_or_sell_row(buy_or_sell_row),
        'fees': [
            parse_broker_commission_fee_row(broker_commission_row),
            parse_settlement_and_administration_fee_row(settlement_and_administration_row),
            parse_investor_protection_levy_fee_row(investor_protection_levy_row),
            parse_vat_on_costs_fee_row(vat_on_costs_row)
        ]
    }


def do_rows_have_same_date(rows):
    return all(x['date'] == rows[0]['date'] for x in rows)


def parse_easy_equities_transaction_history(file):
    workbook = load_workbook(file, read_only=True)
    sheet = workbook['Transaction History']
    rows = sheet.rows
    header_row = rows.__next__()
    assert header_row[0].value == 'Date'
    assert header_row[1].value == 'Comment'
    assert header_row[2].value == 'Debit/Credit'
    none_empty_rows = filter(lambda x: x[0].value is not None, rows)
    parsed_rows = (parse_row(r) for r in none_empty_rows)
    parsed_transactions = []
    for row in parsed_rows:
        if is_recurring_buying_instruction_row(row):
            date = row['date']
            transaction_rows = {
                'release_funds_row': row,
                'buy_or_sell_row': parsed_rows.__next__(),
            }
            IS_ROW_POST_2018 = date >= datetime(2018, 1, 1)
            if IS_ROW_POST_2018:
                transaction_rows['recurring_investment_fee_row'] = parsed_rows.__next__()
            transaction_rows = {
                **transaction_rows,
                'broker_commission_row': parsed_rows.__next__(),
                'settlement_and_administration_row': parsed_rows.__next__(),
                'investor_protection_levy_row': parsed_rows.__next__(),
                'vat_on_costs_row': parsed_rows.__next__(),
            }
            assert do_rows_have_same_date(list(transaction_rows.values()))
            if IS_ROW_POST_2018:
                f = parse_recurring_investment_purchase_post_2018
            else:
                f = parse_recurring_investment_purchase_pre_2018
            transaction = f(**transaction_rows)
            parsed_transactions.append(transaction)
        if is_buy_or_sell_row(row):
            transaction_rows = {
                'buy_or_sell_row': row,
                'broker_commission_row': parsed_rows.__next__(),
                'settlement_and_administration_row': parsed_rows.__next__(),
                'investor_protection_levy_row': parsed_rows.__next__(),
                'vat_on_costs_row': parsed_rows.__next__(),
            }
            assert do_rows_have_same_date(list(transaction_rows.values()))
            transaction = parse_purchase_or_sale(**transaction_rows)
            parsed_transactions.append(transaction)
    return parsed_transactions


EE_TRANSACTION_HISTORY = './file_samples/EasyEquitiesTransactionHistoryExport.xlsx'
f = open(EE_TRANSACTION_HISTORY, 'rb')

transactions = parse_easy_equities_transaction_history(f)
pp(transactions)
