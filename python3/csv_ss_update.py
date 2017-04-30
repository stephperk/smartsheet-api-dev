import csv, smartsheet, json, requests
from datetime import datetime
from settings import ACCESS_TOKENS

access_token = ACCESS_TOKENS['dataranglr']
ss = smartsheet.Smartsheet(access_token)

def find_sheet_id():
    print('Hello! Which Smartsheet would you like to update?')
    name = raw_input()
    print('Fetching {}...'.format(name))
    sheets = ss.Sheets.list_sheets()
    for sheet in sheets.data:
        if sheet.name == name:
            print('Sheet found!')
            return sheet.id
    print('Sheet not found. Make sure spellings match exactly.')

def find_col_ids(sheet_id):
    col_dict = {
        'SKU': None,
        'QOH': None,
        'Priority': None,
        'Date': None
    }
    print('Acquiring columns...')
    columns = ss.Sheets.get_columns(sheet_id, include_all=True)
    for col in columns.data:
        if col.title == 'SKU':
            col_dict['SKU'] = col.id
        elif col.title == 'QOH':
            col_dict['QOH'] = col.id
        elif col.title == 'Priority':
            col_dict['Priority'] = col.id
        elif col.title == 'Priority_QOH_Updated':
            col_dict['Date'] = col.id
    return col_dict

def find_sheet(sheet_id, col_dict):
    sheet = ss.Sheets.get_sheet(sheet_id)
    count = sheet.total_row_count
    sheet = ss.Sheets.get_sheet(sheet_id, page_size=count, column_ids=['{},{},{}'.format(col_dict['SKU'], col_dict['QOH'], col_dict['Priority'])])
    return sheet

def get_csv_data():
    csv_data = {}
    print('What is the exact name of the csv file? Make sure to include file extension (.csv)!')
    csv_name = raw_input()
    with open(csv_name, 'rb') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            csv_data[row[0]] = {
                'Priority': row[1],
                'QOH': row[2]
            }
    print('csv data acquired...')
    return csv_data


def get_ss_data(sheet, col_dict):
    js_sheet = json.loads(str(sheet))
    ss_data = {}
    for row in js_sheet['rows']:
        ss_data[row['id']] = None
        for cell in row['cells']:
            if cell['columnId'] == int(col_dict['SKU']):
                ss_data[row['id']] = cell['displayValue']
    print('Smartsheet data acquired...')
    return ss_data


def clean_ss_data(ss_data):
    cleaned_ss_data = {k: v for k, v in ss_data.items() if v}
    return cleaned_ss_data


def update_rows(sheet_id, row_id, qoh_value, priority_value, col_dict):
    url = 'https://api.smartsheet.com/2.0/sheets/{}/rows/{}'.format(sheet_id, row_id)
    data = {
        'cells': [
                    {
                    'columnId': col_dict['QOH'],
                    'value': qoh_value
                },
                    {
                    'columnId': col_dict['Priority'],
                    'value': priority_value
                },
                    {
                    'columnId': col_dict['Date'],
                    'value': str(datetime.date(datetime.now()))
                }
            ]
        }
    json_data = json.dumps(data)
    headers = {
        'authorization': 'Bearer {}'.format(access_token),
        'content-type': 'application/json',
        'cache-control': 'no-cache'
    }
    req = requests.request('PUT', url, data=json_data, headers=headers)
    try:
        js_resp = json.loads(req.text)
        for i in js_resp['result']:
            row_id = i['id']
            row_num = i['rowNumber']

        for result in js_resp['result']:
            for cell in result['cells']:
                if cell['columnId'] == int(col_dict['SKU']):
                    sku = cell['displayValue']
                elif cell['columnId'] == int(col_dict['QOH']):
                    qoh = cell['displayValue']
                elif cell['columnId'] == int(col_dict['Priority']):
                    priority = cell['displayValue']

        print('{}, row {}'.format(js_resp['message'], row_num))
        print(row_id, sku, qoh, priority)
        print
    except:
        print('Exception hit')


def find_matches(sheet_id, csv_data, cleaned_ss_data, col_dict):
    sku_not_found = []
    for x, y in cleaned_ss_data.items():
        try:
            qoh = csv_data[y]['QOH']
            priority = csv_data[y]['Priority']
            print(x, y, qoh, priority)
            update_rows(sheet_id, x, qoh, priority, col_dict)
        except KeyError:
            sku_not_found.append(y)
    return sku_not_found


def main():
    sheet_id = find_sheet_id()
    col_dict = find_col_ids(sheet_id)
    sheet = find_sheet(sheet_id, col_dict)
    csv_data = get_csv_data()
    ss_data = get_ss_data(sheet, col_dict)
    cleaned_ss_data = clean_ss_data(ss_data)
    find_matches(sheet_id, csv_data, cleaned_ss_data, col_dict)

if __name__ == '__main__':
    main()
