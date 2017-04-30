import smartsheet, json, requests
from settings import ACCESS_TOKENS

ss = smartsheet.Smartsheet(ACCESS_TOKEN['matt'])

canceled_info = {
    'source_id': 5808854987499396,
    'dest_id': 3045607515416452,
    'column_id': 3951545940240260,
    'status': 'Canceled'
}

def gather_sheet(source_sheet):
'''Input sheet id and return json version.'''
    sheet = ss.Sheets.get_sheet(source_sheet)
    count = sheet.total_row_count
    sheet = ss.Sheets.get_sheet(source_sheet, page_size=count)
    json_sheet = json.loads(str(sheet))
    return(json_sheet)

def gather_row_ids(json_sheet, column_id, value):
'''Iterate through json rows and save to list based on value parameter'''
    row_ids = []
    for row in json_sheet['rows']:
        for cell in row['cells']:
            if cell['columnId'] == column_id:
                if cell['displayValue'] == value:
                    row_ids.append(row['id'])
                    print(row['rowNumber'])
                    return row_ids

def move_rows(row_ids, source, dest):
'''Move row operation implemented using the smartsheet-python-sdk'''
    rowobjs = ss.models.CopyOrMoveRowDirective({
        'row_ids': row_ids,
        'to': ss.models.CopyOrMoveRowDestination({
            'sheet_id': dest
        })
    })
    req = ss.Sheets.move_rows(source, rowobjs, include=['attachments','discussions'], ignore_rows_not_found=False)
    print(req)


def move_rows_2(row_ids, source, dest):
'''Alternative move row operation using the requests module'''
    url = 'https://api.smartsheet.com/2.0/sheets/{}/rows/move'.format(source)
    headers = {
        'authorization': 'Bearer {}'.format(ACCESS_TOKEN),
        'content-type': 'application/json',
        'cache-control': 'no-cache'
    }
    mydata = {
        'rowIds': row_ids,
        'to': {
            'sheetId': dest
        }
    }
    data_json = json.dumps(mydata)
    payload = {'include': ['attachments', 'discussions']}
    req = requests.request('POST', url, data=data_json, headers=headers, params=payload)
    print(req.text)


def archive_canceled_rows():
'''Archive process'''
    json_sheet = gather_sheet(canceled_info['source_id'])
    row_ids = gather_row_ids(json_sheet, canceled_info['column_id'], canceled_info['status'])
    move_rows_2(row_ids, canceled_info['source_id'], canceled_info['dest_id'])

def main():
    archive_canceled_rows()


if __name__ == '__main__':
    main()


#TODO continue testing, implement batch processign of rows
