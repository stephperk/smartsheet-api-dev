import smartsheet, json, requests, pprint
from settings import ACCESS_TOKENS

access_token = ACCESS_TOKENS['dataranglr']
ss = smartsheet.Smartsheet(access_token)
pp = pprint.PrettyPrinter()


def gimme_sheet(sheet_name):
'''Takes string of desired sheet's name; outputs list of sheet.name and sheet.id'''
    req = ss.Sheets.list_sheets(include_all=True)
    sheets = req.data
    for sheetinfo in sheets:
        if sheetinfo.name == sheet_name:
            sheet = ss.Sheets.get_sheet(sheetinfo.id)
            result = [sheet.name, sheet.id]
            return result


def show_pending_urs(sheet_name):
'''Takes string of desired sheet's name; prettyprints all pending (unanswered) update requests'''
    sheet = gimme_sheet(sheet_name)
    url = 'https://api.smartsheet.com/2.0/sheets/{}/sentupdaterequests'.format(sheet[1])
    headers = {
        'authorization': "Bearer {}".format(access_token),
        'cache-control': "no-cache",
    }
    req = requests.request("GET", url, headers=headers)
    json_req = json.loads(req.text)
    pending_urs = json_req['data']
    pp.pprint(pending_urs)


def del_pending_urs(sheet_name):
'''Takes string of desired sheet's name; deletes all pending update requests on a sheet'''
    sheet = gimme_sheet(sheet_name)
    url = 'https://api.smartsheet.com/2.0/sheets/{}/sentupdaterequests'.format(sheet[1])
    headers = {
        'authorization': "Bearer {}".format(access_token),
        'cache-control': "no-cache",
    }
    req = requests.request("GET", url, headers=headers)
    json_req = json.loads(req.text)
    for item in json_req['data']:
        url2 = 'https://api.smartsheet.com/2.0/sheets/{}/sentupdaterequests/{}'.format(sheet[1], item['id'])
        req2 = requests.request("DELETE", url2, headers=headers)
        pp.pprint(item)


def gimme_all_sheets():
'''Takes no arguments; outputs dictionary of sheet metadata indexed by sheet name'''
    sheet_dict = {}
    req = ss.Sheets.list_sheets(include_all=True)
    json_req = json.loads(str(req))
    sheets = json_req['data']
    for sheet in sheets:
        sheet_dict[sheet['name']] = sheet
    return sheet_dict


def show_sheet(sheet_name):
'''Takes string of desired sheet's name; pretty prints json-formatted sheet'''
    req = ss.Sheets.list_sheets(include_all=True)
    sheets = req.data
    for sheetinfo in sheets:
        if sheetinfo.name == sheet_name:
            sheet = ss.Sheets.get_sheet(sheetinfo.id)
            count = sheet.total_row_count
            sheet = ss.Sheets.get_sheet(sheetinfo.id, page_size=count)
            jsonsheet = json.loads(str(sheet))
            pp.pprint(jsonsheet)


def gimme_column_codes(sheet_name):
'''Takes string of desired sheet's name; outputs dictionary of column ids indexed by column name'''
    columns = {}
    req = ss.Sheets.list_sheets(include_all=True)
    sheets = req.data
    for sheetinfo in sheets:
        if sheetinfo.name == sheet_name:
            sheet = ss.Sheets.get_sheet(sheetinfo.id)
            jsonsheet = json.loads(str(sheet))
            for column in jsonsheet['columns']:
                columns[column['title']] = column['id']
            return columns


def new_user(fname, lname, email, admin=False, licensed_sheet_creator=False):
'''Add a new user to your account by inputting first name, last name, and email address. Admin and Licensed Sheet Creator arguements are defaulted to False.'''
    req = ss.Users.add_user(
        smartsheet.models.User({
            'first_name': fname,
            'last_name': lname,
            'email': email,
            'admin': admin,
            'licensed_sheet_creator': licensed_sheet_creator
        })
    )
    pp.pprint(req)


def bulk_move_rows(source_sheet_name, dest_sheet_name, rows_to_move=20):
'''Provide the source sheet and destination sheet to move rows. This will move rows starting from the top. There is a defaulted row_to_move parameter set at 20.'''
    source_sheet = gimme_sheet(source_sheet_name)
    dest_sheet = gimme_sheet(dest_sheet_name)
    rowids = []
    req = json.loads(str(ss.Sheets.get_sheet(source_sheet[1], page_size=rows_to_move)))
    rows = req['rows']
    for row in rows:
	       rowids.append(row['id'])
    #Instantiates rows as CopyOrMoveRowDirective object. 'To' in the props is the destination sheet.
    rowobjs = ss.models.CopyOrMoveRowDirective({
	   'row_ids' : rowids,
	    'to' : ss.models.CopyOrMoveRowDestination({
		'sheet_id': dest_sheet[1]
		})
	})
    #Move rows request. Sheetid is the source sheet.
    req = ss.Sheets.move_rows(source_sheet[1], rowobjs, include=['attachments','discussions'], ignore_rows_not_found=False)
    pp.pprint('complete')


def link_column(source_sheet_name, source_col_name, dest_sheet_name, dest_col_name):
'''Takes strings of the source sheet, source column, destination sheet, and destination column; will perform cell linking of destination to source
Assumes data within the source sheet column are unique, where data in the destination column (of either unique or duplicate values) will be linked back to the source.'''

    def sheet_to_json_rows(sheet_id, col_id):
        sheet = ss.Sheets.get_sheet(sheet_id)
        count = sheet.total_row_count
        sheet = ss.Sheets.get_sheet(sheet_id, column_ids=[str(col_id)], page_size=count)
        json_sheet = json.loads(str(sheet))
        json_rows = json_sheet['rows']
        return json_rows

    data = {
        'source_sheet_id': None,
        'source_col_id': None,
        'dest_sheet_id': None,
        'dest_col_id': None,
        'source_rows': {}
    }

    source_search = gimme_sheet(source_sheet_name)
    data['source_sheet_id'] = source_search[1]
    dest_search = gimme_sheet(dest_sheet_name)
    data['dest_sheet_id'] = dest_search[1]

    source_column = ss.Sheets.get_column_by_title(data['source_sheet_id'], source_col_name)
    data['source_col_id'] = source_column.id
    dest_column = ss.Sheets.get_column_by_title(data['dest_sheet_id'], dest_col_name)
    data['dest_col_id'] = dest_column.id

    source_rows = sheet_to_json_rows(data['source_sheet_id'], data['source_col_id'])
    dest_rows = sheet_to_json_rows(data['dest_sheet_id'], data['dest_col_id'])

    for row in source_rows:
        value = row['cells'][0]['displayValue']
        if value in data['source_rows'].keys():
            pass
        else:
            data['source_rows'][value] = row['id']

    for row in dest_rows:
        value = row['cells'][0]['displayValue']
        dest_row_id = row['id']
        source_row_id = data['source_rows'].get(value)
        url = "https://api.smartsheet.com/2.0/sheets/{}/rows/{}".format(data['dest_sheet_id'], dest_row_id)
        mydata = {"cells":[{"columnId":data['dest_col_id'],"value":None,"linkInFromCell":{"sheetId":data['source_sheet_id'],"rowId":source_row_id,"columnId":data['source_col_id']}}]}
        data_json = json.dumps(mydata)

        headers = {
            'authorization': "Bearer {}".format(access_token),
            'content-type': "application/json",
            'cache-control': "no-cache",
            }
        response = requests.request("PUT", url, data=data_json, headers=headers)
        pp.pprint(response.text)
