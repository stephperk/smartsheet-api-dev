import pyodbc, smartsheet, requests, json, ConfigParser, sys

column_mappings = {
    'TEXT_NUMBER': 'VARCHAR(200)',
    'CONTACT_LIST': 'VARCHAR(200)',
    'DATE': 'DATE',
    'PICKLIST': 'VARCHAR(200)',
    'CHECKBOX': 'VARCHAR(200)',
    'DATETIME': 'SMALLDATETIME'
}

def config_and_connect(env):
    config = ConfigParser.ConfigParser()
    if env == 'TEST':
        config.read(['config/pctest.cfg'])
        conn = pyodbc.connect('DRIVER={};SERVER={};DATABASE={};trusted_connection={}'.format(
            config.get('sql_server', 'driver'),
            config.get('sql_server', 'server'),
            config.get('sql_server', 'database'),
            config.get('sql_server', 'connection')
            )
        )
    elif env == 'PROD':
        config.read(['config/production.cfg'])
        conn = pyodbc.connect('DRIVER={};SERVER={};PORT={};DATABASE={};UID={};PWD={}'.format(
            config.get('sql_server', 'driver'),
            config.get('sql_server', 'server'),
            config.get('sql_server', 'port'),
            config.get('sql_server', 'database'),
            config.get('sql_server', 'user'),
            config.get('sql_server', 'password')
            )
        )
    cursor = conn.cursor()
    token = config.get('smartsheet', 'access_token')
    ss = smartsheet.Smartsheet(token)
    return ss, token, conn, cursor

def return_id(sheets, reports, name):
    for sheet in sheets.data:
        if sheet.name == name:
            return 'sheets', sheet.id
        else:
            for report in reports.data:
                if report.name == name:
                    return 'reports', report.id

def find_sheet_or_report(ss):
    print('Which Smartsheet would you like to upload to MS SQL Database?')
    name = raw_input()
    print('Fetching {}...'.format(name))
    sheets = ss.Sheets.list_sheets()
    reports = ss.Reports.list_reports()
    item_id = return_id(sheets, reports, name)
    return item_id

def find_row_count(item_id, ss):
    if item_id[0] == 'sheets':
        sheet = ss.Sheets.get_sheet(item_id[1])
        count = sheet.total_row_count
        return count
    else:
        report = ss.Reports.get_report(item_id[1])
        count = report.total_row_count
        return count

def get_sheet_or_report(item_id, count, token):
    url = 'https://api.smartsheet.com/2.0/{}/{}'.format(item_id[0], item_id[1])
    headers = {
        'authorization': 'Bearer {}'.format(token),
        'cache-control': 'no-cache'
    }
    payload = {'pageSize': str(count)}
    req = requests.request('GET', url, headers=headers, params=payload)
    json_resp = json.loads(req.text)
    return json_resp

def identify_columns(json_resp):
    columns = [(col['title'].replace(" ",""), col['type']) for col in json_resp['columns']]
    return columns

def clean_col_names(columns):
    cleaned_cols = []
    for col, col_type in columns:
        clean_title = '[{}]'.format(col)
        clean_tuple = clean_title, col_type
        cleaned_cols.append(clean_tuple)
    return cleaned_cols

def get_cell_value(cell):
    try:
        value = (cell['value'],)
    except KeyError:
        try:
            value = (cell['displayValue'],)
        except KeyError:
            value = (None,)
    return value

def setup_schema(columns, conn, cursor):
    schema = ""
    for col in columns:
	    col_schema = "{} {}, ".format(col[0],column_mappings.get(col[1]))
	    schema += col_schema
    schema = schema[:-2]
    print(schema)
    cursor.execute("""
    IF OBJECT_ID('smartsheet', 'U') IS NOT NULL
        DROP TABLE smartsheet
    CREATE TABLE smartsheet ({})
    """.format(schema))
    conn.commit()
    print('Schema produced...')

def insert_into_table(json_resp, cleaned_cols, conn, cursor):
    qmarks = ""
    for i in range(0, len(cleaned_cols) - 1):
        qmarks += "?, "
    rows_to_insert = []
    for row in json_resp['rows']:
        row_values = ()
        for cell in row['cells']:
            row_tuple = get_cell_value(cell)
            row_values += row_tuple
        rows_to_insert.append(row_values)
    cursor.executemany(
        "INSERT INTO smartsheet VALUES ({}?)".format(qmarks), rows_to_insert)
    conn.commit()
    print('Table updated; please refresh to view data!')


def main():
    env = sys.argv[1]
    ss, token, conn, cursor = config_and_connect(env)
    item_id = find_sheet_or_report(ss)
    count = find_row_count(item_id, ss)
    json_resp = get_sheet_or_report(item_id, count, token)
    columns = identify_columns(json_resp)
    cleaned_cols = clean_col_names(columns)
    setup_schema(cleaned_cols, conn, cursor)
    insert_into_table(json_resp, cleaned_cols, conn, cursor)


if __name__ == '__main__':
    main()
