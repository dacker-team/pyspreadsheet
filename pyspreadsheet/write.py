from pyspreadsheet.manage_sheet import get_row_count
from .tool import construct_range, delete_none
from . import init_connection


def write(project, sheet_id, data, account=None):
    worksheet_name = data["worksheet_name"]
    columns_name = data["columns_name"]
    rows = data["rows"]
    rows = delete_none(rows)
    values = [columns_name] + rows

    body = {
        "value_input_option": "USER_ENTERED",
        "data": [

            {
                "range": construct_range(worksheet_name, 'A1'),
                "majorDimension": "ROWS",
                "values": values,
            }
        ]
    }
    if not account:
        account = init_connection.get_api_account(project)
    response = account.values().batchUpdate(spreadsheetId=sheet_id,
                                            body=body).execute()
    row_count = get_row_count(project, sheet_id, worksheet_name, account=account)
    if row_count > len(values):
        clean_row = ["" for r in columns_name]
        clean_rows = [clean_row for i in range(row_count - len(values))]
        clean_body = {
            "value_input_option": "USER_ENTERED",
            "data": [

                {
                    "range": construct_range(worksheet_name, 'A' + str(len(values) + 1)),
                    "majorDimension": "ROWS",
                    "values": clean_rows,
                }
            ]
        }
        account.values().batchUpdate(spreadsheetId=sheet_id,
                                     body=clean_body).execute()

    return response
