from .tool import construct_range
from . import init_connection


def write(project, sheet_id, data):
    worksheet_name = data["worksheet_name"]
    columns_name = data["columns_name"]
    rows = data["rows"]
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
    account = init_connection.get_api_account(project)
    response = account.values().batchUpdate(spreadsheetId=sheet_id,
                                            body=body).execute()
    return response
