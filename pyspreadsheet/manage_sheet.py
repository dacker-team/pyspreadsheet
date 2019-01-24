from . import init_connection


def get_row_count(project, sheet_id, worksheet_name, account=None):
    info = get_sheet_info(project, sheet_id, account=account)
    for i in info["sheets"]:
        properties = i["properties"]
        if properties["title"] == worksheet_name:
            return properties["gridProperties"]["rowCount"]
    return 0


def get_sheet_info(project, sheet_id, account=None):
    if not account:
        account = init_connection.get_api_account(project)
    sheet = account.get(
        spreadsheetId=sheet_id
    ).execute()
    return sheet


def find_worksheet(project, sheet_id, worksheet_name, account=None):
    sheet = get_sheet_info(project, sheet_id, account=account)
    wks_list = sheet.get("sheets")
    for wks in wks_list:
        if wks.get("properties").get("title").lower() == worksheet_name.lower():
            return wks
    print("This worksheet doesn't exist")
    return None


def add_worksheet(project, sheet_id, worksheet_name, account=None):
    if not account:
        account = init_connection.get_api_account(project)
    requests = [
        {
            "addSheet": {
                "properties": {
                    "title": worksheet_name,
                    "tabColor": {
                        "red": 0.0,
                        "green": 0.0,
                        "blue": 1.0
                    }
                }
            }
        }
    ]
    body = {
        'requests': requests
    }
    response = account.batchUpdate(spreadsheetId=sheet_id,
                                   body=body).execute()
    return response
