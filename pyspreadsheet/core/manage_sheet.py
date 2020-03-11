

def get_row_count(spreadsheet, sheet_id, worksheet_name):
    info = get_sheet_info(spreadsheet, sheet_id)
    for i in info["sheets"]:
        properties = i["properties"]
        if properties["title"] == worksheet_name:
            return properties["gridProperties"]["rowCount"]
    return 0


def get_sheet_info(spreadsheet, sheet_id):
    sheet = spreadsheet.account.get(
        spreadsheetId=sheet_id
    ).execute()
    return sheet


def find_worksheet(spreadsheet, sheet_id, worksheet_name):
    sheet = get_sheet_info(spreadsheet, sheet_id)
    wks_list = sheet.get("sheets")
    for wks in wks_list:
        if wks.get("properties").get("title").lower() == worksheet_name.lower():
            return wks
    print("This worksheet doesn't exist")
    return None


def add_worksheet(spreadsheet, sheet_id, worksheet_name):
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
    response = spreadsheet.account.batchUpdate(spreadsheetId=sheet_id, body=body).execute()
    return response
