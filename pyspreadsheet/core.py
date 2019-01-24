import googleapiclient

from . import write
from .manage_sheet import add_worksheet
from .read import check_availability_column


def send_to_sheet(project, sheet_id, data):
    try:
        available_response = check_availability_column(project, sheet_id, data)

    except googleapiclient.errors.HttpError:
        add_worksheet(project, sheet_id, data["worksheet_name"])
        available_response = {
            "completely_available": True
        }
    available = available_response["completely_available"]
    if not available:
        user_response = input("Columns " +
                              ", ".join(available_response["warning_columns"]) +
                              " will be overwritten, do you want to continue ?" +
                              "\nType yes to continue\n")
        if user_response.lower() not in ["yes", "y"]:
            return 0
    write.write(project, sheet_id, data)
    return 0


def send_to_sheet_from_account(account, sheet_id, data):
    project = "DATA_COCKPIT"  # Just to know
    try:
        available_response = check_availability_column(project, sheet_id, data, account=account)

    except googleapiclient.errors.HttpError:
        add_worksheet(project, sheet_id, data["worksheet_name"], account=account)
        available_response = {
            "completely_available": True
        }
    available = available_response["completely_available"]
    if not available:
        user_response = input("Columns " +
                              ", ".join(available_response["warning_columns"]) +
                              " will be overwritten, do you want to continue ?" +
                              "\nType yes to continue\n")
        if user_response.lower() not in ["yes", "y"]:
            return 0
    write.write(project, sheet_id, data,account=account)
    return 0
