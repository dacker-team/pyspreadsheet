import datetime

import decimal
import googleapiclient
from googleauthentication import GoogleAuthentication
from dbstream import DBStream

from pyspreadsheet.core import write
from pyspreadsheet.core.manage_sheet import add_worksheet

from pyspreadsheet.core.read import check_availability_column


class Spreadsheet:
    def __init__(self, googleauthentication: GoogleAuthentication, dbstream: DBStream = None):
        self.googleauthentication = googleauthentication
        self.dbstream = dbstream
        self.account = googleauthentication.get_account("sheets").spreadsheets()

    def send(self, sheet_id, data):
        try:
            available_response = check_availability_column(self, sheet_id, data)
        except googleapiclient.errors.HttpError:
            add_worksheet(self, sheet_id, data["worksheet_name"])
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
        write.write(self, sheet_id, data)
        return 0

    def query_to_sheet(self, sheet_id, worksheet_name, query):
        if not self.dbstream:
            print("No DBStream set")
            exit()
        items = self.dbstream.execute_query(query)
        if not items:
            return 0
        columns_name = list(items[0].keys())
        rows = [list(i.values()) for i in items]
        for row in rows:
            for i in range(len(row)):
                r = row[i]
                if isinstance(r, datetime.datetime):
                    row[i] = str(r)
                if isinstance(r, decimal.Decimal):
                    row[i] = float(r)
        data = {
            "worksheet_name": worksheet_name,
            "columns_name": columns_name,
            "rows": rows
        }
        self.send(sheet_id, data)
        return 0
