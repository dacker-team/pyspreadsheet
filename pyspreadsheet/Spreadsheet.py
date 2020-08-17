import string
from datetime import datetime
from time import sleep
import decimal
import os
import copy
import json

import pygsheets
import requests
import googleapiclient
import yaml
from googleauthentication import GoogleAuthentication
from dbstream import DBStream

from pyspreadsheet.core import write
from pyspreadsheet.core.manage_sheet import add_worksheet

from pyspreadsheet.core.read import check_availability_column
from pyspreadsheet.core.tool import value_or_none, remove_col


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
        data_copy = copy.deepcopy(data)
        url = os.environ.get("MONITORING_EXPORT_URL")
        if url:
            body = {
                "dbstream_instance_id": self.googleauthentication.dbstream.dbstream_instance_id,
                "instance_name": self.googleauthentication.dbstream.instance_name,
                "client_id": self.googleauthentication.dbstream.client_id,
                "instance_type_prefix": self.googleauthentication.dbstream.instance_type_prefix,
                "sheet_id": sheet_id,
                "worksheet_name": data_copy["worksheet_name"],
                "nb_rows": len(data_copy["rows"]),
                "nb_columns": len(data_copy["columns_name"]),
                "timestamp": str(datetime.now()),
                "ssh_tunnel": True if self.googleauthentication.dbstream.ssh_tunnel else False,
                "local_absolute_path": os.getcwd()
            }
            r = requests.post(url=url, data=json.dumps(body))
            print(r.status_code)
        return 0

    def _try_query_to_sheet(self, sheet_id, worksheet_name, query, seconds, beautiful_columns_name):
        try:
            items = self.dbstream.execute_query(query)
            if not items:
                return 0
            columns_name = list(items[0].keys())
            rows = [list(i.values()) for i in items]
            for row in rows:
                for i in range(len(row)):
                    r = row[i]
                    if isinstance(r, datetime):
                        row[i] = str(r)
                    if isinstance(r, decimal.Decimal):
                        row[i] = float(r)
            data = {
                "worksheet_name": worksheet_name,
                "columns_name": columns_name if not beautiful_columns_name else [string.capwords(
                    c.replace("_", " ")
                ).replace("Arr", "ARR").replace("Mrr", "MRR").replace("Tcv", "TCV").replace("Acv", "ACV") for
                                                                                 c in columns_name],
                "rows": rows
            }
            self.send(sheet_id, data)
        except Exception as e:
            if seconds <= 20:
                print(str(e))
                print("Waiting %s seconds..." % seconds)
                sleep(seconds)
                self._try_query_to_sheet(sheet_id, worksheet_name, query, seconds * 2, beautiful_columns_name)
            else:
                raise e
        return 0

    def query_to_sheet(self, sheet_id, worksheet_name, query, beautiful_columns_name=False):
        seconds = 5
        self._try_query_to_sheet(sheet_id=sheet_id,
                                 worksheet_name=worksheet_name,
                                 query=query,
                                 seconds=seconds,
                                 beautiful_columns_name=beautiful_columns_name)

    # NEW

    def _get_worksheets_by_id(self, sheet_id, worksheet_name):
        ps = pygsheets.client.Client(credentials=self.googleauthentication.credentials())
        sh = ps.open_by_key(sheet_id)
        wks_list = sh.worksheets()
        for wks in wks_list:
            if wks.title.lower() == worksheet_name.lower():
                return wks
        print("This worksheet doesn't exist")
        exit()
        return None

    def get_column_names(self, row):
        result = []
        for i in row:
            if i != '':
                column_name = str \
                    .lower(str(i)) \
                    .replace(" ", "_") \
                    .replace("(", "") \
                    .replace(")", "") \
                    .replace("\n", "_") \
                    .replace("/", "_") \
                    .replace("'s", "") \
                    .replace("-", "_")

                if column_name in result:
                    column_name = column_name + "_%s" % (str(result.count(column_name) + 1))

                result.append(column_name)
        return result

    def get_info_from_worksheets(self, config_path, fr_to_us_date=False, avoid_lines=None,
                                 transform_comma=False,
                                 format_date_from=None, list_col_to_remove=None, special_table_name=None,
                                 remove_comma=False, treat_int_column=False, remove_comma_float=False):
        config = yaml.load(open(config_path), Loader=yaml.FullLoader)
        for key in config:
            worksheet_name = config[key]['worksheet_name']
            wks = self._get_worksheets_by_id(config[key]['sheet_id'], worksheet_name)
            table_name = "spreadsheet." + worksheet_name.replace(" ", "_").lower()

            rows = []
            c = 0
            if avoid_lines:
                wks_list = []
                for row in wks:
                    wks_list.append(row)
                wks = wks_list[avoid_lines:]
            columns_names = self.get_column_names(wks[0])

            for row in wks:
                if c > 0:
                    for i in range(len(columns_names)):
                        if "€" in row[i]:
                            row[i] = row[i].replace(" ", "").replace("€", "")
                        if "#DIV/0" in row[i]:
                            row[i] = None
                        if row[i] == "#N/A":
                            row[i] = None
                        if row[i] == "#REF!":
                            row[i] = None
                        if treat_int_column and row[i] == "NA" or row[i] == '?':
                            row[i] = None
                        try:
                            row[i] = int(row[i].replace("\u202f", ""))
                        except Exception as e:
                            pass
                        if transform_comma:
                            try:
                                row[i] = float(row[i].replace(",", ".").replace("\u202f", ""))
                            except Exception as e:
                                pass
                        if remove_comma:
                            try:
                                row[i] = int(row[i].replace(",", ""))
                            except:
                                pass
                        if remove_comma_float:
                            try:
                                row[i] = float(row[i].replace(",", ""))
                            except:
                                pass
                        if fr_to_us_date:
                            if "date" in columns_names[i]:
                                if row[i] and row[i] != "":
                                    try:
                                        row[i] = datetime.strptime(row[i], "%d/%m/%Y")
                                    except:
                                        try:
                                            row[i] = datetime.strptime(row[i], "%d/%m/%y")
                                        except:
                                            row[i] = datetime.strptime(row[i][:10], "%Y-%m-%d")
                        if format_date_from:
                            if "date" in columns_names[i]:
                                if row[i] and row[i] != "":
                                    row[i] = datetime.strptime(row[i], format_date_from)

                    row = tuple(row[:len(columns_names)])
                    rows.append(list(map(lambda x: value_or_none(x), row)))
                c = 1
            if special_table_name:
                table_name = "spreadsheet." + special_table_name.replace(" ", "_")
            result = {
                "table_name": table_name,
                "columns_name": columns_names,
                "rows": rows,
            }
            if list_col_to_remove:
                result = remove_col(result, list_col_to_remove)

            self.dbstream.send_data(result)
            print('table %s created' % table_name)
