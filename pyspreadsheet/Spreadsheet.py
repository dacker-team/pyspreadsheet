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

GET_INFO_DEFAULT_ARGS = {
    "fr_to_us_date": False,
    "avoid_lines": None,
    "transform_comma": False,
    "table_name_from_key": False,
    "format_date_from": None,
    "list_col_to_remove": None,
    "special_table_name": None,
    "remove_comma": False,
    "treat_int_column": False,
    "remove_comma_float": False,
    "replace": True,
    "unformatting": False
}


def _get_args(key_config, param, dict_param):
    if dict_param.get(param) is not None:
        return dict_param.get(param)
    return key_config.get(param) if key_config.get(param) is not None else GET_INFO_DEFAULT_ARGS.get(param)


class Spreadsheet:
    def __init__(self,
                 googleauthentication: GoogleAuthentication,
                 dbstream: DBStream = None,
                 dbstream_spreadsheet_schema_name=None,
                 data_collection_config_path=None):
        self.googleauthentication = googleauthentication
        self.dbstream = dbstream
        self.account = googleauthentication.get_account("sheets").spreadsheets()
        self.dbstream_spreadsheet_schema_name = "spreadsheet" if dbstream_spreadsheet_schema_name is None \
            else dbstream_spreadsheet_schema_name
        self.data_collection_config_path = data_collection_config_path

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
        return None

    @staticmethod
    def get_columns_name(row):
        columns_name = []
        for i in row:
            if i != '':
                column_name = str \
                    .lower(str(i)) \
                    .strip() \
                    .replace(" ", "_") \
                    .replace("(", "") \
                    .replace(")", "") \
                    .replace("\n", "_") \
                    .replace("/", "_") \
                    .replace("'s", "") \
                    .replace("-", "_") \
                    .replace(".", "_")
                if column_name in columns_name:
                    column_name = column_name + "_%s" % (str(columns_name.count(column_name) + 1))
                columns_name.append(column_name)
            elif i == '':
                column_name = "__blank_header_" + str(len(columns_name) + 1)
                if column_name in columns_name:
                    column_name = column_name + "%s" % (str(columns_name.count(column_name) + 1))
                columns_name.append(column_name)

        for i in range(0, len(columns_name)):
            if "__blank_header_" in columns_name[i]:
                t = 0
                for y in range(i, len(columns_name)):
                    if not ("__blank_header_" in columns_name[y]):
                        t = 1
                        break
                if t == 0:
                    columns_name = columns_name[:i]
                    break

        return columns_name

    def get_last_spreadsheet_update_time(self, spreadsheet_id):
        drive_api_service = self.googleauthentication.get_account("drive", "v3")
        r = drive_api_service.files().get(
            fileId=spreadsheet_id,
            fields="lastModifyingUser,modifiedTime"
        ).execute()
        return r["modifiedTime"], r["lastModifyingUser"].get("emailAddress")

    def get_info_from_worksheet(self, key, **kwargs):
        config = yaml.load(open(self.data_collection_config_path), Loader=yaml.FullLoader)
        key_config = config[key]
        worksheet_name = key_config['worksheet_name']
        spreadsheet_id = key_config['sheet_id']
        last_spreadsheet_update_time, last_spreadsheet_update_by = self.get_last_spreadsheet_update_time(
            spreadsheet_id)
        print(last_spreadsheet_update_time, last_spreadsheet_update_by)
        loaded_sheet_table = "_loaded_sheets"
        max_in_datamart = self.dbstream.get_max(
            schema=self.dbstream_spreadsheet_schema_name,
            table=loaded_sheet_table, field="last_spreadsheet_update_time",
            filter_clause="WHERE worksheet_name='%s' and last_spreadsheet_update_time='%s'"
                          % (worksheet_name, last_spreadsheet_update_time))
        print(max_in_datamart)
        if max_in_datamart:
            return 0
        wks = self._get_worksheets_by_id(spreadsheet_id, worksheet_name)
        table_name_from_key = _get_args(key_config=key_config, param="table_name_from_key", dict_param=kwargs)
        if table_name_from_key:
            table_name = self.dbstream_spreadsheet_schema_name + "." + key
        else:
            table_name = self.dbstream_spreadsheet_schema_name + "." + worksheet_name.replace(" - ", "_").replace(
                "-", "_").replace(" ", "_").lower()
        try:
            avoid_lines = _get_args(key_config=key_config, param="avoid_lines", dict_param=kwargs)
            if avoid_lines:
                wks_list = []
                for row in wks:
                    wks_list.append(row)
                wks = wks_list[avoid_lines:]
            columns_names = self.get_columns_name(wks[0])

            transform_comma = _get_args(key_config=key_config, param="transform_comma", dict_param=kwargs)
            remove_comma = _get_args(key_config=key_config, param="remove_comma", dict_param=kwargs)
            remove_comma_float = _get_args(key_config=key_config, param="remove_comma_float", dict_param=kwargs)
            fr_to_us_date = _get_args(key_config=key_config, param="fr_to_us_date", dict_param=kwargs)
            special_table_name = _get_args(key_config=key_config, param="special_table_name", dict_param=kwargs)
            format_date_from = _get_args(key_config=key_config, param="format_date_from", dict_param=kwargs)
            list_col_to_remove = _get_args(key_config=key_config, param="list_col_to_remove", dict_param=kwargs)
            treat_int_column = _get_args(key_config=key_config, param="treat_int_column", dict_param=kwargs)
            unformatting = _get_args(key_config=key_config, param="unformatting", dict_param=kwargs)

            if unformatting:
                _l = 0
                for row in wks:
                    _l += 1
                wks = wks.get_values((1, 1), (_l, len(columns_names)), value_render="UNFORMATTED_VALUE",
                                     date_time_render_option="FORMATTED_STRING")
                for row in wks:
                    for i in range(len(columns_names)):
                        row[i] = str(row[i])

            _etl___loaded_at__ = str(datetime.now())
            rows = []
            c = 0
            for row in wks:
                if c == 0:
                    c = 1
                    continue
                for i in range(len(columns_names)):
                    if "€" in row[i]:
                        row[i] = row[i].replace(" ", "").replace("€", "")
                    if "#DIV/0" in row[i]:
                        row[i] = None
                    elif "#N/A" in row[i]:
                        row[i] = None
                    elif "#REF!" in row[i]:
                        row[i] = None
                    elif "#NAME?" in row[i]:
                        row[i] = None
                    elif treat_int_column and (row[i] == "NA" or row[i] == '?'):
                        row[i] = None
                    elif row[i].replace(" ", "") == "":
                        row[i] = None
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
                                    row[i] = row[i].replace(" ", "")
                                    row[i] = datetime.strptime(row[i], "%d/%m/%Y")
                                except:
                                    try:
                                        row[i] = row[i].replace(" ", "")
                                        row[i] = datetime.strptime(row[i], "%d/%m/%y")
                                    except:
                                        row[i] = datetime.strptime(row[i][:10], "%Y-%m-%d")
                    if format_date_from:
                        if "date" in columns_names[i]:
                            if row[i] and row[i] != "":
                                row[i] = datetime.strptime(row[i], format_date_from)
                row = row[:len(columns_names)]
                row.append(_etl___loaded_at__)
                row = tuple(row)
                rows.append(list(map(lambda x: value_or_none(x), row)))
            if special_table_name is not None:
                table_name = self.dbstream_spreadsheet_schema_name + "." + special_table_name.replace(" ", "_")
            columns_names.append("_etl___loaded_at__")
            result = {
                "table_name": table_name,
                "columns_name": columns_names,
                "rows": rows,
            }
            if list_col_to_remove:
                result = remove_col(result, list_col_to_remove)

            replace = _get_args(key_config=key_config, param="replace", dict_param=kwargs)
            self.dbstream.send_data(result, replace=replace)
            self.dbstream.send_data(
                {
                    "table_name": self.dbstream_spreadsheet_schema_name + "." + loaded_sheet_table,
                    "columns_name": ["spreadsheet_id", "worksheet_name", "last_spreadsheet_update_time",
                                     "last_spreadsheet_update_by"],
                    "rows": [
                        [spreadsheet_id, worksheet_name, last_spreadsheet_update_time, last_spreadsheet_update_by]],
                }, replace=False
            )
            print('table %s created' % table_name)
            return result
        except Exception as e:
            raise Exception(
                "error to treat and/or send %s which ID is %s in schema %s : %s" % (
                    str(worksheet_name), str(spreadsheet_id), str(table_name), str(e))
            )

    def get_info_from_worksheets(self, **kwargs):
        config = yaml.load(open(self.data_collection_config_path), Loader=yaml.FullLoader)
        dict_error = dict()
        for key in config:
            try:
                self.get_info_from_worksheet(key, **kwargs)
            except Exception as e:
                dict_error[key] = str(e)
                pass
        if len(dict_error) != 0:
            raise Exception(dict_error)
