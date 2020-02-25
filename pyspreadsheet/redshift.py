import decimal
import pyred

from . import core
import datetime


def query_to_sheet(project, sheet_id, worksheet_name, instance, query):
    seconds = 5
    try_query_to_sheet(project, sheet_id, worksheet_name, instance, query, seconds)


def try_query_to_sheet(project, sheet_id, worksheet_name, instance, query, seconds):
    try:
        items = pyred.execute.execute_query(instance, query)
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
        core.send_to_sheet(project, sheet_id, data)
    except Exception as e:
        if seconds <= 20:
            print("Waiting %s seconds..." % seconds)
            datetime.time.sleep(seconds)
            try_query_to_sheet(project, sheet_id, worksheet_name, instance, query, seconds * 2)
        else:
            raise e
    return 0
