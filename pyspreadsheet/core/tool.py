def construct_range(worksheet_name, start, end=None):
    range_ = "'" + worksheet_name + "'" + "!" + start
    if end:
        range_ = range_ + ":" + end
    return range_


def column_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def delete_none(rows):
    for i in range(len(rows)):
        for j in range(len(rows[i])):
            if rows[i][j] is None:
                rows[i][j] = ""
    return rows


def value_or_none(value):
    if not value == "":
        return value


def remove_col(result, list_col_to_remove):
    all_index = [result['columns_name'].index(col) for col in list_col_to_remove]
    for row in result['rows']:
        for i in all_index:
            del row[i]
    for i in all_index:
        del result['columns_name'][i]
    return result
