def construct_range(worksheet_name, start, end=None):
    range = "'" + worksheet_name + "'" + "!" + start
    if end:
        range = range + ":" + end
    return range


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
