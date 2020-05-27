import argparse
import xlrd
import datetime

parser = argparse.ArgumentParser(description='Excel file to sql script')
# parser.add_argument('-l', metavar='l', type=str, nargs='?', help='link to Google Sheet', dest='file_link')
parser.add_argument('-f', metavar='f', type=str, nargs='?', help='path to file', dest='file_path')
parser.add_argument('-d', metavar='d', type=str, nargs=1, help='path to destination file', dest='dest_file_path')
# parser.add_argument('-s', metavar='s', type=str, nargs='?', help='sql dialect', dest='sql_d')
args = parser.parse_args()

sql_d = 'my'
cast = {
    "my": "UNSIGNED",
    "pg": "INTEGER"
}


def output_insert_sql(excel_path, output_path):
    book = xlrd.open_workbook(excel_path)
    out_f = open(output_path, "w")

    for sheet_name in book.sheet_names():
        table_name = sheet_name
        sheet = book.sheet_by_name(sheet_name)
        out_f.write(f"\n-- {table_name}\n")

        column_types = []
        for col in range(sheet.ncols):
            column_types.append(sheet.cell(0, col).value)
        column_names = []
        for col in range(sheet.ncols):
            column_names.append(sheet.cell(1, col).value)

        sql_prefix = f"INSERT INTO {table_name} ({','.join(column_names)}) \n\tVALUES("
        sql_suffix = ");\n"

        for row in range(2, sheet.nrows):
            value_list = [None] * sheet.ncols
            for col in range(sheet.ncols):
                val = sheet.cell(row, col).value
                if column_types[col] == 'int':
                    value_list[col] = fr"CAST('{int(val)}' as {cast.get(sql_d)})"
                elif column_types[col] == 'date':
                    date = datetime.datetime.utcfromtimestamp((val - 25569) * 86400.0).strftime('%Y-%m-%d')
                    value_list[col] = fr"CAST('{date}' as date)"
                elif column_types[col] == 'float':
                    value_list[col] = f"CAST('{val}' as datetime)"
                elif column_types[col] != 'float' and type(val) == float:
                    value_list[col] = f"'{str(val).replace('.0','')}'"
                else:
                    value_list[col] = str(fr"'{val}'")
            sql = sql_prefix + ','.join(value_list) + sql_suffix

            out_f.write(sql)
    out_f.close()


output_insert_sql(args.file_path, args.dest_file_path[0])
