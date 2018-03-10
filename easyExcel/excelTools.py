from easyExcel import easyExcel
import os
def groupBooks(path, destination, all_copy=True):
    listdir = os.listdir(path)
    if destination in listdir:
        listdir.remove(destination)
    destination = os.path.join(path, destination)

    Excel = easyExcel()
    try:
        to_book = Excel.chBook(destination)
        for index, file in enumerate(listdir):
            xlspath = os.path.join(path, file)
            from_book = Excel.chBook(xlspath)
            to_sheet = Excel.chSheet(to_book, index + 1)
            from_sheet = Excel.chSheet(from_book, 1)

            if all_copy:
                ##复制工作簿
                Excel.cpSheet(from_sheet, to_sheet)
            else:
                ##复制区域
                start = [1, 1]
                end = [Excel.getMaxRows(from_sheet), Excel.getMaxCols(from_sheet)]
                Excel.cpRange(from_sheet, to_sheet, start, end, start, end)

            Excel.close(xlspath)
            to_sheet = None
            from_sheet = None
        Excel.save(destination, destination)
        Excel.close(destination)
    except Exception as e:
        print(e)
    finally:
        Excel.exit()

def groupSheets(path,destination,from_start,from_end):
    destination = os.path.join(path, destination)
    Excel = easyExcel()
    try:
        book = Excel.chBook(destination)
        to_sheet = Excel.chSheet(book, 1)
        values = set()
        from_values = list()
        from_sheets = Excel.chSheets(book, range(2, book.Worksheets.Count + 1))
        for from_sheet in from_sheets:
            from_value = Excel.getRange(from_sheet, from_start, from_end)
            from_values.append(from_value)
            from_sheet = None

        flag = True
        com_value = set()
        for from_value in from_values:
            values.update(from_value)
            if flag:
                com_value.update(from_value)
                flag = False
            else:
                com_value.intersection_update(from_value)
        row = len(com_value)
        Excel.setRange(to_sheet, com_value, from_start, [row, from_end[1]])

        diff_values = values.difference(com_value)
        diff_values = [value for value in diff_values if None not in value]
        diff_values = sorted(diff_values,key=lambda x:x[0])
        Excel.setRange(to_sheet, diff_values, [row+1, from_start[1]], [row+len(diff_values), from_end[1]])

        to_sheet = None
        Excel.save(destination,destination)
        Excel.close(destination)
    except Exception as e:
        print(e)
    finally:
        Excel.exit()

def groupBookstoOne(path, destination, from_start,from_end):
    listdir = os.listdir(path)
    if destination in listdir:
        listdir.remove(destination)
    destination = os.path.join(path, destination)

    Excel = easyExcel()
    try:
        from_values = list()
        for file in listdir:
            xlspath = os.path.join(path, file)
            from_book = Excel.chBook(xlspath)
            from_sheet = Excel.chSheet(from_book, 1)
            from_values.append(Excel.getRange(from_sheet,from_start,from_end))
            from_sheet = None
            Excel.close(xlspath)
        values = set()
        flag = True
        com_value = set()
        for from_value in from_values:
            values.update(from_value)
            if flag:
                com_value.update(from_value)
                flag = False
            else:
                com_value.intersection_update(from_value)
        row = len(com_value)
        to_book = Excel.chBook(destination)
        to_sheet = Excel.chSheet(to_book,1)
        Excel.setRange(to_sheet, com_value, from_start, [row, from_end[1]])

        diff_values = values.difference(com_value)
        diff_values = [value for value in diff_values if None not in value]
        diff_values = sorted(diff_values,key=lambda x:x[0])
        Excel.setRange(to_sheet, diff_values, [row+1, from_start[1]], [row+len(diff_values), from_end[1]])

        to_sheet = None
        Excel.save(destination,destination)
        Excel.close(destination)
    except Exception as e:
        print(e)
    finally:
        Excel.exit()