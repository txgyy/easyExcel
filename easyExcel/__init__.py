from win32com.client import Dispatch
import pythoncom

class easyExcel:
    """A utility to make it easier to get at Excel.    Remembering
    to save the data is your problem, as is    error handling.
    Operates on one workbook at a time."""

    def __init__(self):                             # 打开文件或者新建文件（如果不存在的话）
        pythoncom.CoInitialize()
        self.xlBooks = {}
        self.xlApp = Dispatch('Excel.Application')
        self.xlApp.Visible = False
        self.xlApp.DisplayAlerts = False

    def save(self, oldbookname, newbookname=None):  # 保存文件
        if newbookname:
            self.xlBooks.get(oldbookname).SaveAs(newbookname)
        else:
            self.xlBooks.get(oldbookname).Save()

    def close(self, bookname):                      # 关闭文件
        book = self.xlBooks.pop(bookname)
        book.Close(SaveChanges=0)
        book = None
        while book != None:
            pythoncom.PumpWaitingMessages()

    def exit(self):                                 #退出程序
        self.xlApp.Quit()
        self.xlApp = None
        while self.xlApp != None:
            pythoncom.PumpWaitingMessages()
        pythoncom.CoUninitialize()

    def getCell(self, sheet, *args):                # 获取单元格的数据
        "Get value of one cell"
        if len(args) == 1:
            return sheet.Range(args[0]).Value
        elif len(args) == 2:
            return sheet.Cells(args[0],args[1]).Value
        else:
            raise Exception('参数数量错误')

    def setCell(self, sheet,value,*args):           # 设置单元格的数据
        "set value of one cell"
        if len(args) == 1:
            sheet.Range(args[0]).Value=value
        elif len(args) == 2:
            sheet.Cells(args[0],args[1]).Value=value
        else:
            raise Exception('参数数量错误')
    def delCell(self, sheet, *args):                # 删除单元格的数据
        "Get value of one cell"
        if len(args) == 1:
            sheet.Range(args[0]).ClearContents()
        elif len(args) == 1:
            sheet.Cells(args[0],args[1]).ClearContents()
        else:
            raise Exception('参数数量错误')

    def setCellformat(self, sheet, row, col):  # 设置单元格的数据
        "set value of one cell"
        sheet.Cells(row, col).Font.Size = 15  # 字体大小
        sheet.Cells(row, col).Font.Bold = True  # 是否黑体
        sheet.Cells(row, col).Name = "Arial"  # 字体类型
        sheet.Cells(row, col).Interior.ColorIndex = 3  # 表格背景
        # sht.Range("A1").Borders.LineStyle = xlDouble
        sheet.Cells(row, col).BorderAround(1, 4)  # 表格边框
        sheet.Rows(3).RowHeight = 30  # 行高
        sheet.Cells(row, col).HorizontalAlignment = -4131  # 水平居中xlCenter
        sheet.Cells(row, col).VerticalAlignment = -4160  #

    def getMaxRows(self, sheet):                    #得到行数
        return sheet.UsedRange.Rows.Count

    def getMaxCols(self, sheet):                    #得到列数
        return sheet.UsedRange.Columns.Count

    def getRow(self, sheet, index):                 #得到行数据
        return sheet.UsedRange.Rows(index).Value

    def getCol(selfsheet, sheet, index):            #得到列数据
        return sheet.UsedRange.Columns(index).Value

    def setRow(self, sheet, index, value):          #设置行数据
        sheet.UsedRange.Rows(index).Value = tuple(value)

    def setCol(selfsheet, sheet, index, value):     #设置列数据
        sheet.UsedRange.Columns(index).Value = tuple(value)

    def delRow(self, sheet, index):                 # 删除行
        sheet.Rows(index).Delete()

    def delCol(self, sheet, index):                 # 删除列
        sheet.Columns(index).Delete()

    def getRange(self, sheet, *args):               # 获得一块区域的数据，返回为一个二维元组
        "return a 2d array (i.e. tuple of tuples)"
        if len(args)==1:
            return sheet.Range(args[0]).Value
        elif len(args)==2:
            return sheet.Range(sheet.Cells(args[0][0],args[0][1]), sheet.Cells(args[1][0],args[1][1])).Value
        else:
            raise Exception('参数数量错误')

    def setRange(self, sheet, value,*args):         # 设置一块区域的数据
        "return a 2d array (i.e. tuple of tuples)"
        if len(args)==1:
            sheet.Range(args[0]).Value=tuple(value)
        elif len(args)==2:
            sheet.Range(sheet.Cells(args[0][0],args[0][1]), sheet.Cells(args[1][0],args[1][1])).Value=tuple(value)
        else:
            raise Exception('参数数量错误')

    def delRange(self, sheet, *args):               #删除区域
        if len(args)==1:
            sheet.Range(args[0]).ClearContents()
        elif len(args)==2:
            sheet.Range(sheet.Cells(args[0][0],args[0][1]), sheet.Cells(args[1][0],args[1][1])).ClearContents()
        else:
            raise Exception('参数数量错误')

    def cpRange(self, from_sheet, to_sheet, *args): #区域复制
        if len(args)==2:
            from_sheet.Range(args[0]).Copy()
            to_sheet.Range(args[1])
            to_sheet.Paste()
        elif len(args)==4:
            from_sheet.Range(from_sheet.Cells(args[0][0],args[0][1]), from_sheet.Cells(args[1][0],args[1][1])).Copy()
            to_sheet.Range(to_sheet.Cells(args[2][0], args[2][1]), to_sheet.Cells(args[3][0], args[3][1]))
            to_sheet.Paste()
        else:
            raise Exception('参数数量错误')

    def getSheetNames(self, book):              #得到该文件的所有工作簿名
        counts = book.Worksheets.Count
        names = [book.Worksheets(index).Name for index in range(1, counts + 1)]
        return names

    def chSheet(self, book, sheetname):         #选择工作簿
        counts = book.Worksheets.Count
        names = self.getSheetNames(book)
        sheet = book.Worksheets(sheetname) if (sheetname in names) or (sheetname in range(1, counts + 1)) \
            else book.Worksheets.Add(None, book.Worksheets(counts))
        if isinstance(sheetname,str):
            sheet.Name = sheetname
        return sheet

    def chSheets(self, book, sheetnames):       #选择多个工作簿
        counts = book.Worksheets.Count
        names = self.getSheetNames(book)
        sheets = list()
        for sheetname in sheetnames:
            sheet = book.Worksheets(sheetname) if (sheetname in names) or (sheetname in range(1, counts + 1))\
                else book.Worksheets.Add(None, book.Worksheets(counts))
            if isinstance(sheetname,str):
                sheet.Name = sheetname
            sheets.append(sheet)
        return sheets

    def cpSheet(self, from_sheet, to_sheet):    # 复制工作表
        "copy sheet"
        from_sheet.Copy(None, to_sheet)

    def delSheet(self, sheet):                   # 删除工作表
        sheet.Delete()

    def mvSheet(self, from_sheet, to_sheet):    # 移动工作表
        from_sheet.Move(None, to_sheet)

    def chBook(self, bookname):                 #选择文件
        xlBook = self.xlApp.Workbooks.Open(bookname) if bookname and os.path.isfile(bookname) \
            else self.xlApp.Workbooks.Add()
        self.xlBooks.update({bookname: xlBook})
        return xlBook

    def addPicture(self, sheet, pictureName, Left, Top, Width, Height):  # 插入图片
        "Insert a picture in sheet"
        sheet.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)