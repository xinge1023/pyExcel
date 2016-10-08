#(author:xinge1023@163.com)
#(date:2016-08-15)
import win32com,os
from win32com.client import Dispatch, constants

#※※※※※※※※※※※※※※※※(变量定义)※※※※※※※※※※※※※※※※#
#(在2007下测试正常)
excelApp = win32com.client.Dispatch('Excel.Application')    #excel应用
excel = None                #类似ActiveWorkbook,当前处理的xls
Workbooks = None            #Workbook集合
ActiveWorkbook = None       #当前活动的Workbook
Worksheets = None           #Worksheet集合
ActiveSheet = None          #当前活动的WorkSheet
Selection = None            #已选择对象
showTips = True             #是否打印提示信息
#※※※※※※※※※※※※※※※※(方法定义)※※※※※※※※※※※※※※※※#
#新建excel文件
#isMakedirs(是否自动创建目录)isCover(是否自动覆盖旧文件) 
def newExcel(fileaddr,isMakedirs=False,isCover=False):
    global excelApp,excel
    excel = excelApp.Workbooks.Add()
    fileSplit = fileaddr.split("\\")
    fileaddr = ""
    for i in range(len(fileSplit)-1):
        fileaddr += fileSplit[i] + "/"
    leastSplit = fileSplit[len(fileSplit)-1]
    leastSplit = leastSplit.split("\\")
    leastSplit = str(leastSplit).replace("\\","/").replace("/x0","/")
    leastSplit = leastSplit.replace("['","").replace("']","") 
    fileaddr +=  leastSplit
    if(isCover or not os.path.exists(fileaddr)):
        fileDir = os.path.split(fileaddr)[0]
        if(isMakedirs and not os.path.exists(fileDir)):
            os.makedirs(fileDir)
        try:
            excel.SaveAs(fileaddr)
        except:
            print("excel文件保存异常.")
    
#打开excel文件
def openExcel(fileaddr):
    global excelApp,excel
    fileSplit = fileaddr.split("\\")
    fileaddr = ""
    for i in range(len(fileSplit)-1):
        fileaddr += fileSplit[i] + "/"
    leastSplit = fileSplit[len(fileSplit)-1]
    leastSplit = leastSplit.split("\\")
    leastSplit = str(leastSplit).replace("\\","/").replace("/x0","/")
    leastSplit = leastSplit.replace("['","").replace("']","") 
    fileaddr +=  leastSplit
    if(os.path.exists(fileaddr)):
        excel = excelApp.Workbooks.Open(fileaddr)
        excelInit()
    else:
        print("文件打开失败,请检查路径是否正确("+str(fileaddr)+")")


#初始化
def excelInit():
    global excelApp,Workbooks,ActiveWorkbook,Worksheets,ActiveSheet,Selection
    Workbooks = excelApp.Workbooks
    ActiveWorkbook = excelApp.ActiveWorkbook
    Worksheets = excelApp.ActiveWorkbook.Worksheets
    ActiveSheet = ActiveWorkbook.ActiveSheet
    Selection = excelApp.Selection
    print("当前Excel版本:"+getExcelVersion())

# 后台运行,不显示,不警告
def setExcelVisible(visible=True,displayAlerts=False):
    #运行情况可见(1为可见,0为不可见)
    excelApp.Visible = visible
    #屏蔽提示(1为显示,0为不显示)
    excelApp.DisplayAlerts = displayAlerts

#列名转换列编号
def getColumnId(columnsName="A"):
    if(type(columnsName) == type(1)):
        return columnsName
    columnsName = str(columnsName).upper()
    letter = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    letterDict = {}
    one = [x  for x in letter ]
    two = [x + y for x in letter for y in letter]
    three = [x + y + z for x in letter for y in letter for z in letter]
    allLetter = one + two + three
    for columnId,columnName in enumerate(allLetter):
        letterDict[columnName] = (columnId+1)
    return letterDict[columnsName]

#列编号转换列名
def getColumnName(columnsId=1):
    if(type(columnsId) == type("A")):
        return columnsId.upper()
    letter = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    letterDict = {}
    one = [x  for x in letter ]
    two = [x + y for x in letter for y in letter]
    three = [x + y + z for x in letter for y in letter for z in letter]
    allLetter = one + two + three
    for columnId,columnName in enumerate(allLetter):
        letterDict[columnId+1] = columnName
    return letterDict[columnsId]

#*******************************************************#
#                       页面设置                        #
#*******************************************************#
#设置页边距(默认单位为厘米)
def setPageMargin(top=2.54,bottom=2.54,left=1.91,right=1.91,header=1.3,footer=1.3):
    #设置页面方向,纵向=1,横向=2(与Word不同)
    ActiveSheet.PageSetup.Orientation = 1
    #上边距(3cm,1cm=28.35pt)
    ActiveSheet.PageSetup.TopMargin = float(top)*28.35
    #下边距
    ActiveSheet.PageSetup.BottomMargin = float(bottom)*28.35
    #左边距
    ActiveSheet.PageSetup.LeftMargin  = float(left)*28.35
    #右边距
    ActiveSheet.PageSetup.RightMargin  = float(right)*28.35
    #页眉
    ActiveSheet.PageSetup.HeaderMargin = float(header)*28.35
    #页脚
    ActiveSheet.PageSetup.FooterMargin = float(footer)*28.35

#设置页眉
def setPageHeader(pos="center",show="pageNum"):
    #pos可选值(center.left,right)
    #show可选值(date,time,fileName,pageNum,bold,italic,underline)
    show = show.replace("date","&D").replace("time","&T")
    show = show.replace("fileName","&F").replace("pageNum","&P")
    show = show.replace("bold","&B").replace("italic","&I").replace("underline","&U")
    show = show.replace(",","")
    if(pos == "center"):
        ActiveSheet.PageSetup.CenterHeader = show
    if(pos == "left"):
        ActiveSheet.PageSetup.LeftHeader = show
    if(pos == "right"):
        ActiveSheet.PageSetup.RightHeader = show    

#设置页脚
def setPageFooter(pos="center",show="pageNum"):
    #pos可选值(center.left,right)
    #show可选值(date,time,fileName,pageNum,bold,italic,underline)
    show = show.replace("date","&D").replace("time","&T")
    show = show.replace("fileName","&F").replace("pageNum","&P")
    show = show.replace("bold","&B").replace("italic","&I").replace("underline","&U")
    show = show.replace(",","")
    if(pos == "center"):
        ActiveSheet.PageSetup.CenterFooter = show
    if(pos == "left"):
        ActiveSheet.PageSetup.LeftFooter = show
    if(pos == "right"):
        ActiveSheet.PageSetup.RightFooter = show    


#*******************************************************#
#                       sheet处理                       #
#*******************************************************#
#增加一个sheet并放到最后面
def addSheet():
    pos = Worksheets(getSheetsCount())
    sheet = Worksheets.Add(Type='xlWorksheet')
    sheet.Move(None,Worksheets(Worksheets.count))

#删除一个sheet
def deleteSheet(SheetIdOrName=None):
    if(type(SheetIdOrName)!=type(None)):
        Worksheets(SheetIdOrName).Delete()

#获得excel中的sheet个数
def getSheetsCount():
    return Worksheets.Count

#获得当前活动的sheet
def getActiveSheet():
    return excel.ActiveSheet

#获得sheet对象通过id(sheet名或者编号)
def getSheetById(SheetIdOrName=None):
    if(type(SheetIdOrName)!=type(None)):
        sheet = Worksheets(SheetIdOrName)
        return sheet
    else:
        return ActiveSheet

#设置sheet的名称
def setSheetName(id,name):
    Worksheets(id).Name = name

#获取sheet的名称
def getSheetName(id):
    return Worksheets(id).Name

#设置当前活动sheet
def setActiveSheet(SheetIdOrName=None):
    if(type(SheetIdOrName)!=type(None)):
        sheet = Worksheets(SheetIdOrName)
        sheet.Activate()


#*******************************************************#
#                       单元格操作                      #
#*******************************************************#
#获取单元格对象
def getCell(row,col=None):
    row = str(row)
    if("$" in row):
        return ActiveSheet.Range(row)
    elif(len(row) == 2 and row[0].isalpha()):
        if(type(col) != type(None) and len(col) == 2 and col[0].isalpha()):
            return ActiveSheet.Range(row+":"+str(col))
        else:
            return ActiveSheet.Range(row)   
    elif(":" in row):
        return ActiveSheet.Range(row)
    elif(len(row) == 1 and type(col) == type(None)):
        if(row.isdigit()):
            return ActiveSheet.Rows(int(row))
        elif(row.isalpha()):
            return ActiveSheet.Columns(row)                 
    else:
        row = int(row)
        return ActiveSheet.Cells(row,col)

#设置单元格的值
def setCellValue(row,col,value=None):
    row = str(row)
    if("$" in row):
        col = str(col)
        ActiveSheet.Range(row).Value = col
    elif(len(row) == 2 and row[0].isalpha()):
        if(type(col) != type(None) and len(col) == 2 and col[0].isalpha()):
            if(type(value) != type(None)):
                ActiveSheet.Range(row+":"+col).Value = value
            else:
                ActiveSheet.Range(row).Value = col  
    elif(":" in row):
         ActiveSheet.Range(row).Value = col
    else:
        row = int(row)
        col = getColumnId(col)
        ActiveSheet.Cells(row,col).Value = value

#获取单元格的值(合并单元格需取左上角坐标)
def getCellValue(row,col=None):
    cellValue = getCell(row,col).Value
    return cellValue

#清除单元格的值
def clearCell(row,col=None):
    getCell(row,col).Clear()

#设置偏移单元格的值(合并单元格需取左上角坐标)
def setOffsetCellValue(row,col,offsetrow=0,offsetcol=0,value=""):
    cell = getCell(row,col)
    cell.Offset(int(offsetrow)+1,int(offsetcol)+1).Value = value

#获取偏移单元格的值
def getOffsetCellValue(row=1,col=1,offsetrow=0,offsetcol=0):
    cell = getCell(row,col)
    return cell.Offset(int(offsetrow)+1,int(offsetcol)+1).Value

#设置所有单元格样式
def setAllCellStyle(fontSize=None,fontName=None,bold=None,italic=None,underline=None,color=None,bgColor=None,borderLeft=None,borderRight=None,borderTop=None,borderBottom=None,numberFormat=None,wrapText=None,horizontalAlign=None,verticalAlign=None):
    cell = ActiveSheet.Cells
    #默认(fontSize=12)
    if(type(fontSize) != type(None)):
        cell.Font.Size = fontSize
    #默认(fontName="宋体")  
    if(type(fontName) != type(None)):   
        cell.Font.Name  = fontName
    #默认(bold=False)     
    if(type(bold) != type(None)):   
        cell.Font.Bold  = bold
    #默认(italic=False)       
    if(type(italic) != type(None)): 
        cell.Font.Italic  = italic
    #默认(underline=False)        
    if(type(underline) != type(None)):
        cell.Font.Underline  = underline
    #单元格字体颜色(ColorIndex等于0为默认色,数字对应当前调色板中颜色的编号) 
    if(type(color) != type(None)):  
        cell.Font.ColorIndex = int(color)
    #单元格背景色
    if(type(bgColor) != type(None)):    
        cell.Interior.ColorIndex = int(bgColor)
    #Borders(1为左,2为右,3为上,4为下),LineStyle(0为无边框,1为实线,2为虚线)    
    if(type(borderLeft) != type(None)): 
        cell.Borders(1).LineStyle = int(borderLeft)
    if(type(borderRight) != type(None)):    
        cell.Borders(2).LineStyle = int(borderRight)
    if(type(borderTop) != type(None)):
        cell.Borders(3).LineStyle = int(borderTop)
    if(type(borderBottom) != type(None)):
        cell.Borders(4).LineStyle = int(borderBottom)
    #其他自定义格式(yyyy"年"m"月"d"日";@)(yyyy-mm-dd)(h:mm:ss;@)(0.00)(#,##0)
    if(type(numberFormat) != type(None)):
    	if(numberFormat == "text"):
        	cell.NumberFormat = "@"   
    	else:
    		cell.NumberFormat = numberFormat
    if(type(horizontalAlign) != type(None)):    
        #文本水平对齐方式(1为两端,2为左(默认),3为中,4为右) 
        cell.HorizontalAlignment = int(horizontalAlign)
    if(type(verticalAlign) != type(None)):  
        #文本垂直对齐方式(1为上,2为中(默认),3为下)  
        cell.VerticalAlignment = int(verticalAlign)
    if(type(wrapText) != type(None)):
        #是否自动换行,默认不换行
        cell.WrapText = wrapText


#设置单元格样式
def setCellStyle(row,col=None,fontSize=None,fontName=None,bold=None,italic=None,underline=None,color=None,bgColor=None,borderLeft=None,borderRight=None,borderTop=None,borderBottom=None,numberFormat=None,wrapText=None,horizontalAlign=None,verticalAlign=None):
    cell = getCell(row,col)
    #默认(fontSize=12)
    if(type(fontSize) != type(None)):
        cell.Font.Size = fontSize
    #默认(fontName="宋体")  
    if(type(fontName) != type(None)):   
        cell.Font.Name  = fontName
    #默认(bold=False)     
    if(type(bold) != type(None)):   
        cell.Font.Bold  = bold
    #默认(italic=False)       
    if(type(italic) != type(None)): 
        cell.Font.Italic  = italic
    #默认(underline=False)        
    if(type(underline) != type(None)):
        cell.Font.Underline  = underline
    #单元格字体颜色(ColorIndex等于0为默认色,数字对应当前调色板中颜色的编号) 
    if(type(color) != type(None)):  
        cell.Font.ColorIndex = int(color)
    #单元格背景色
    if(type(bgColor) != type(None)):    
        cell.Interior.ColorIndex = int(bgColor)
    #Borders(1为左,2为右,3为上,4为下),LineStyle(0为无边框,1为实线,2为虚线)    
    if(type(borderLeft) != type(None)): 
        cell.Borders(1).LineStyle = int(borderLeft)
    if(type(borderRight) != type(None)):    
        cell.Borders(2).LineStyle = int(borderRight)
    if(type(borderTop) != type(None)):
        cell.Borders(3).LineStyle = int(borderTop)
    if(type(borderBottom) != type(None)):
        cell.Borders(4).LineStyle = int(borderBottom)
    #其他自定义格式(yyyy"年"m"月"d"日";@)(yyyy-mm-dd)(h:mm:ss;@)(0.00)(#,##0)    
    if(type(numberFormat) != type(None)):
    	if(numberFormat == "text"):
        	cell.NumberFormat = "@"   
    	else:
    		cell.NumberFormat = numberFormat
    if(type(horizontalAlign) != type(None)):    
        #文本水平对齐方式(1为两端,2为左(默认),3为中,4为右) 
        cell.HorizontalAlignment = int(horizontalAlign)
    if(type(verticalAlign) != type(None)):  
        #文本垂直对齐方式(1为上,2为中(默认),3为下)  
        cell.VerticalAlignment = int(verticalAlign)
    if(type(wrapText) != type(None)):
        #是否自动换行,默认不换行
        cell.WrapText = wrapText

#设置Selection样式(被选择的单元格)
def setSelectionStyle(fontSize=None,fontName=None,bold=None,italic=None,underline=None,color=None,bgColor=None,borderLeft=None,borderRight=None,borderTop=None,borderBottom=None,numberFormat="general",wrapText=None,horizontalAlign=None,verticalAlign=None):
    selection = excelApp.Selection
    #默认(fontSize=12)
    if(type(fontSize) != type(None)):
        selection.Font.Size = fontSize
    #默认(fontName="宋体")  
    if(type(fontName) != type(None)):   
        selection.Font.Name  = fontName
    #默认(bold=False)     
    if(type(bold) != type(None)):   
        selection.Font.Bold  = bold
    #默认(italic=False)       
    if(type(italic) != type(None)): 
        selection.Font.Italic  = italic
    #默认(underline=False)        
    if(type(underline) != type(None)):
        selection.Font.Underline  = underline
    #单元格字体颜色(ColorIndex等于0为默认色,数字对应当前调色板中颜色的编号) 
    if(type(color) != type(None)):  
        selection.Font.ColorIndex = int(color)
    #单元格背景色
    if(type(bgColor) != type(None)):    
        selection.Interior.ColorIndex = int(bgColor)
    #Borders(1为左,2为右,3为上,4为下),LineStyle(0为无边框,1为实线,2为虚线)    
    if(type(borderLeft) != type(None)): 
        selection.Borders(1).LineStyle = int(borderLeft)
    if(type(borderRight) != type(None)):    
        selection.Borders(2).LineStyle = int(borderRight)
    if(type(borderTop) != type(None)):
        selection.Borders(3).LineStyle = int(borderTop)
    if(type(borderBottom) != type(None)):
        selection.Borders(4).LineStyle = int(borderBottom)
    #其他自定义格式(yyyy"年"m"月"d"日";@)(yyyy-mm-dd)(h:mm:ss;@)(0.00)(#,##0)    
    if(numberFormat == "general"):
        selection.NumberFormat = "General"
    elif(numberFormat == "text"):
        selection.NumberFormat = "@"
    else:
        selection.NumberFormat = numberFormat
    if(type(horizontalAlign) != type(None)):    
        #文本水平对齐方式(1为两端,2为左(默认),3为中,4为右) 
        selection.HorizontalAlignment = int(horizontalAlign)
    if(type(verticalAlign) != type(None)):  
        #文本垂直对齐方式(1为上,2为中(默认),3为下)  
        selection.VerticalAlignment = int(verticalAlign)
    if(type(wrapText) != type(None)):
        #是否自动换行,默认不换行
        selection.WrapText = wrapText

#合并单元格(开始坐标,结束坐标,是否强制覆盖数据)
def mergeCell(startPos,endPos=None,isCover=False):
    if(not isCover):
        excelApp.DisplayAlerts = 1
    if(":" in startPos):
        cells = ActiveSheet.Range(startPos)
        cells.Merge()
        #文本水平对齐方式(1为两端,2为左,3为中,4为右)
        cells.HorizontalAlignment = 3
        #文本垂直对齐方式(1为上,2为中,3为下)
        cells.VerticalAlignment = 2
    else:
        cells = ActiveSheet.Range(startPos+":"+endPos)
        cells.Merge()
        cells.HorizontalAlignment = 3
        cells.VerticalAlignment = 2
    excelApp.DisplayAlerts = 0  

#取消合并单元格(开始坐标,结束坐标)
def unMergeCell(startPos,endPos=None):
    if(":" in startPos):
        cells = ActiveSheet.Range(startPos)    
        cells.UnMerge()
    else:
        cells = ActiveSheet.Range(startPos+":"+endPos)    
        cells.UnMerge() 

#剪切粘贴单元格(pastePos为要粘贴的起始点,也就是左上角)
def cutValue(copyStartPos,copyEndPos,pastePos=None):
    if(":" in copyStartPos or type(pastePos) == type(None)):
        ActiveSheet.Range(copyStartPos).Select()
    else:
        ActiveSheet.Range(copyStartPos+":"+copyEndPos).Select()
    if(type(pastePos) == type(None)):
        pastePos = copyEndPos    
    excelApp.Selection.Cut()
    ActiveSheet.Range(pastePos).Select()
    ActiveSheet.Paste()

#复制粘贴单元格(pastePos为要粘贴的起始点,也就是左上角)
def copyValue(copyStartPos,copyEndPos,pastePos=None):
    if(":" in copyStartPos or type(pastePos) == type(None)):
        ActiveSheet.Range(copyStartPos).Select()
    else:
        ActiveSheet.Range(copyStartPos+":"+copyEndPos).Select()
    if(type(pastePos) == type(None)):
        pastePos = copyEndPos    
    excelApp.Selection.Copy()
    ActiveSheet.Range(pastePos).Select()
    ActiveSheet.Paste()
    

#格式刷
def formatCell(copyPos,pasteStartPos,pasteEndPos=None):
    ActiveSheet.Range(copyPos).Select()
    excelApp.Selection.Copy()
    if(type(pasteEndPos) != type(None)):
        ActiveSheet.Range(pasteStartPos+":"+pasteEndPos).Select()
    else:
        ActiveSheet.Range(pasteStartPos).Select()   
    #Paste(-4122为粘贴格式)
    excelApp.Selection.PasteSpecial(Paste=-4122,Operation=0,SkipBlanks=False,Transpose=False)
    excelApp.CutCopyMode = False

#获取最后一个在使用的单元格
def getLeastUsedCell():
    rowsCount = ActiveSheet.UsedRange.Rows.count
    colsCount = ActiveSheet.UsedRange.Columns.count
    leastCell = ActiveSheet.Cells(rowsCount,colsCount)
    return leastCell

#获取最后一个单元格
def getLeastCell():
    rowsCount = ActiveSheet.Rows.count
    colsCount = ActiveSheet.Columns.count
    leastCell = ActiveSheet.Cells(rowsCount,colsCount)
    return leastCell    
	
#*******************************************************#
#                       行列操作                        #
#*******************************************************#
#插入行(内容下移,rowId为插入行,count为插入的行数)
def addRow(rowId,count):
    count = str(int(count)+int(rowId)-1)
    ActiveSheet.Rows(str(rowId)+":"+count).Insert()

#插入列(内容右移,可输入a-zz或者列编号)
def addCol(col,count=1):
    for i in range(count):
        ActiveSheet.Columns(col).Insert()

#删除行
def deleteRow(rowId):
    ActiveSheet.Rows(rowId).Delete()

#删除列
def deleteCol(colId):
    ActiveSheet.Columns(colId).Delete()

#选择行(大于1的整数)
def selectRow(stratRow,endRow):
    ActiveSheet.Rows(str(stratRow)+":"+str(endRow)).Select()

#选择列(大于A的字母)
def selectCol(stratCol,endCol):
    stratCol = getColumnName(stratCol)
    endCol = getColumnName(endCol)
    ActiveSheet.Columns(str(stratCol)+":"+str(endCol)).Select()

#选择已用行(大于1的整数)
def selectUsedRow(rowId):
    colsCount = ActiveSheet.UsedRange.Columns.count
    ActiveSheet.Range(getCell(int(rowId),1),getCell(int(rowId),colsCount)).Select()

#选择已用列(大于A的字母)
def selectUsedCol(colId):
    colId = getColumnId(colId)
    rowsCount = ActiveSheet.UsedRange.Rows.count
    ActiveSheet.Range(getCell(1,int(colId)),getCell(rowsCount,int(colId))).Select()


#合并行
def mergerRow(stratRow,endRow):
    ActiveSheet.Rows(str(stratRow)+":"+str(endRow)).Merge()

#合并列
def mergerCol(stratCol,endCol):
    stratCol = getColumnName(stratCol)
    endCol = getColumnName(endCol)
    ActiveSheet.Columns(str(stratCol)+":"+str(endCol)).Merge()

#设置行高(不指定stratRow则从头设置到已用行,不指定height则自适应) 
def setRowHeight(stratRow=None,endRow=None,height=None):
    leastRow = ActiveSheet.UsedRange.Rows.count
    if(type(stratRow) == type(None)):
        stratRow = 1
        endRow = leastRow
        ActiveSheet.Rows(str(stratRow)+":"+str(endRow)).AutoFit()
    elif(type(stratRow) == type(1)):
        if(type(endRow) != type(None)):
            if(type(endRow) == type(1)):
                if(type(height) != type(None)):
                    ActiveSheet.Rows(str(stratRow)+":"+str(endRow)).RowHeight = float(height)
                else:
                    ActiveSheet.Rows(str(stratRow)+":"+str(endRow)).AutoFit()    
            elif(type(endRow) == type(1.0)):
                height = endRow
                ActiveSheet.Rows(str(stratRow)).RowHeight = float(height)
        else:
            ActiveSheet.Rows(str(stratRow)).AutoFit()    

#设置列宽(不指定stratCol则从头设置到已用列,不指定width则自适应)
def setColHeight(stratCol=None,endCol=None,width=None):
    leastCol = ActiveSheet.UsedRange.Columns.count
    if(type(stratCol) == type(None)):
        stratCol = getColumnName(1)
        endCol = getColumnName(leastCol)
        ActiveSheet.Columns(str(stratCol)+":"+str(endCol)).AutoFit()
    elif(type(stratCol) == type(1)):
        stratCol = getColumnName(stratCol)
        if(type(endCol) != type(None)):
            if(type(endCol) == type(1)):
                endCol = getColumnName(endCol)
                if(type(width) != type(None)):
                    ActiveSheet.Columns(str(stratCol)+":"+str(endCol)).ColumnWidth = float(width)
                else:
                    ActiveSheet.Columns(str(stratCol)+":"+str(endCol)).AutoFit()    
            elif(type(endCol) == type(1.0)):
                width = endCol
                ActiveSheet.Columns(str(stratCol)).ColumnWidth = float(width)
        else:
            ActiveSheet.Columns(str(stratCol)).AutoFit()

#返回行数据(比较慢)
def getRowData(rowId=1):
    colsCount = ActiveSheet.UsedRange.Columns.count
    dataDict = {}
    for i in range(1,colsCount):
        if(type(getCellValue(int(rowId),i))!=type(None)):
            key = ActiveSheet.Cells(int(rowId),i).address
            value = getCellValue(int(rowId),i)
            dataDict[key] = value
    return dataDict

#返回列数据(比较慢)
def getColData(colId=1):
    rowsCount = ActiveSheet.UsedRange.Rows.count
    dataDict = {}
    colId = getColumnId(colId)
    for i in range(1,rowsCount):
        if(type(getCellValue(i,colId)) != type(None)):
            key = ActiveSheet.Cells(i,colId).address
            value = getCellValue(i,colId)
            dataDict[key] = value
    # sorted([(v, k) for k, v in dataDict.items()], reverse=True)       
    return dataDict


#*******************************************************#
#                       图片操作                        #
#*******************************************************#
#插入图片
def insertPic(pos,picAddr):
    try:
        if(os.path.exists(picAddr)):
            cell = ActiveSheet.Range(pos)
            pic = ActiveSheet.Shapes.AddPicture(Filename=picAddr,LinkToFile=1,SaveWithDocument=1,Left=cell.Left+1,Top=cell.Top+1,Width=cell.Width-1,Height=cell.Height-1)
            #1为图片大小和位置随单元格变化，2为大小不变位置变，3为大小和位置均不变(浮动)
            pic.Placement = 1
            #防止出错,二次定位图片
            thisPic = getPic(getPicCount())
            thisPic.Left = cell.Left+1
            thisPic.Top = cell.Top+1
        else:
            print("图片路径有误")
    except:
        print("图片插入失败") 

#剪切图片(isAdaptation是否自适应单元格)
def cutPic(picId,pastePos,isAdaptation = True):
    shape = getPic(int(picId))
    if(type(shape) != type(None)):
        shape.Select()
        excelApp.Selection.Cut()
        getCell(pastePos).Select()
        ActiveSheet.Paste()
        getCell(pastePos).Select()
        leastPicNum = getPicCount()
        pasteCell = getCell(pastePos)
        if(isAdaptation):
            setPic(int(leastPicNum),left=pasteCell.Left+1,top=pasteCell.Top+1,width=pasteCell.width-1,height=pasteCell.Height-1,visible = True)
    else:
        if(showTips):
            print("图片源不存在,无法剪切.(pic编号:"+str(picId)+")")
        
#复制图片(isAdaptation为自适应单元格大小)
def copyPic(picId,pastePos,isAdaptation = True):
    shape = getPic(int(picId))
    if(type(shape) != type(None)):
        shape.Select()
        excelApp.Selection.Copy()
        getCell(pastePos).Select()
        ActiveSheet.Paste()
        getCell(pastePos).Select()
        leastPicNum = getPicCount()
        pasteCell = getCell(pastePos)
        if(isAdaptation):
            setPic(int(leastPicNum),left=pasteCell.Left+1,top=pasteCell.Top+1,width=pasteCell.width-1,height=pasteCell.Height-1,visible = True)
    else:
        if(showTips):
            print("图片源不存在,无法复制.(pic编号:"+str(picId)+")")    
 
#设置图形(isAdaptation为自适应单元格大小)
def setPic(picId,left=None,top=None,width=None,height=None,placement=1,rotation=0,scale=None,visible=True):
    if(type(getPic(int(picId))) != type(None)):
        #取消等比例缩放    
        ActiveSheet.Shapes.Item(int(picId)).LockAspectRatio = 0
        if(type(left) != type(None)):
            ActiveSheet.Shapes.Item(int(picId)).Left = left
        if(type(top) != type(None)):
            ActiveSheet.Shapes.Item(int(picId)).Top = top   
        if(type(width) != type(None)):
            ActiveSheet.Shapes.Item(int(picId)).Width = width
        if(type(height) != type(None)):
            ActiveSheet.Shapes.Item(int(picId)).Height = height
        #跟随单元格变化
        if(type(placement) != type(None)):
            ActiveSheet.Shapes.Item(int(picId)).Placement = placement   
        #旋转度数   
        if(type(rotation) != type(None)):
            ActiveSheet.Shapes.Item(int(picId)).Rotation = rotation
        #等比例缩放图片(第一个参数缩放比率scale为小数,第二个参数为相对原图还是当前大小缩放,第三个参数为缩放轴0为左上,1为中,2为右下)
        if(type(scale) != type(None)):
            ActiveSheet.Shapes.Item(int(picId)).ScaleWidth(scale,False,1) 
            ActiveSheet.Shapes.Item(int(picId)).ScaleHeight(scale,False,1)     
        #是否可见       
        if(type(visible) != type(None)):
            ActiveSheet.Shapes.Item(int(picId)).Visible = visible
    else:
        if(showTips):
                print("图片不存在,无法修改.(pic编号:"+str(picId)+")")          
        
#获取图形
def getPic(picId):
    #Left Top Width Height 
    #TopLeftCell.Address    获取左上角下面的单元格的地址
    #BottomRightCell        右下角单元格
    if(getPicCount()>=int(picId)):
        shape = ActiveSheet.Shapes.Item(int(picId))
        return shape
    return None 

#获取图形所在单元格
def getPicCell(picId):
    shape = getPic(picId)
    if(type(shape) != type(None)):
        picLeft = shape.Left
        picTop = shape.Top
        row,col = None,None
        #行搜索
        for r in range(1,getLeastUsedCell().Row):
            currentCell = getCell(r,1)
            if(currentCell.Top  <= picTop  and (currentCell.Top+currentCell.Height) >= picTop):
                row =currentCell.Row
        #列搜索
        for c in range(1,getLeastUsedCell().Column):
            currentCell = currentCell.Offset(row,c)
            if(currentCell.Left <= picLeft and (currentCell.Left+currentCell.Width) >= picLeft):
                col = currentCell.Column
        return getCell(row,col)
    else:
        if(showTips):
            print("图片不存在(pic编号:"+str(picId)+").")
        return None
    
    
#删除图形
def deletePic(picId):
    if(type(getPic(int(picId))) != type(None)):
        ActiveSheet.Shapes.Item(int(picId)).Delete()

#另存为图片
def saveAsPic(picId,picPos):
    shape = getPic(int(picId))
    if(type(shape) != type(None)):
        shape.Copy()
        chart = ActiveSheet.ChartObjects().Add(shape.Left,shape.Top+3, shape.Width, shape.Height)
        # chart.Activate()
        print(shape.Width,shape.Height,chart.Width,chart.Height)
        chart.Chart.Paste()
        try:
            chart.Chart.Export(picPos)
            chart.Delete()
        except:
            print("图片保存失败,请检查路径是否正确,应改为/")
    else:
        if(showTips):
            print("图片保存失败(pic编号:"+str(picId)+").")
        return None

#获取图形数量
def getPicCount():
    count = ActiveSheet.Shapes.Count
    return count

#*******************************************************#
#                       公式运算                        #
#*******************************************************#
#求和(开始位置，结束位置，赋值位置)
def sum(startPos,endPos,assignPos):
    total = excelApp.WorksheetFunction.Sum(ActiveSheet.Range(startPos+":"+endPos))
    address = ActiveSheet.Range(assignPos).Address
    setCellValue(address,total)
    return total

#公式字母转大写
def upper(pos):
    address = ActiveSheet.Range(pos).Address
    cellValue = getCellValue(address)
    ActiveSheet.Range(pos).FormulaR1C1 = '=UPPER("'+cellValue+'")'
    ActiveSheet.Range(pos).Select()
    return getCellValue(address)

#公式字母转小写
def lower(pos):
    address = ActiveSheet.Range(pos).Address
    cellValue = getCellValue(address)
    ActiveSheet.Range(pos).FormulaR1C1 = '=LOWER("'+cellValue+'")'
    ActiveSheet.Range(pos).Select()
    return getCellValue(address)

#*******************************************************#
#                       其他操作                        #
#*******************************************************#
#判断excel版本
def getExcelVersion():
	version = int(float(excelApp.Version))
	if(version == 8):
		return "Excel 97"
	elif(version == 9):
		return "Excel 2000"
	elif(version == 10):
		return "Excel 2002"	
	elif(version == 11):
		return "Excel 2003"
	elif(version == 12):
		return "Excel 2007"
	elif(version == 14):
		return "Excel 2010"
	elif(version == 15):
		return "Excel 2013"
	else:
		return "Excel 未知版本"		

#全文查找字符串,目前只能模糊查找,无法精确.(LookAt属性不起作用)
def findContent(value):
    #What为要查找的内容,LookAt分全部查找和匹配查找(xlWhole{1},xlPart{2}),SearchOrder为按行查找还是列查找,SearchDirection为查找方向
    startCell = getLeastUsedCell()
    findObj = ActiveSheet.Cells
    #要查询的字符串
    what = str(value)
    #查询开始位置
    after = startCell
    lookIn = -4163
    #1是全部匹配，2是部分匹配
    lookAt = 1
    #按行查找为1，按列查找为2
    searchOrder = 1
    #查找方向(1为向后，2为向前)
    searchDirection = 1
    #是否匹配大小写
    matchCase = True
    matchByte = False
    searchFormat = True
    text = findObj.Find(what,after,lookIn,lookAt,searchOrder,searchDirection,matchCase,matchByte,searchFormat)
    if(type(text)!=type(None)):
        ActiveSheet.Range(text.address).Select()
        return (text.address,text.value,findObj,text)
    else:
        return None 

#块中查找字符串
def findContentInRange(startPos,endPos,value):
    currentRange = ActiveSheet.Range(startPos+":"+endPos)
    rowsCount = currentRange.Rows.count
    colsCount = currentRange.Columns.count
    cell = currentRange.Offset(rowsCount,colsCount)
    findObj = ActiveSheet.Range(startPos+":"+endPos)
    #要查询的字符串
    what = str(value)
    #查询开始位置
    after = cell
    lookIn = -4163
    #1是全部匹配，2是部分匹配
    lookAt = 1
    #1为按行查找，2为按列查找
    searchOrder = 1
    #查找方向(1为向后，2为向前)
    searchDirection = 1
    #是否匹配大小写
    matchCase = True
    matchByte = False
    searchFormat = True
    text = findObj.Find(what,after,lookIn,lookAt,searchOrder,searchDirection,matchCase,matchByte,searchFormat)
    if(type(text)!=type(None)):
        ActiveSheet.Range(text.address).Select()
        return (text.address,text.value,findObj,text)
    else:
        return None 

#行中查找字符串
def findContentInRow(rowNum,value):
    colCount = ActiveSheet.Columns.count
    cell = ActiveSheet.Cells(rowNum,colCount)
    findObj = ActiveSheet.Rows(rowNum)
    #要查询的字符串
    what = str(value)
    #查询开始位置
    after = cell
    lookIn = -4163
    #1是全部匹配，2是部分匹配
    lookAt = 1
    #按行查找为1，按列查找为2
    searchOrder = 1
    #查找方向(1为向后，2为向前)
    searchDirection = 1
    #是否匹配大小写
    matchCase = True
    matchByte = False
    searchFormat = True
    text = findObj.Find(what,after,lookIn,lookAt,searchOrder,searchDirection,matchCase,matchByte,searchFormat)
    if(type(text)!=type(None)):
        ActiveSheet.Range(text.address).Select()
        return (text.address,text.value,findObj,text)
    else:
        return None 

#列中查找字符串
def findContentInColumn(colNum,value):
    rowCount = ActiveSheet.Rows.count
    cell = ActiveSheet.Cells(rowCount,colNum)
    findObj = ActiveSheet.Columns(colNum)
    #要查询的字符串
    what = str(value)
    #查询开始位置
    after = cell
    lookIn = -4163
    #1是全部匹配，2是部分匹配
    lookAt = 1
    #按行查找为1，按列查找为2
    searchOrder = 1
    #查找方向(1为向后，2为向前)
    searchDirection = 1
    #是否匹配大小写
    matchCase = True
    matchByte = False
    searchFormat = True
    text = findObj.Find(what,after,lookIn,lookAt,searchOrder,searchDirection,matchCase,matchByte,searchFormat)
    if(type(text)!=type(None)):
        ActiveSheet.Range(text.address).Select()
        return (text.address,text.value,findObj,text)
    else:
        return None 

#查找下一个
def findNextContent(findObject):
    if(type(findObject)!=type(None)):
        findObj = findObject[2]
        value = findObject[3]
        text = findObj.FindNext(value)
        if(type(text)!=type(None)):
            ActiveSheet.Range(text.address).Select()
            return (text.address,text.value,findObj,text)
        else:
            return None 
    else:
        return None 
    

#查找替换全部
def findAndReplace(oldValue,newValue):
    #What为要替换的内容,Replacement为替换值,LookAt分全部查找和匹配查找(xlWhole,xlPart)
    text = ActiveSheet.Cells.Replace(What=str(oldValue),Replacement=str(newValue),LookAt=1,SearchOrder=1,MatchCase=True)
    if(type(text)!=type(None)):
        return text
    else:
        return False    
    return False

#获取搜索值的数量
def getFindValueCount():
    pass
    
#打印预览
def printReview():
    ActiveSheet.PrintPreview()

#打印当前sheet
#From为开始页码,To为结束页码,Copies为打印份数,Preview为是否打印前预览,ActivePrinter为设置活动打印机的名称
def printSheet(startPage=None,endPage=None,copies=1):
    ActiveSheet.PrintOut(From=startPage,To=endPage,Copies=copies,Preview=False,ActivePrinter=None)

#关闭excel
def closeExcel():
    #关闭前保存活动工作表所做的更改
    if(type(ActiveWorkbook)!=type(None)):
        ActiveWorkbook.Save()
    #关闭前显示提示
    excelApp.DisplayAlerts = 1
    if(type(excel)!=type(None)):
        excel.Close()
    

#退出excel应用
def quitExcel():
    if(type(excelApp)!=type(None)):
        excelApp.Quit()
        print("excel程序执行完毕...")

