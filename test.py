#导入pyExcel模块
from  pyExcel  import *

setExcelVisible(True,False)
newExcel("D:/test.xls")
openExcel("D:/test.xls")
setCellValue(1,1,"pyExcel")
setColHeight()
setRowHeight()
setAllCellStyle(bold=True)
mergeCell("a1:c4",isCover=True)
closeExcel()
quitExcel()

