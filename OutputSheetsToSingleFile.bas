Sub OutputSheetsToSingleFile()
'
' 导出sheet到单个文件 宏
'

Application.ScreenUpdating = False
For Each sht In Sheets
sht.Copy
  ActiveWorkbook.SaveAs "你的磁盘位置" & sht.Name & ".xlsx"
ActiveWorkbook.Close
Next
Application.ScreenUpdating = True
'
End Sub
