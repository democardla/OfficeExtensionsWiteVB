Sub 导出sheet到单个文件()
'
' 导出sheet到单个文件 宏
'

Application.ScreenUpdating = False
For Each sht In Sheets
sht.Copy
ActiveWorkbook.SaveAs "/Users/democardla/Desktop/excel/" & sht.Name
ActiveWorkbook.Close
Next
Application.ScreenUpdating = True
'
End Sub
