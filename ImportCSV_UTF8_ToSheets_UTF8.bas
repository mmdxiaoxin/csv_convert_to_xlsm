Function SanitizeSheetName(name As String) As String
    Dim invalidChars As Variant, ch As Variant
    invalidChars = Array(":", "\", "/", "?", "*", "[", "]")
    For Each ch In invalidChars
        name = Replace(name, ch, "_")
    Next
    If Len(name) > 31 Then name = Left(name, 31)
    SanitizeSheetName = name
End Function

Sub ReplaceEmptyCells(ws As Worksheet)
    Dim usedRange As Range
    Dim cell As Range
    
    ' 获取已使用区域
    Set usedRange = ws.UsedRange
    If usedRange Is Nothing Then Exit Sub
    
    ' 遍历所有单元格，将空值替换为"-"
    For Each cell In usedRange
        If IsEmpty(cell) Or Trim(cell.Value) = "" Then
            cell.Value = "-"
        End If
    Next cell
End Sub

Sub FormatTable(ws As Worksheet)
    On Error Resume Next
    
    Dim usedRange As Range
    Dim tableRange As Range
    Dim lastRow As Long, lastCol As Long
    
    ' 检查工作表是否有数据
    If ws.UsedRange Is Nothing Then
        Exit Sub
    End If
    
    ' 获取已使用区域
    Set usedRange = ws.UsedRange
    If usedRange Is Nothing Then
        Exit Sub
    End If
    
    lastRow = usedRange.Rows.Count
    lastCol = usedRange.Columns.Count
    
    If lastRow = 0 Or lastCol = 0 Then
        Exit Sub
    End If
    
    ' 设置表格范围
    Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' 清除所有边框
    tableRange.Borders.LineStyle = xlNone
    
    ' 设置三线表格式
    With tableRange
        ' 上边框
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThick
        
        ' 表头下边框（使用细线）
        .Range(.Cells(1, 1), .Cells(1, lastCol)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(.Cells(1, 1), .Cells(1, lastCol)).Borders(xlEdgeBottom).Weight = xlMedium
        
        ' 下边框
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
    End With
    
    ' 设置字体
    With tableRange
        .Font.Name = "宋体"
        .Font.Size = 10.5  ' 5号字体对应10.5磅
    End With
    
    ' 设置表头格式
    With tableRange.Rows(1)
        .Font.Name = "黑体"
        .Font.Size = 10.5
        .HorizontalAlignment = xlCenter
    End With
    
    ' 调整列宽以适应内容
    tableRange.Columns.AutoFit
    
    ' 设置表格位置（居中）
    tableRange.HorizontalAlignment = xlCenter
    tableRange.VerticalAlignment = xlCenter
    
    ' 添加表号（示例：表4-2）
    Dim tableNumber As String
    tableNumber = "表" & ws.Index & "-" & ws.Index
    
    ' 在表格上方插入一行用于表号
    ws.Rows(1).Insert
    ws.Range("A1").Value = tableNumber
    ws.Range("A1").Font.Name = "黑体"
    ws.Range("A1").Font.Size = 10.5
    ws.Range("A1").HorizontalAlignment = xlCenter
    
    On Error GoTo 0
End Sub

Sub ImportCSV_UTF8_ToSheets()
    On Error Resume Next
    
    Dim folderPath As String, fileName As String
    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim sheetName As String
    Dim originalName As String
    Dim suffix As Integer

    folderPath = "D:\桌面\毕业设计\assets\static"  ' 修改为你的路径
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    fileName = Dir(folderPath & "*.csv")
    Do While fileName <> ""
        originalName = Left(fileName, InStrRev(fileName, ".") - 1)
        sheetName = SanitizeSheetName(originalName)
        
        ' 确保工作表名唯一
        suffix = 1
        Do While Not SheetNameIsUnique(sheetName)
            sheetName = Left(SanitizeSheetName(originalName), 28) & "_" & suffix
            suffix = suffix + 1
        Loop

        Set ws = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
        If ws Is Nothing Then
            MsgBox "无法创建工作表，请检查Excel是否正常运行。", vbExclamation
            Exit Sub
        End If
        
        ws.name = sheetName
        ws.Cells.Clear

        Set qt = ws.QueryTables.Add(Connection:="TEXT;" & folderPath & fileName, Destination:=ws.Range("A1"))
        If qt Is Nothing Then
            MsgBox "无法创建查询表，请检查CSV文件格式是否正确。", vbExclamation
            Exit Sub
        End If
        
        With qt
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileCommaDelimiter = True
            .TextFilePlatform = 65001
            .TextFileParseType = xlDelimited
            .TextFileColumnDataTypes = Array(1)
            .Refresh BackgroundQuery:=False
        End With
        
        ' 替换空单元格为"-"
        ReplaceEmptyCells ws
        
        ' 调用格式化表格的子程序
        FormatTable ws

        fileName = Dir
    Loop
    
    On Error GoTo 0
End Sub

Function SheetNameIsUnique(name As String) As Boolean
    Dim ws As Worksheet
    SheetNameIsUnique = True
    For Each ws In ThisWorkbook.Sheets
        If ws.name = name Then
            SheetNameIsUnique = False
            Exit Function
        End If
    Next
End Function 