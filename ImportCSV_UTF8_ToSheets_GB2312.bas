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
    
    ' ��ȡ��ʹ������
    Set usedRange = ws.UsedRange
    If usedRange Is Nothing Then Exit Sub
    
    ' �������е�Ԫ�񣬽���ֵ�滻Ϊ"-"
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
    
    ' ��鹤�����Ƿ�������
    If ws.UsedRange Is Nothing Then
        Exit Sub
    End If
    
    ' ��ȡ��ʹ������
    Set usedRange = ws.UsedRange
    If usedRange Is Nothing Then
        Exit Sub
    End If
    
    lastRow = usedRange.Rows.Count
    lastCol = usedRange.Columns.Count
    
    If lastRow = 0 Or lastCol = 0 Then
        Exit Sub
    End If
    
    ' ���ñ��Χ
    Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' ������б߿�
    tableRange.Borders.LineStyle = xlNone
    
    ' �������߱��ʽ
    With tableRange
        ' �ϱ߿�
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = 1.5
        
        ' ��ͷ�±߿�
        .Range(.Cells(1, 1), .Cells(1, lastCol)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(.Cells(1, 1), .Cells(1, lastCol)).Borders(xlEdgeBottom).Weight = 1
        
        ' �±߿�
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = 1.5
    End With
    
    ' ��������
    With tableRange
        .Font.Name = "����"
        .Font.Size = 10.5  ' 5�������Ӧ10.5��
    End With
    
    ' ���ñ�ͷ��ʽ
    With tableRange.Rows(1)
        .Font.Name = "����"
        .Font.Size = 10.5
        .HorizontalAlignment = xlCenter
    End With
    
    ' �����п�����Ӧ����
    tableRange.Columns.AutoFit
    
    ' ���ñ��λ�ã����У�
    tableRange.HorizontalAlignment = xlCenter
    tableRange.VerticalAlignment = xlCenter
    
    ' ��ӱ�ţ�ʾ������4-2��
    Dim tableNumber As String
    tableNumber = "��" & ws.Index & "-" & ws.Index
    
    ' �ڱ���Ϸ�����һ�����ڱ��
    ws.Rows(1).Insert
    ws.Range("A1").Value = tableNumber
    ws.Range("A1").Font.Name = "����"
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

    folderPath = "D:\����\��ҵ���\assets\static"  ' �޸�Ϊ���·��
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    fileName = Dir(folderPath & "*.csv")
    Do While fileName <> ""
        originalName = Left(fileName, InStrRev(fileName, ".") - 1)
        sheetName = SanitizeSheetName(originalName)
        
        ' ȷ����������Ψһ
        suffix = 1
        Do While Not SheetNameIsUnique(sheetName)
            sheetName = Left(SanitizeSheetName(originalName), 28) & "_" & suffix
            suffix = suffix + 1
        Loop

        Set ws = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
        If ws Is Nothing Then
            MsgBox "�޷���������������Excel�Ƿ��������С�", vbExclamation
            Exit Sub
        End If
        
        ws.name = sheetName
        ws.Cells.Clear

        Set qt = ws.QueryTables.Add(Connection:="TEXT;" & folderPath & fileName, Destination:=ws.Range("A1"))
        If qt Is Nothing Then
            MsgBox "�޷�������ѯ������CSV�ļ���ʽ�Ƿ���ȷ��", vbExclamation
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
        
        ' �滻�յ�Ԫ��Ϊ"-"
        ReplaceEmptyCells ws
        
        ' ���ø�ʽ�������ӳ���
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