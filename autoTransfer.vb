'@author: Adonis Settouf
'@mail: adonis.settouf@gmail.com

Sub writeDatas()
    Dim val As Variant, row As Long, range As String
    Dim wb As Workbook
    Dim KAList
    KAList = Array("France", "UK", "Germany", "Spain", "Portugal", "Belgium", "Netherlands", "Norway")
    MsgBox (KAList(0))
    'Workbooks.Open ("C:\Users\asettouf\Documents\Projecto\Gab-Macro\phone_report_Full_May.xls")
    Set wb = findWorkbook("phone_report")
    val = wb.Worksheets("PSSD combined").range("B4:G22").Value
    'ThisWorkbook.Worksheets(val(1, 1)).Range("D127").Value = val(1, UBound(val, 2))
    row = findRangeToWrite()
    range = "D" & CStr(row)
    'MsgBox (range)
    For i = 1 To (UBound(val, 1))
      'MsgBox (ThisWorkbook.Worksheets(val(1, 1)).Name)
      'MsgBox (val(i, 1))
      If (InStr(val(i, 1), "Ireland") > 0) Then
        ThisWorkbook.Worksheets("UK").range(range).Value = val(i, UBound(val, 2))
      ElseIf (InStr(val(i, 1), "Africa") > 0) Then
        ThisWorkbook.Worksheets("South Africa").range(range).Value = val(i, UBound(val, 2))
      ElseIf (InStr(val(i, 1), "LexLIME") > 0) Then
        
      Else
        ThisWorkbook.Worksheets(val(i, 1)).range(range).Value = val(i, UBound(val, 2))
      End If
    Next
End Sub

'Find the good row to write datas
Function findRangeToWrite() As Long
    Dim lastRow As Long, firstRow As Long, counter As Long
    Dim val As Variant
    counter = 0
    lastRow = ThisWorkbook.Worksheets("France").Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).row
    firstRow = lastRow - 23
    val = ThisWorkbook.Worksheets("France").range("A" & CStr(firstRow) & ":B" & CStr(lastRow)).Value
    'MsgBox (InStr(val(1, 1), CStr(Year(Date))))
    For i = 1 To UBound(val, 1)
        If (InStr(val(i, 1), CStr(Year(Date))) > 0 And InStr(MonthName(Month(Date) - 1), val(i, 2)) > 0) Then
            counter = i
            Exit For
        End If
    Next
    findRangeToWrite = firstRow + counter - 1
End Function

'find Workbooks with names (partial name ok)
Function findWorkbook(nameWB As String) As Workbook
    Dim buffWb As Workbook
    For Each book In Workbooks
        If (InStr(book.name, nameWB) > 0) Then
            Set findWorkbook = book
            Exit For
        End If
    Next book
End Function


'Proc for testing new func
Sub test()
   Dim wb As Workbook
   Set wb = findWorkbook("phone_report")
   MsgBox (wb.name)
End Sub

Private Sub TransferScript_Click()
    writeDatas
End Sub


