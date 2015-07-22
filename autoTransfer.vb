'@author: Adonis Settouf
'@mail: adonis.settouf@gmail.com

Sub writeDatas()
    Dim val As Variant, KAVal As Variant, row As Long, Range As String, KARange As String, BAUVal As Variant, KACountry As Boolean
    Dim wb As Workbook
    Dim KAList As Variant
    Application.Calculation = xlCalculationManual
    KAList = Array("France", "UK", "Germany", "Spain", "Portugal", "Belgium", "Netherlands", "Norway")
    Set wb = findWorkbook("phone_report")
    val = wb.Worksheets("PSSD combined").Range("B4:G22").value
    
    row = findRangeToWrite()
    Range = "D" & CStr(row)
    KARange = "P" & CStr(row)
    KAVal = wb.Worksheets("PSSD KA").Range("B4:G22").value
    BAUVal = wb.Worksheets("PSSD BAU").Range("B4:G22").value

    For i = 1 To (UBound(KAVal, 1))
         For Each country In KAList
            If (InStr(KAVal(i, 1), country)) > 0 Then
                Call writeToSheet(Range, BAUVal(i, UBound(BAUVal, 2)), KAVal(i, 1))
                Call writeToSheet(KARange, KAVal(i, UBound(KAVal, 2)), country)
                KACountry = True
                Exit For
            End If
        KACountry = False
        Next country
        If Not KACountry Then
            Call writeToSheet(Range, val(i, UBound(val, 2)), val(i, 1))
        End If
    Next
    Application.Calculation = xlCalculationAutomatic
End Sub

'writeToTheGlobalForecast Workbook
Function writeToSheet(ByVal Range As String, ByVal value As Integer, ByVal country As String)
    If (InStr(country, "Ireland") > 0) Then
      ThisWorkbook.Worksheets("UK").Range(Range).value = value
    ElseIf (InStr(country, "Africa") > 0) Then
      ThisWorkbook.Worksheets("South Africa").Range(Range).value = value
    ElseIf (InStr(country, "LexLIME") > 0) Then
    Else
      ThisWorkbook.Worksheets(country).Range(Range).value = value
    End If
End Function
'Find the good row to write datas
Function findRangeToWrite() As Long
    Dim lastRow As Long, firstRow As Long, counter As Long
    Dim val As Variant
    counter = 0
    lastRow = ThisWorkbook.Worksheets("France").Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).row
    firstRow = lastRow - 23
    val = ThisWorkbook.Worksheets("France").Range("A" & CStr(firstRow) & ":B" & CStr(lastRow)).value
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
Function findWorkbook(ByVal nameWB As String) As Workbook
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




Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
