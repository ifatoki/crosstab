Attribute VB_Name = "Module1"
Option Explicit

Dim myWorkbook As Workbook
Dim sourceSheet As Worksheet
Dim controlSheet As Worksheet
Dim totalRows() As Range
Dim lastRow As Integer
Dim initialRow As Integer

Private Sub initialize()
    Set myWorkbook = Application.ThisWorkbook
    Set sourceSheet = myWorkbook.Sheets("Original weighted")
    sourceSheet.Copy After:=sourceSheet
    Set controlSheet = myWorkbook.Sheets("Original weighted (2)")
    controlSheet.Name = "Control"
    With controlSheet
        initialRow = .Range("A1").End(xlDown).End(xlDown).Offset(1, 0).Row
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    totalRows = getTotals()
    
    
'    room_data = Range(Range("aboveTable").Offset(1, 0), Range("total").Offset(-1, 0)).Value
'    room_count = UBound(room_data) / 3
'    description_count = Range(Range("descriptionField"), Range("measurementField").Offset(-1)).Rows.Count
'    Application.DisplayAlerts = False
'    If action = GENERATE_REPORT Then
'        stat_upper = Range("static_upper").Value
'        stat_lower = Range("static_lower").Value
'        Application.StatusBar = "Generating Report in Microsoft Word. Please wait..."
'    End If
End Sub

Private Sub finalize()
    sourceSheet.Protect
    Application.StatusBar = ""
    Set myWorkbook = Nothing
    Set sourceSheet = Nothing
    Set sourceSheet = Nothing
    Set stat_upper = Nothing
    Set stat_lower = Nothing
    Set room_data = Nothing
    room_count = 0
    description_count = 0
    Sheet1.Activate
    Application.ScreenUpdating = True
End Sub

Private Function getTotals()
    Dim currentTotal As Range
    Dim currentIndex As Integer
    Dim totals() As Range
    
    currentIndex = -1
    ReDim totals(2000)
    With controlSheet
        Set currentTotal = .Range("B" & lastRow)
        While currentTotal.Row >= initialRow
            If (IsNumeric(currentTotal.Value) And Trim(currentTotal.Offset(-1, 0).Value) <> "") Then
                currentIndex = currentIndex + 1
                Set totals(currentIndex) = currentTotal
            End If
            Set currentTotal = currentTotal.End(xlUp)
        Wend
    End With
    ReDim Preserve totals(currentIndex)
    getTotals = totals
End Function

Sub main()
    initialize
End Sub
