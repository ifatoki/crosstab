Attribute VB_Name = "Module1"
Option Explicit

Public Enum outputTypes
    TOP2_BOTTOM2 = 2
    TOP3_BOTTOM3 = 3
    PERCENT_SORTED = 1
End Enum

Private sourceWorkbook As Workbook
Private sourceSheet As Worksheet
Private controlSheet As Worksheet
Private totalRows() As Range
Private totalRowsInt() As Integer
Private batchCols() As Range
Private lastRow As Integer
Private lastCol As Integer
Private firstHeader As Range
Private initialRow As Integer
Private fileType As Integer
Private initialized As Boolean
Private Enum FileTypes
    Default = 0
    Mean = 1
    Index = 2
End Enum

Private Sub initialize()
    Dim dialog As FileDialog
    Dim dialogResult As Long
    Const ERROR_NO_SOURCE As Long = vbObjectError + 513
    
    Debug.Print ERROR_NO_SOURCE
    On Error GoTo errHandler
    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
    With dialog
        .Title = "Select the source file"
        .Filters.Clear
        .Filters.Add "Spreadsheets", "*.xlsx; *.xls", 1
        dialogResult = .Show
        If dialogResult <> 0 Then
            DoEvents
            Set sourceWorkbook = Workbooks.Open(.SelectedItems(1))
        End If
        If sourceWorkbook Is Nothing Then Err.Raise ERROR_NO_SOURCE
        Set sourceSheet = sourceWorkbook.Sheets(1)
        sourceSheet.Copy After:=sourceSheet
        Set controlSheet = sourceWorkbook.Sheets(2)
        controlSheet.Name = "Control"
    End With
    With controlSheet.UsedRange
        Set firstHeader = .Range("B1").End(xlDown)
        initialRow = .Range("A1").End(xlDown).End(xlDown).Offset(1, 0).Row
        lastRow = .Rows.Count
        lastCol = .Columns.Count
    End With
    fileType = getFileType()
    totalRows = getTotals()
    totalRowsInt = getTotalRows(totalRows)
    If fileType = FileTypes.Mean Then insertMeanRows
    lastRow = controlSheet.UsedRange.Rows.Count
    batchCols = getbatches()
    initialized = True
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Exit Sub
errHandler:
    Debug.Print Err.Number
    If Err.Number = ERROR_NO_SOURCE Then
        MsgBox "Script stopped. No file selected.", vbOKOnly + vbExclamation, "Cancelled"
    Else
        MsgBox "Script stopped. Please check the selected file and ensure its compliant to the formatting requirements.", vbOKOnly + vbExclamation, "Invalid File"
        sourceWorkbook.Close False
    End If
End Sub

Private Sub finalize()
    If Not controlSheet Is Nothing And Not controlSheet.Name Like "Output*" Then controlSheet.Name = "Output"
    Set sourceWorkbook = Nothing
    Set sourceSheet = Nothing
    ReDim totalRows(1)
    ReDim batchCols(1)
    Set firstHeader = Nothing
    lastCol = 0
    lastRow = 0
    initialRow = 0
    fileType = FileTypes.Default
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Private Function getFileType()
    Dim cellValue As String
    With controlSheet
        cellValue = .Cells(initialRow, .UsedRange.Columns.Count).Offset(-2, 0).Value
        cellValue = LCase(Trim(cellValue))
        If cellValue = "mean" Then
            getFileType = FileTypes.Mean
        ElseIf cellValue = "index on column %" Then
            getFileType = FileTypes.Index
        Else
            getFileType = FileTypes.Default
        End If
    End With
End Function

Private Function getTotals()
    Dim currentTotal As Range
    Dim currentIndex As Integer
    Dim totals() As Range
    
    currentIndex = -1
    ReDim totals(2000)
    With controlSheet
        Set currentTotal = .Range("B" & lastRow)
        While currentTotal.Row >= initialRow
            If (IsNumeric(currentTotal.Value) And Trim(currentTotal.Offset(0, -1).Value) = "Total") Then
                currentIndex = currentIndex + 1
                Set totals(currentIndex) = currentTotal
            End If
            Set currentTotal = currentTotal.End(xlUp)
        Wend
    End With
    ReDim Preserve totals(currentIndex)
    getTotals = totals
End Function

Private Function getTotalRows(totalRows)
    Dim rows() As Integer
    Dim index As Integer
    ReDim rows(UBound(totalRows))
    
    index = 0
    While index <= UBound(totalRows)
        rows(index) = totalRows(index).Row
        index = index + 1
    Wend
    getTotalRows = rows
End Function

Private Function getbatches()
    Dim currentBatch As Range
    Dim currentIndex As Integer
    Dim batches() As Range
    
    currentIndex = -1
    ReDim batches(2000)
    With controlSheet
        Set currentBatch = .Cells(firstHeader.Row, .Columns.Count).End(xlToLeft)
        While currentBatch.Column >= 2
            currentIndex = currentIndex + 1
            Set batches(currentIndex) = currentBatch
            Set currentBatch = currentBatch.End(xlToLeft)
        Wend
    End With
    ReDim Preserve batches(currentIndex)
    getbatches = batches
End Function

Private Sub processBatch(ByVal batch As Range, Optional isPercentSort As Boolean = False, Optional includeCounts As Boolean = False)
    Dim lastBatchCol As Integer
    Dim headerRange As Range
    
    With controlSheet
        lastBatchCol = .Cells(initialRow - 2, batch.Column).End(xlToRight).Column
        Set headerRange = .Range(batch, batch.Offset(initialRow - 2 - batch.Row))
        headerRange.UnMerge
        Set headerRange = .Range(headerRange, headerRange.Offset(0, lastBatchCol - batch.Column))
        If Not includeCounts Then headerRange.Copy Destination:=batch.Offset(, 1)
        If fileType = FileTypes.index Then
            Dim headerSubRange As Range
            Dim nextHeader As Range
            
            Set nextHeader = batch.Offset(1, 1).End(xlToRight)
            Set headerSubRange = .Range(nextHeader, .Cells(nextHeader.Row, lastBatchCol))
            headerSubRange.Copy Destination:=nextHeader.Offset(, 1)
        End If
        With .Range(.Cells(firstHeader.Row, batch.Column), .Cells(lastRow, lastBatchCol))
            .Font.Name = "Gotham Book"
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
        End With
    End With
    processTotals batch, lastBatchCol, includeCounts
    deleteColumns batch, lastBatchCol, isPercentSort, includeCounts
End Sub

Private Sub deleteColumns(ByVal batch As Range, lastBatchCol As Integer, Optional isPercentSort As Boolean = False, Optional includeCounts As Boolean = False)
    Dim currentCol As Integer
    Dim firstCol As Integer
    Dim offsetIncrement As Integer
    Dim firstRow As Integer
    
    offsetIncrement = 2
    If fileType > FileTypes.Default Then offsetIncrement = 3
    firstCol = batch.Column
    With controlSheet
        currentCol = lastBatchCol - (offsetIncrement - 1)
        .Columns(currentCol + offsetIncrement).Clear
        While currentCol >= firstCol
            If Not includeCounts Then
                .Columns(currentCol).Delete
                If fileType = FileTypes.index And currentCol > firstCol Then
                    .Columns(currentCol).Delete
                ElseIf fileType = FileTypes.Mean Or (fileType = FileTypes.index And currentCol = firstCol) Then
                    .Columns(currentCol + 1).Delete
                End If
            End If
            If isPercentSort Then
                Dim count As Integer
                
                count = LBound(totalRows)
                If currentCol > 2 Then .Columns(currentCol).Insert Shift:=xlToRight
                While count <= UBound(totalRows)
                    Dim sortKeyRange As Range
                    Dim sortRange As Range
                    
                    firstRow = initialRow
                    If count < UBound(totalRows) Then firstRow = totalRowsInt(count + 1) + 2
                    Set sortKeyRange = .Range(.Cells(firstRow, currentCol + 1), .Cells(totalRowsInt(count) - 1, currentCol + 1))
                    sortKeyRange.ColumnWidth = 9.98
                    If currentCol <= 2 Then
                        Set sortKeyRange = .Range(.Cells(firstRow, currentCol), .Cells(totalRowsInt(count) - 1, currentCol))
                        sortKeyRange.ColumnWidth = 9.98
                    Else
                        sortKeyRange.Offset(, -1).ColumnWidth = 9.98
                        .Range("A" & firstRow & ":A" & totalRowsInt(count) - 1).Copy .Cells(firstRow, currentCol)
                    End If
                    Set sortRange = Range(sortKeyRange.Offset(, -1), sortKeyRange)
                    If includeCounts Then
                        Set sortKeyRange = sortKeyRange.Offset(, 1)
                        Set sortRange = Range(sortKeyRange.Offset(, -2), sortKeyRange)
                    End If
                    .Sort.SortFields.Clear
                    .Sort.SortFields.Add Key:=sortKeyRange, SortOn:=xlSortOnValues, Order:= _
                        xlDescending, DataOption:=xlSortNormal
                    With .Sort
                        .SetRange sortRange
                        .MatchCase = False
                        .Orientation = xlTopToBottom
                        .SortMethod = xlPinYin
                        .Apply
                    End With
                    count = count + 1
                Wend
            End If
            currentCol = currentCol - offsetIncrement
        Wend
    End With
End Sub

Private Sub processTotals(ByVal batch As Range, lastBatchCol As Integer, Optional includeCounts As Boolean = False)
    Dim total As Variant
    Dim batchOffset As Integer
    Dim currentTotal As Range
    Dim currentOffset As Integer
    Dim offsetIncrement As Integer
    Dim conditionalOffset As Integer
    
    offsetIncrement = 2
    If fileType > FileTypes.Default Then offsetIncrement = 3
    batchOffset = batch.Column - totalRows(0).Column
    currentOffset = 0
    While currentOffset + batch.Column <= lastBatchCol
        For Each total In totalRows
            Set currentTotal = total.Offset(, batchOffset).Offset(, currentOffset)
            conditionalOffset = 1
            If includeCounts Then
                currentTotal.NumberFormat = "0.0,,""M"""
                GoTo continue
            End If
            If fileType = FileTypes.index And currentOffset <> 0 Then conditionalOffset = 2
            With currentTotal.Offset(, conditionalOffset)
                If IsNumeric(currentTotal.Value) Then
                    .Value = Round(currentTotal.Value / 1000000, 6)
                Else
                    .Value = 0
                End If
                .NumberFormat = "0.0"
            End With
            If fileType = FileTypes.Mean Then
                With currentTotal.Offset(-1, conditionalOffset)
                    If IsNumeric(currentTotal.Value) Then
                        .Value = currentTotal.Offset(, 2).Value
                    Else
                        .Value = 0
                    End If
                    .NumberFormat = "0.0"
                End With
            End If
continue:
        Next total
        currentOffset = currentOffset + offsetIncrement
    Wend
End Sub

Private Sub insertMeanRows()
    Dim total As Variant
    
    For Each total In totalRows
        total.EntireRow.Insert
        total.Offset(-1, -1).Value = "Mean"
    Next total
End Sub

Private Sub fixHeaders()
    Dim currentHeader As Range
    Dim nextHeader As Range
    
    With controlSheet
        Set currentHeader = .Range("B" & initialRow).End(xlUp)
        .Rows(currentHeader.Row).Delete
    End With
End Sub

Sub main(Optional outputType As Integer = 0, Optional includeCounts As Boolean = False)
Attribute main.VB_ProcData.VB_Invoke_Func = "X\n14"
    Dim batch As Variant
    Dim isPercentSort As Boolean
    
    initialize
    isPercentSort = outputType = outputTypes.PERCENT_SORTED
    If initialized = True Then
        For Each batch In batchCols
            processBatch batch, isPercentSort, includeCounts
        Next batch
        If Not includeCounts Then Call fixHeaders
        If outputType > 1 Then averagesModule.main outputType, controlSheet
        controlSheet.Cells.Font.Name = "Gotham Book"
        finalize
    End If
End Sub
