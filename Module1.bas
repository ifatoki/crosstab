Attribute VB_Name = "Module1"
Option Explicit

Dim sourceWorkbook As Workbook
Dim sourceSheet As Worksheet
Dim controlSheet As Worksheet
Dim totalRows() As Range
Dim batchCols() As Range
Dim lastRow As Integer
Dim lastCol As Integer
Dim firstHeader As Range
Dim initialRow As Integer

Private Sub initialize()
    Dim dialog As FileDialog
    Dim dialogResult As Long
    
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
        If sourceWorkbook Is Nothing Then Set sourceWorkbook = Application.ThisWorkbook
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
    totalRows = getTotals()
    batchCols = getbatches()
    Application.DisplayAlerts = False
End Sub

Private Sub finalize()
    sourceSheet.Activate
    ' controlSheet.Delete
    Application.StatusBar = ""
    Set sourceWorkbook = Nothing
    Set sourceSheet = Nothing
    Set controlSheet = Nothing
    ReDim totalRows(1)
    ReDim batchCols(1)
    Set firstHeader = Nothing
    lastCol = 0
    lastRow = 0
    initialRow = 0
    Application.DisplayAlerts = True
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

Private Sub processBatch(ByVal batch As Range)
    Dim lastBatchCol As Integer
    Dim headerRange As Range
    
    With controlSheet
        lastBatchCol = .Cells(initialRow, batch.Column).End(xlToRight).Column
        Set headerRange = .Range(batch, batch.End(xlDown))
        headerRange.UnMerge
        Set headerRange = .Range(headerRange, headerRange.Offset(0, lastBatchCol - batch.Column))
        headerRange.Copy Destination:=batch.Offset(, 1)
        With .Range(.Cells(firstHeader.Row, batch.Column), .Cells(lastRow, lastBatchCol))
            .Font.Name = "Gotham Book"
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
        End With
    End With
    processTotals batch, lastBatchCol
    deleteColumns batch, lastBatchCol
End Sub

Private Sub deleteColumns(ByVal batch As Range, lastBatchCol As Integer)
    Dim currentCol As Integer
    Dim firstCol As Integer
    
    firstCol = batch.Column
    With controlSheet
        currentCol = lastBatchCol - 1
        .Columns(currentCol + 2).Delete
        While currentCol >= firstCol
            .Columns(currentCol).Delete
            currentCol = currentCol - 2
        Wend
    End With
End Sub

Private Sub processTotals(ByVal batch As Range, lastBatchCol As Integer)
    Dim total As Variant
    Dim batchOffset As Integer
    Dim currentTotal As Range
    Dim currentOffset As Integer
    
    batchOffset = batch.Column - totalRows(0).Column
    currentOffset = 0
    While currentOffset + batch.Column <= lastBatchCol
        For Each total In totalRows
            Set currentTotal = total.Offset(, batchOffset).Offset(, currentOffset)
            With currentTotal.Offset(, 1)
                .Value = Round(currentTotal.Value / 1000000, 6)
                .NumberFormat = "0.0"
            End With
        Next total
        currentOffset = currentOffset + 2
    Wend
End Sub

Private Sub fixHeaders()
    Dim currentHeader As Range
    Dim nextHeader As Range
    
    With controlSheet
        Set currentHeader = .Range("B" & initialRow).End(xlUp)
        While currentHeader.Row > 1
            Set nextHeader = currentHeader.Offset(-1)
            If LCase(Trim(currentHeader.Value)) <> "total" And Trim(currentHeader.Value) <> "" Then
                .Rows(currentHeader.Row).Delete
            End If
            Set currentHeader = nextHeader
        Wend
    End With
End Sub

Sub main()
Attribute main.VB_ProcData.VB_Invoke_Func = "X\n14"
    Dim batch As Variant
    
    initialize
    For Each batch In batchCols
        processBatch batch
    Next batch
    Call fixHeaders
    finalize
End Sub
