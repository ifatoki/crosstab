Attribute VB_Name = "averagesModule"
'@Folder("VBAProject")
Option Explicit

Dim sourceWorkbook As Workbook
Dim sourceSheet As Worksheet
Dim controlSheet As Worksheet
Dim dataGroups() As Range
Dim lastRow As Integer
Dim lastCol As Integer
Dim lastCol2 As Integer
Dim initialRow As Integer
Dim initialized As Boolean

Private Sub initialize(Optional srcSheet As Worksheet)
    Dim dialog As FileDialog
    Dim dialogResult As Long
    
    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
    With dialog
        If srcSheet Is Nothing Then
            .Title = "Select the source file"
            .Filters.Clear
            .Filters.Add "Spreadsheets", "*.xlsx; *.xls", 1
            dialogResult = .Show
            If dialogResult <> 0 Then
                DoEvents
                Set sourceWorkbook = Workbooks.Open(.SelectedItems(1))
            End If
            Set sourceSheet = sourceWorkbook.Worksheets("Output")
        Else
            Set sourceSheet = srcSheet
            Set sourceWorkbook = srcSheet.Parent
        End If
        sourceSheet.Copy After:=sourceSheet
        Set controlSheet = sourceWorkbook.Worksheets("Output (2)")
        controlSheet.Name = "Control"
        sourceSheet.Name = "Output1"
    End With
    
    With controlSheet.UsedRange
        initialRow = .Range("A1").End(xlDown).End(xlDown).Offset(1, 0).Row
        lastRow = .Rows.Count
        lastCol = .Columns.Count
        lastCol2 = .Cells(initialRow, 10000).End(xlToLeft).Column
        If lastCol2 < lastCol Then
            lastCol = lastCol2
        End If
    End With
    dataGroups = getDataGroups()
    initialized = True
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
End Sub

Private Sub finalize()
    Application.StatusBar = ""
    If Not sourceSheet Is Nothing Then sourceSheet.Delete
    If Not controlSheet Is Nothing Then controlSheet.Name = "Output"
    Set sourceWorkbook = Nothing
    Set sourceSheet = Nothing
    Set controlSheet = Nothing
    ReDim dataGroups(1)
    lastCol = 0
    lastRow = 0
    initialRow = 0
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Private Function getDataGroups()
    Dim dataGroups() As Range
    Dim currentGroup As Range
    Dim currentIndex As Integer
    
    currentIndex = -1
    ReDim dataGroups(2000)
    With controlSheet
        Dim firstCell As Range
        Dim lastCell As Range
        
        Set firstCell = .Cells(initialRow, 2)
        While firstCell.Row <= lastRow
            Set lastCell = .Cells(firstCell.End(xlDown).Row - 1, lastCol)
            currentIndex = currentIndex + 1
            Set currentGroup = .Range(firstCell, lastCell)
            Set dataGroups(currentIndex) = currentGroup
            Set firstCell = .Cells(lastCell.Offset(1, 0).End(xlDown).Row, 2)
        Wend
    End With
    ReDim Preserve dataGroups(currentIndex)
    getDataGroups = dataGroups
End Function

Private Function processDataGroup(group)
    Dim topValues() As Variant
    Dim bottomValues() As Variant
    Dim length As Integer
    Dim width As Integer
    Dim unitSize As Integer
    Dim groupValues() As Variant
    Dim i, j As Integer
    Dim topText As String
    Dim bottomText As String
    
    topText = ""
    bottomText = ""
    width = UBound(group.Value2, 2)
    length = UBound(group.Value2, 1)
    unitSize = Int(length / 2)
    groupValues = group.Value2
    ReDim topValues(1 To 1, 1 To width)
    ReDim bottomValues(1 To 1, 1 To width)

    For i = 1 To unitSize
        topText = topText + group(i, 1).Offset(, -1).Value + " & "
        For j = 1 To width
            topValues(1, j) = topValues(1, j) + groupValues(i, j)
            If groupValues(i, j) = "" Then topValues(1, j) = ""
        Next
    Next
    topText = Left(topText, Len(topText) - 3)
    
    For i = unitSize + 2 To length
        bottomText = bottomText + group(i, 1).Offset(, -1).Value + " & "
        For j = 1 To width
            bottomValues(1, j) = bottomValues(1, j) + groupValues(i, j)
            If groupValues(i, j) = "" Then bottomValues(1, j) = ""
        Next
    Next
    bottomText = Left(bottomText, Len(bottomText) - 3)
    
    For i = group.Row + length - 1 To group.Row + unitSize + 2 Step -1
        controlSheet.Rows(i).Delete
    Next
    For i = group.Row + unitSize - 2 To group.Row Step -1
        controlSheet.Rows(i).Delete
    Next
    group.Rows(1).Value2 = topValues
    group.Rows(3).Value2 = bottomValues
    group.Cells(3, 1).Offset(, -1).Value = bottomText
    group.Cells(1, 1).Offset(, -1).Value = topText
End Function

Sub main(Optional srcWorksheet As Worksheet)
    Dim group As Variant
    
    initialize srcWorksheet
    If initialized = True Then
        For Each group In dataGroups
            processDataGroup group
        Next group
    Else
        MsgBox "Script stopped. No file selected.", vbOKOnly + vbExclamation, "Cancelled"
    End If
    finalize
End Sub
