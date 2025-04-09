Option Explicit

Function _ReadCell(columnIndex As Long, rowIndex As Long) As Variant
    Dim dataTable As Object, cellValue As Variant
    Set dataTable = ActiveDocument.CurrentDataItem.DataTable
    cellValue = dataTable.GetData(columnIndex, rowIndex, columnIndex, rowIndex)
    _ReadCell = cellValue(0, 0)
End Function

Function _GetMaxCol() As Long
    Dim maxCol As Long, maxRow As Long, dataTable As Object
    Set dataTable = ActiveDocument.CurrentDataItem.DataTable
    DataTable.GetMaxUsedSize(maxCol, maxRow)
    _GetMaxCol = maxCol
End Function

Function _SetColumnName(columnIndex As Long, columnName As String)
   ActiveDocument.CurrentDataItem.DataTable.NamedRanges.Add columnName, columnIndex, 0, 1, -1, True, True
End Function

Function _DeleteFirstRow()
    ActiveDocument.CurrentDataItem.DeleteCells(0, 0, 31999, 0, DeleteUp)
    ActiveDocument.CurrentDataItem.Open
End Function

Sub FirstRowToColumnNames()
    Dim columnIndex As Long
    Dim firstRowValue As String
    Dim MaxColumns As Long

    MaxColumns = _GetMaxCol()

    For columnIndex = 0 To MaxColumns
        On Error Resume Next
        firstRowValue = CStr(_ReadCell(columnIndex, 0))
        _SetColumnName(columnIndex, firstRowValue)
        On Error GoTo 0
    Next columnIndex

    If columnIndex > 0 Then
        On Error Resume Next       
        _DeleteFirstRow
        On Error GoTo 0        
    End If
End Sub

Sub Main()
   FirstRowToColumnNames()
End Sub
