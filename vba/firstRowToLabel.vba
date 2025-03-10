Option Explicit

Function FlagOn(flag As Long) As Long
    FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function

Function FlagOff(flag As Long) As Long
    FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Clipboard Settings
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Add this at the top of your module:
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Function GetClipboardText() As String
    Dim hClipMemory As Long
    Dim lpClipMemory As Long
    Dim MyString As String
    Dim RetVal As Long

    OpenClipboard (0&)

    hClipMemory = GetClipboardData(1)
    If hClipMemory > 0 Then
        lpClipMemory = GlobalLock(hClipMemory)
        MyString = Space$(1000)
        RetVal = lstrcpy(MyString, lpClipMemory)
        GlobalUnlock (hClipMemory)
        MyString = Left$(MyString, InStr(1, MyString, Chr$(0)) - 1)
    End If

    CloseClipboard

    GetClipboardText = MyString
End Function

Sub CutAndCopyColumn(columnIndex As Long, ByRef copiedData As String)
    ReDim Selection(3)

    Selection(0) = columnIndex
    Selection(1) = 0
    Selection(2) = columnIndex
    Selection(3) = 0

    ActiveDocument.CurrentDataItem.SelectionExtent = Selection
    ActiveDocument.CurrentItem.Copy
    ActiveDocument.CurrentItem.Clear
    copiedData = GetClipboardText()

End Sub

Sub PasteAndNameColumn(columnIndex As Long, copiedData As String)
    ReDim Selection(3)

    Selection(0) = columnIndex
    Selection(1) = 0
    Selection(2) = columnIndex
    Selection(3) = -1

    ActiveDocument.CurrentDataItem.SelectionExtent = Selection
    ActiveDocument.CurrentDataItem.Paste
    ActiveDocument.CurrentDataItem.Goto 0, columnIndex

    ActiveDocument.CurrentDataItem.DataTable.NamedRanges.Add copiedData, columnIndex, 0, 1, -1, True, True
End Sub

Sub DeleteFirstRow()
    'ActiveDocument.CurrentDataItem.InsertCells(0, 0, 18, 0, InsertDown)
    'ActiveDocument.CurrentDataItem.Open
    ActiveDocument.CurrentDataItem.DeleteCells(0, 0, 18, 0, DeleteUp)
    ActiveDocument.CurrentDataItem.Open
    ActiveDocument.CurrentDataItem.Open
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Main
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Purpose: Process columns in the active document, copying and pasting data
'          while naming ranges. Stops processing after encountering a set
'          number of consecutive empty columns.
'
' Behavior:
' - Iterates through columns up to MaxColumns (128)
' - For each column:
'   1. Cuts and copies the column data
'   2. If data exists, pastes it back and names the range
'   3. If empty, increments the empty column counter
' - Stops if MaxConsecutiveEmptyColumns (7) empty columns are found
' - Deletes the first row after processing
'
' Usage:
' Select a cell of a datasheet and Run this macro
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Main()
    Dim MaxColumns As Long
    Dim columnIndex As Long
    Dim Selection()    
    Dim ConsecutiveEmptyColumnsCounter As Long
    Dim maxConsecutiveEmptyColumns As Long
    Dim copiedData As String

    MaxColumns = 128
    maxConsecutiveEmptyColumns = 7

    For columnIndex = 0 To MaxColumns - 1
        CutAndCopyColumn columnIndex, copiedData

        If Len(copiedData) > 0 Then
            PasteAndNameColumn columnIndex, copiedData
            ConsecutiveEmptyColumnsCounter = 0
        Else
            ConsecutiveEmptyColumnsCounter = ConsecutiveEmptyColumnsCounter + 1
        End If

        If ConsecutiveEmptyColumnsCounter > maxConsecutiveEmptyColumns Then
           MsgBox maxConsecutiveEmptyColumns & " consective empty columns found. Exit."
           Exit For
        End If
    Next columnIndex

    If columnIndex > 0 Then
        MsgBox "Delete the first row called"
        DeleteFirstRow
    End If

End Sub
