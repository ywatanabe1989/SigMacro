Option Explicit

Function FlagOn(flag As Long) As Long
    FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function

Function FlagOff(flag As Long) As Long
    FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function

Function ParseArguments() As String()
    ' Read arguments from the text file and parse them
    Dim filePath As String
    Dim fileNum As Integer
    Dim argsString As String
    Dim argArray() As String
    
    ' Get the path to the arguments file
    filePath = "C:\Users\wyusu\Documents\SigmaPlot\SPW12\arguments.txt"
    
    ' Check if the arguments file exists
    If Dir(filePath) <> "" Then
        ' Open and read the file
        fileNum = FreeFile
        Open filePath For Input As #fileNum
        Line Input #fileNum, argsString
        Close #fileNum
        
        ' Split the comma-separated arguments into an array
        argArray = Split(argsString, ",")
        
        ' Trim whitespace from each argument
        Dim i As Integer
        For i = 0 To UBound(argArray)
            argArray(i) = Trim(argArray(i))
        Next i
        
        ParseArguments = argArray
    Else
        ' Return empty array if file doesn't exist
        ReDim argArray(0)
        ParseArguments = argArray
    End If
End Function


Function helloWorld() As String
    ' Get parsed arguments
    Dim args() As String
    args = ParseArguments()
    
    ' Initialize message
    Dim message As String
    message = "Hello World from 3!"
    
    ' Check if we have arguments
    If UBound(args) >= 0 And Len(Join(args, "")) > 0 Then
        ' Append arguments to the message
        message = message & " (arguments: " & Join(args, ", ") & ")"
    End If
    
    ' Display the message box
    MsgBox message
    
    ' Return the message
    helloWorld = message
End Function

Sub Main()
    On Error GoTo ErrorHandler
    Dim result As String
    result = helloWorld()
    ' Do something with result if needed
    Exit Sub
ErrorHandler:
    MsgBox "An error has occurred: " & Err.Description
End Sub