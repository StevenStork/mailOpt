Attribute VB_Name = "Logger"
'==============================================================================
' Logger
' Lightweight debug logging for the mail framework.
'
' By default writes to the Immediate Window (Ctrl+G in the VBA editor).
' Flip ENABLE_FILE_LOG to True and set LOG_FILE_PATH to also append to a file.
'==============================================================================
Option Explicit

Private Const ENABLE_FILE_LOG As Boolean = False
Private Const LOG_FILE_PATH As String = "C:\Temp\outlook-mail-framework.log"

Public Sub Log(ByVal message As String)
    Dim line As String
    line = Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & message

    Debug.Print line

    If ENABLE_FILE_LOG Then
        AppendToFile line
    End If
End Sub

Public Sub LogError(ByVal source As String, ByVal number As Long, ByVal description As String)
    Log "ERROR [" & source & "] " & number & " — " & description
End Sub

Private Sub AppendToFile(ByVal line As String)
    Dim fileNum As Integer

    On Error Resume Next
    fileNum = FreeFile
    Open LOG_FILE_PATH For Append As #fileNum
    If Err.Number = 0 Then
        Print #fileNum, line
        Close #fileNum
    End If
    On Error GoTo 0
End Sub
