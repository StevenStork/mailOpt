Attribute VB_Name = "FolderHelpers"
'==============================================================================
' FolderHelpers
' Resolve and create mail folders under the default store / Inbox.
' Folder paths use backslash separators:
'   "Newsletters"              → under Inbox
'   "Projects\Alpha"           → nested under Inbox
'   "\\BAE Comms\ES Comms"     → under mailbox root (sibling of Inbox)
'==============================================================================
Option Explicit

Public Function GetDefaultInbox() As Outlook.Folder
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")
    Set GetDefaultInbox = ns.GetDefaultFolder(olFolderInbox)
End Function

' Returns an existing folder or creates each segment along the path.
' path is relative to the Inbox unless rooted with "\\" for mailbox root.
Public Function GetOrCreateFolder(ByVal folderPath As String) As Outlook.Folder
    Dim parent As Outlook.Folder
    Dim parts() As String
    Dim i As Long
    Dim name As String
    Dim child As Outlook.Folder

    If Len(Trim$(folderPath)) = 0 Then
        Set GetOrCreateFolder = Nothing
        Exit Function
    End If

    folderPath = Replace(folderPath, "/", "\")

    If Left$(folderPath, 2) = "\\" Then
        Set parent = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Parent
        folderPath = Mid$(folderPath, 3)
    Else
        Set parent = GetDefaultInbox()
    End If

    parts = Split(folderPath, "\")

    For i = LBound(parts) To UBound(parts)
        name = Trim$(parts(i))
        If Len(name) = 0 Then GoTo NextPart

        Set child = Nothing
        On Error Resume Next
        Set child = parent.Folders(name)
        On Error GoTo 0

        If child Is Nothing Then
            Set child = parent.Folders.Add(name)
            Logger.Log "Created folder: " & name
        End If

        Set parent = child
NextPart:
    Next i

    Set GetOrCreateFolder = parent
End Function

Public Function FolderExists(ByVal folderPath As String) As Boolean
    Dim parent As Outlook.Folder
    Dim parts() As String
    Dim i As Long
    Dim name As String
    Dim child As Outlook.Folder

    FolderExists = False
    If Len(Trim$(folderPath)) = 0 Then Exit Function

    folderPath = Replace(folderPath, "/", "\")

    If Left$(folderPath, 2) = "\\" Then
        Set parent = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Parent
        folderPath = Mid$(folderPath, 3)
    Else
        Set parent = GetDefaultInbox()
    End If

    parts = Split(folderPath, "\")

    For i = LBound(parts) To UBound(parts)
        name = Trim$(parts(i))
        If Len(name) = 0 Then GoTo NextPart

        Set child = Nothing
        On Error Resume Next
        Set child = parent.Folders(name)
        On Error GoTo 0

        If child Is Nothing Then Exit Function
        Set parent = child
NextPart:
    Next i

    FolderExists = True
End Function
