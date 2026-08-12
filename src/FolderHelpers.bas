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

Public Function GetMailStoreRoot() As Outlook.Folder
    Set GetMailStoreRoot = GetDefaultInbox().Parent
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
        End If

        Set parent = child
NextPart:
    Next i

    Set GetOrCreateFolder = parent
End Function

Public Function FolderExists(ByVal folderPath As String) As Boolean
    FolderExists = Not GetFolder(folderPath) Is Nothing
End Function

' Resolve an existing folder without creating missing segments.
' Same path rules as GetOrCreateFolder.
Public Function GetFolder(ByVal folderPath As String) As Outlook.Folder
    Dim parent As Outlook.Folder
    Dim parts() As String
    Dim i As Long
    Dim name As String
    Dim child As Outlook.Folder

    Set GetFolder = Nothing
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

    Set GetFolder = parent
End Function

' Immediate child folder names under parentPath (sorted A–Z).
Public Function ListChildFolderNames(ByVal parentPath As String) As Collection
    Dim parent As Outlook.Folder
    Dim child As Outlook.Folder
    Dim names As Collection
    Dim sorted() As String
    Dim i As Long
    Dim j As Long
    Dim tmp As String
    Dim count As Long

    Set names = New Collection
    Set ListChildFolderNames = names

    Set parent = GetFolder(parentPath)
    If parent Is Nothing Then Exit Function

    count = parent.Folders.Count
    If count = 0 Then Exit Function

    ReDim sorted(1 To count)
    i = 0
    For Each child In parent.Folders
        i = i + 1
        sorted(i) = child.Name
    Next child

    ' Simple insertion sort — folder counts are small.
    For i = 2 To count
        tmp = sorted(i)
        j = i - 1
        Do While j >= 1
            If StrComp(sorted(j), tmp, vbTextCompare) <= 0 Then Exit Do
            sorted(j + 1) = sorted(j)
            j = j - 1
        Loop
        sorted(j + 1) = tmp
    Next i

    For i = 1 To count
        names.Add sorted(i)
    Next i
End Function
