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
    Set GetDefaultInbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
End Function

Public Function GetMailStoreRoot() As Outlook.Folder
    Set GetMailStoreRoot = GetDefaultInbox().Parent
End Function

Public Function GetOrCreateFolder(ByVal folderPath As String) As Outlook.Folder
    Set GetOrCreateFolder = ResolveFolder(folderPath, True)
End Function

Public Function GetFolder(ByVal folderPath As String) As Outlook.Folder
    Set GetFolder = ResolveFolder(folderPath, False)
End Function

Public Function FolderExists(ByVal folderPath As String) As Boolean
    FolderExists = Not GetFolder(folderPath) Is Nothing
End Function

' createMissing=True creates each missing segment; False returns Nothing if absent.
Private Function ResolveFolder(ByVal folderPath As String, ByVal createMissing As Boolean) As Outlook.Folder
    Dim parent As Outlook.Folder
    Dim parts() As String
    Dim i As Long
    Dim name As String
    Dim child As Outlook.Folder

    Set ResolveFolder = Nothing
    If Len(Trim$(folderPath)) = 0 Then Exit Function

    folderPath = Replace(folderPath, "/", "\")

    If Left$(folderPath, 2) = "\\" Then
        Set parent = GetMailStoreRoot()
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
            If Not createMissing Then Exit Function
            Set child = parent.Folders.Add(name)
        End If

        Set parent = child
NextPart:
    Next i

    Set ResolveFolder = parent
End Function

' Immediate child folder names under parentPath (sorted A-Z).
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
