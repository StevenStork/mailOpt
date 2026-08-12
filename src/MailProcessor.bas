Attribute VB_Name = "MailProcessor"
'==============================================================================
' MailProcessor
' Startup scans, on-demand runs, and per-item processing.
'==============================================================================
Option Explicit

Private m_Initialized As Boolean

Public Sub Initialize()
    If m_Initialized Then Exit Sub
    MailRules.LoadRules
    m_Initialized = True
End Sub

Public Sub Shutdown()
    m_Initialized = False
End Sub

' On-demand: sort all mail in every folder. Alt+F8 → SortAllEmails
Public Sub SortAllEmails()
    If Not m_Initialized Then Initialize
    MailRules.LoadRules
    ProcessAllFolders False
End Sub

Public Sub ProcessAllFoldersOnStartup()
    If Not m_Initialized Then Initialize
    ProcessAllFolders True
End Sub

Public Sub ProcessNewMail(ByVal EntryIDCollection As String)
    Dim entryIDs() As String
    Dim i As Long
    Dim item As Object
    Dim ns As Outlook.NameSpace

    If Not m_Initialized Then Initialize
    If Len(Trim$(EntryIDCollection)) = 0 Then Exit Sub

    Set ns = Application.GetNamespace("MAPI")
    entryIDs = Split(EntryIDCollection, ",")

    For i = LBound(entryIDs) To UBound(entryIDs)
        Set item = Nothing
        On Error Resume Next
        Set item = ns.GetItemFromID(Trim$(entryIDs(i)))
        On Error GoTo 0

        If Not item Is Nothing Then ProcessOutlookItem item
    Next i
End Sub

Private Sub ProcessAllFolders(ByVal unreadOnly As Boolean)
    ProcessFolderTree FolderHelpers.GetMailStoreRoot(), unreadOnly
End Sub

Private Sub ProcessFolderTree(ByVal folder As Outlook.Folder, ByVal unreadOnly As Boolean)
    Dim subFolder As Outlook.Folder
    Dim items As Outlook.Items
    Dim i As Long

    On Error Resume Next

    If unreadOnly Then
        Set items = folder.Items.Restrict("[Unread] = true")
    Else
        Set items = folder.Items
    End If

    If Err.Number <> 0 Then
        Err.Clear
        GoTo NextFolder
    End If

    On Error GoTo 0
    items.Sort "[ReceivedTime]", True

    For i = items.Count To 1 Step -1
        ProcessOutlookItem items(i)
    Next i

NextFolder:
    On Error Resume Next
    For Each subFolder In folder.Folders
        ProcessFolderTree subFolder, unreadOnly
    Next subFolder
    On Error GoTo 0
End Sub

Public Sub ProcessOutlookItem(ByRef item As Object)
    Dim action As MailAction
    Dim targetFolder As Outlook.Folder

    If item Is Nothing Then Exit Sub

    action = MailRules.EvaluateItem(item)

    Select Case action.ActionType
        Case maNone

        Case maMove
            Set targetFolder = FolderHelpers.GetOrCreateFolder(action.FolderPath)
            If Not targetFolder Is Nothing Then item.Move targetFolder

        Case maMarkReadAndMove
            item.UnRead = False
            item.Save
            Set targetFolder = FolderHelpers.GetOrCreateFolder(action.FolderPath)
            If Not targetFolder Is Nothing Then item.Move targetFolder

        Case maFlag
            item.FlagStatus = olFlagMarked
            item.FlagRequest = "Follow up"
            item.Save
    End Select
End Sub
