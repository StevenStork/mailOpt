Attribute VB_Name = "MailProcessor"
'==============================================================================
' MailProcessor
' Central entry points for startup scans, on-demand runs, and per-item processing.
' Rules live in MailRules; folder utilities live in FolderHelpers.
'==============================================================================
Option Explicit

Private m_Initialized As Boolean
Private m_Inbox As Outlook.Folder

'------------------------------------------------------------------------------
' Lifecycle
'------------------------------------------------------------------------------

Public Sub Initialize()
    If m_Initialized Then Exit Sub

    Set m_Inbox = FolderHelpers.GetDefaultInbox()
    MailRules.LoadRules
    m_Initialized = True

    Logger.Log "MailProcessor initialized."
End Sub

Public Sub Shutdown()
    Set m_Inbox = Nothing
    m_Initialized = False
End Sub

'------------------------------------------------------------------------------
' On-demand: run all filters against every Inbox item
' Call from the Immediate Window, a button, or a custom ribbon control:
'   MailProcessor.RunInboxFilters
'------------------------------------------------------------------------------

Public Sub RunInboxFilters()
    Dim items As Outlook.Items
    Dim i As Long

    If Not m_Initialized Then Initialize

    Set items = m_Inbox.Items
    items.Sort "[ReceivedTime]", True

    Logger.Log "Manual inbox run: " & items.Count & " item(s)."

    For i = items.Count To 1 Step -1
        ProcessOutlookItem items(i)
    Next i

    Logger.Log "Manual inbox run complete."
End Sub

'------------------------------------------------------------------------------
' Startup: scan unread Inbox items
'------------------------------------------------------------------------------

Public Sub ProcessInboxOnStartup()
    Dim items As Outlook.Items
    Dim i As Long

    If Not m_Initialized Then Initialize

    Set items = m_Inbox.Items.Restrict("[Unread] = true")
    items.Sort "[ReceivedTime]", True

    Logger.Log "Startup scan: " & items.Count & " unread item(s)."

    For i = items.Count To 1 Step -1
        ProcessOutlookItem items(i)
    Next i
End Sub

'------------------------------------------------------------------------------
' New mail: EntryIDCollection from Application_NewMailEx
'------------------------------------------------------------------------------

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

        If Not item Is Nothing Then
            ProcessOutlookItem item
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
' Single-item pipeline — mail and meeting responses
'------------------------------------------------------------------------------

Public Sub ProcessMailItem(ByRef mail As Outlook.MailItem)
    ProcessOutlookItem mail
End Sub

Public Sub ProcessOutlookItem(ByRef item As Object)
    Dim action As MailAction
    Dim targetFolder As Outlook.Folder

    If item Is Nothing Then Exit Sub

    action = MailRules.EvaluateItem(item)

    Select Case action.ActionType
        Case maNone
            ' No matching rule — leave the item alone.

        Case maMarkRead
            item.UnRead = False
            item.Save
            Logger.Log "Marked read: " & ItemSubject(item)

        Case maMove
            Set targetFolder = FolderHelpers.GetOrCreateFolder(action.FolderPath)
            If Not targetFolder Is Nothing Then
                item.Move targetFolder
                Logger.Log "Moved to [" & action.FolderPath & "]: " & ItemSubject(item)
            End If

        Case maDelete
            item.Delete
            Logger.Log "Deleted: " & ItemSubject(item)

        Case maMarkReadAndMove
            item.UnRead = False
            item.Save
            Set targetFolder = FolderHelpers.GetOrCreateFolder(action.FolderPath)
            If Not targetFolder Is Nothing Then
                item.Move targetFolder
                Logger.Log "Marked read & moved to [" & action.FolderPath & "]: " & ItemSubject(item)
            End If

        Case maFlag
            item.FlagStatus = olFlagMarked
            item.FlagRequest = "Follow up"
            item.Save
            Logger.Log "Flagged: " & ItemSubject(item)
    End Select
End Sub

Private Function ItemSubject(ByRef item As Object) As String
    On Error Resume Next
    ItemSubject = item.Subject
End Function
