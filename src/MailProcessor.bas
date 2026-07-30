Attribute VB_Name = "MailProcessor"
'==============================================================================
' MailProcessor
' Central entry points for startup scans and per-message processing.
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
' Startup: optionally scan the Inbox for items that still need processing
'------------------------------------------------------------------------------

Public Sub ProcessInboxOnStartup()
    Dim items As Outlook.Items
    Dim i As Long
    Dim mail As Outlook.MailItem

    If Not m_Initialized Then Initialize

    ' Restrict to unread mail to keep startup light. Expand later if needed.
    Set items = m_Inbox.items.Restrict("[Unread] = true")
    items.Sort "[ReceivedTime]", True

    Logger.Log "Startup scan: " & items.Count & " unread item(s)."

    ' Walk backwards so moves/deletes do not skip items.
    For i = items.Count To 1 Step -1
        If TypeOf items(i) Is Outlook.MailItem Then
            Set mail = items(i)
            ProcessMailItem mail
        End If
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
            If TypeOf item Is Outlook.MailItem Then
                ProcessMailItem item
            End If
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
' Single-message pipeline — evaluate rules in priority order
'------------------------------------------------------------------------------

Public Sub ProcessMailItem(ByRef mail As Outlook.MailItem)
    Dim action As MailAction
    Dim targetFolder As Outlook.Folder

    If mail Is Nothing Then Exit Sub

    action = MailRules.Evaluate(mail)

    Select Case action.ActionType
        Case maNone
            ' No matching rule — leave the message alone.

        Case maMarkRead
            mail.UnRead = False
            mail.Save
            Logger.Log "Marked read: " & mail.Subject

        Case maMove
            Set targetFolder = FolderHelpers.GetOrCreateFolder(action.FolderPath)
            If Not targetFolder Is Nothing Then
                mail.Move targetFolder
                Logger.Log "Moved to [" & action.FolderPath & "]: " & mail.Subject
            End If

        Case maDelete
            ' Soft-delete: moves to Deleted Items. Use mail.Delete after
            ' PermanentDelete if hard-delete is required later.
            mail.Delete
            Logger.Log "Deleted: " & mail.Subject

        Case maMarkReadAndMove
            mail.UnRead = False
            mail.Save
            Set targetFolder = FolderHelpers.GetOrCreateFolder(action.FolderPath)
            If Not targetFolder Is Nothing Then
                mail.Move targetFolder
                Logger.Log "Marked read & moved to [" & action.FolderPath & "]: " & mail.Subject
            End If
    End Select
End Sub
