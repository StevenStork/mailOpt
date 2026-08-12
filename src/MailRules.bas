Attribute VB_Name = "MailRules"
'==============================================================================
' MailRules
' Define and evaluate filtering rules.
'
' Matchers start conservative: only explicitly coded rules take action.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Action types returned by Evaluate / EvaluateItem
'------------------------------------------------------------------------------

Public Enum MailActionType
    maNone = 0
    maMove = 1
    maMarkReadAndMove = 2
    maFlag = 3
End Enum

Public Type MailAction
    ActionType As MailActionType
    FolderPath As String   ' Used by maMove / maMarkReadAndMove (e.g. "Newsletters")
    RuleName As String     ' For logging / debugging
End Type

Private Const MSG_MEETING_ACCEPTED As String = "IPM.Schedule.Meeting.Resp.Pos"
Private Const MSG_MEETING_TENTATIVE As String = "IPM.Schedule.Meeting.Resp.Tent"
Private Const MSG_MEETING_DECLINED As String = "IPM.Schedule.Meeting.Resp.Neg"
Private Const MSG_MEETING_REQUEST As String = "IPM.Schedule.Meeting.Request"

'------------------------------------------------------------------------------
' Load / refresh rule configuration
' (Hook for future: read from a worksheet, JSON file, or registry.)
'------------------------------------------------------------------------------

Public Sub LoadRules()
    SortRules.LoadAllSortRules
End Sub

'------------------------------------------------------------------------------
' Evaluate any Outlook item (mail, meeting response, etc.)
'
' Priority: meeting responses → meeting requests → sortComms → sortTickets
'           → sortProductLines → purchase request content rules
'------------------------------------------------------------------------------

Public Function EvaluateItem(ByRef item As Object) As MailAction
    Dim result As MailAction
    Dim msgClass As String

    result.ActionType = maNone
    result.FolderPath = vbNullString
    result.RuleName = vbNullString

    If item Is Nothing Then
        EvaluateItem = result
        Exit Function
    End If

    On Error Resume Next
    msgClass = item.MessageClass
    On Error GoTo 0

    Select Case StrComp(msgClass, MSG_MEETING_ACCEPTED, vbTextCompare)
        Case 0
            result.ActionType = maMarkReadAndMove
            result.FolderPath = "Accepted Invites"
            result.RuleName = "MeetingAccepted"
            EvaluateItem = result
            Exit Function
    End Select

    Select Case StrComp(msgClass, MSG_MEETING_TENTATIVE, vbTextCompare)
        Case 0
            result.ActionType = maMove
            result.FolderPath = "Tentative Invites"
            result.RuleName = "MeetingTentative"
            EvaluateItem = result
            Exit Function
    End Select

    Select Case StrComp(msgClass, MSG_MEETING_DECLINED, vbTextCompare)
        Case 0
            result.ActionType = maFlag
            result.RuleName = "MeetingDeclined"
            EvaluateItem = result
            Exit Function
    End Select

    Select Case StrComp(msgClass, MSG_MEETING_REQUEST, vbTextCompare)
        Case 0
            result.ActionType = maMove
            result.FolderPath = "Meeting Requests"
            result.RuleName = "MeetingRequest"
            EvaluateItem = result
            Exit Function
    End Select

    If TypeOf item Is Outlook.MailItem Then
        EvaluateItem = Evaluate(item)
    Else
        EvaluateItem = result
    End If
End Function

'------------------------------------------------------------------------------
' Evaluate a mail item against sort-file move rules.
' First matching rule wins. Return maNone when nothing matches.
'------------------------------------------------------------------------------

Public Function Evaluate(ByRef mail As Outlook.MailItem) As MailAction
    Dim result As MailAction

    result.ActionType = maNone
    result.FolderPath = vbNullString
    result.RuleName = vbNullString

    If MatchesMoveRule(mail, result.FolderPath) Then
        result.ActionType = maMove
        result.RuleName = "Move"
        Evaluate = result
        Exit Function
    End If

    Evaluate = result
End Function

'==============================================================================
' Rule matchers
'==============================================================================

Private Function MatchesMoveRule(ByRef mail As Outlook.MailItem, ByRef folderPath As String) As Boolean
    folderPath = vbNullString
    MatchesMoveRule = False

    ' Each sort file tries conversation-root sender, then current sender,
    ' then control moves to the next filter below.

    ' BAE Comms — sender mappings from sortComms text file.
    If SortRules.MatchCommsRule(mail, folderPath) Then
        MatchesMoveRule = True
        Exit Function
    End If

    ' Tickets — sender mappings from sortTickets text file.
    If SortRules.MatchTicketsRule(mail, folderPath) Then
        MatchesMoveRule = True
        Exit Function
    End If

    ' Program Groups — sender mappings from sortProductLines text file.
    If SortRules.MatchProductLineRule(mail, folderPath) Then
        MatchesMoveRule = True
        Exit Function
    End If
End Function

'------------------------------------------------------------------------------
' Shared helpers for building matchers
'------------------------------------------------------------------------------

Public Function SenderContains(ByRef mail As Outlook.MailItem, ByVal needle As String) As Boolean
    Dim addr As String
    addr = SenderAddress(mail)
    SenderContains = (InStr(1, addr, needle, vbTextCompare) > 0)
End Function

' True when the local part (before @) equals `userName` (case-insensitive).
Public Function SenderUserNameIs(ByRef mail As Outlook.MailItem, ByVal userName As String) As Boolean
    Dim addr As String
    Dim atPos As Long
    Dim localPart As String

    addr = Trim$(SenderAddress(mail))
    userName = Trim$(userName)
    If Len(addr) = 0 Or Len(userName) = 0 Then Exit Function

    atPos = InStr(1, addr, "@")
    If atPos = 0 Then
        localPart = addr
    Else
        localPart = Left$(addr, atPos - 1)
    End If

    SenderUserNameIs = (StrComp(localPart, userName, vbTextCompare) = 0)
End Function

' True when the sender host (after @) starts with `prefix` (case-insensitive).
' Example: prefix "servicenow.us" matches user@servicenow.us and
' user@servicenow.us.example.com — not body URLs, not mail.servicenow.us.
Public Function SenderStartsWith(ByRef mail As Outlook.MailItem, ByVal prefix As String) As Boolean
    Dim addr As String
    Dim atPos As Long
    Dim host As String

    addr = Trim$(SenderAddress(mail))
    prefix = Trim$(prefix)
    If Len(addr) = 0 Or Len(prefix) = 0 Then Exit Function

    atPos = InStrRev(addr, "@")
    If atPos = 0 Then
        host = addr
    Else
        host = Mid$(addr, atPos + 1)
    End If

    If Len(host) < Len(prefix) Then Exit Function
    SenderStartsWith = (StrComp(Left$(host, Len(prefix)), prefix, vbTextCompare) = 0)
End Function

Public Function SubjectContains(ByRef mail As Outlook.MailItem, ByVal needle As String) As Boolean
    SubjectContains = (InStr(1, mail.Subject, needle, vbTextCompare) > 0)
End Function

Public Function SenderAddress(ByRef mail As Outlook.MailItem) As String
    On Error Resume Next
    If mail.SenderEmailType = "EX" Then
        ' Exchange address — resolve to SMTP when possible
        Dim ae As Outlook.AddressEntry
        Dim exch As Outlook.ExchangeUser
        Set ae = mail.Sender
        If Not ae Is Nothing Then
            Set exch = ae.GetExchangeUser
            If Not exch Is Nothing Then
                SenderAddress = exch.PrimarySmtpAddress
                Exit Function
            End If
        End If
    End If
    SenderAddress = mail.SenderEmailAddress
End Function

'------------------------------------------------------------------------------
' Conversation root sender (for sort-file matching)
'------------------------------------------------------------------------------

Private m_SortSenderMailId As String
Private m_SortRootAddr As String
Private m_SortCurrentAddr As String

' Returns lowercase root sender (when available) and current sender.
' Cached per mail EntryID so sortComms → sortTickets → sortProductLines
' only resolves the conversation once.
Public Sub GetSortSenderAddresses(ByRef mail As Outlook.MailItem, _
    ByRef rootAddr As String, ByRef currentAddr As String)

    Dim entryId As String
    Dim rootMail As Outlook.MailItem

    rootAddr = vbNullString
    currentAddr = vbNullString
    If mail Is Nothing Then Exit Sub

    On Error Resume Next
    entryId = mail.EntryID
    On Error GoTo 0

    If Len(entryId) > 0 And StrComp(entryId, m_SortSenderMailId, vbBinaryCompare) = 0 Then
        rootAddr = m_SortRootAddr
        currentAddr = m_SortCurrentAddr
        Exit Sub
    End If

    currentAddr = LCase$(Trim$(SenderAddress(mail)))

    Set rootMail = ConversationRootMail(mail)
    If Not rootMail Is Nothing Then
        rootAddr = LCase$(Trim$(SenderAddress(rootMail)))
    End If

    m_SortSenderMailId = entryId
    m_SortRootAddr = rootAddr
    m_SortCurrentAddr = currentAddr
End Sub

' Earliest root MailItem in the conversation, or Nothing when unavailable.
Public Function ConversationRootMail(ByRef mail As Outlook.MailItem) As Outlook.MailItem
    Dim parentFolder As Outlook.Folder
    Dim store As Outlook.Store
    Dim conv As Outlook.Conversation
    Dim roots As Outlook.SimpleItems
    Dim item As Object
    Dim candidate As Outlook.MailItem
    Dim best As Outlook.MailItem
    Dim bestTime As Date
    Dim itemTime As Date
    Dim i As Long

    Set ConversationRootMail = Nothing
    If mail Is Nothing Then Exit Function

    On Error Resume Next
    Set parentFolder = mail.Parent
    If parentFolder Is Nothing Then Exit Function

    Set store = parentFolder.Store
    If store Is Nothing Then Exit Function
    If Not store.IsConversationEnabled Then Exit Function

    Set conv = mail.GetConversation
    If conv Is Nothing Then Exit Function

    Set roots = conv.GetRootItems
    If roots Is Nothing Then Exit Function
    If roots.Count = 0 Then Exit Function

    For i = 1 To roots.Count
        Set item = roots.Item(i)
        Set candidate = Nothing
        If TypeOf item Is Outlook.MailItem Then
            Set candidate = item
        End If
        If candidate Is Nothing Then GoTo NextRoot

        itemTime = candidate.ReceivedTime
        If itemTime = 0 Then itemTime = candidate.SentOn

        If best Is Nothing Then
            Set best = candidate
            bestTime = itemTime
        ElseIf itemTime < bestTime Then
            Set best = candidate
            bestTime = itemTime
        End If
NextRoot:
    Next i

    Set ConversationRootMail = best
    On Error GoTo 0
End Function
