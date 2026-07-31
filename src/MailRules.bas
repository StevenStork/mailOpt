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
    maMarkRead = 1
    maMove = 2
    maDelete = 3
    maMarkReadAndMove = 4
    maFlag = 5
End Enum

Public Type MailAction
    ActionType As MailActionType
    FolderPath As String   ' Used by maMove / maMarkReadAndMove (e.g. "Newsletters")
    RuleName As String     ' For logging / debugging
End Type

Private Const MSG_MEETING_ACCEPTED As String = "IPM.Schedule.Meeting.Resp.Pos"
Private Const MSG_MEETING_TENTATIVE As String = "IPM.Schedule.Meeting.Resp.Tent"
Private Const MSG_MEETING_DECLINED As String = "IPM.Schedule.Meeting.Resp.Neg"

'------------------------------------------------------------------------------
' Load / refresh rule configuration
' (Hook for future: read from a worksheet, JSON file, or registry.)
'------------------------------------------------------------------------------

Public Sub LoadRules()
    Logger.Log "MailRules loaded."
End Sub

'------------------------------------------------------------------------------
' Evaluate any Outlook item (mail, meeting response, etc.)
' Meeting-response rules run first; remaining mail uses Evaluate.
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

    If TypeOf item Is Outlook.MailItem Then
        EvaluateItem = Evaluate(item)
    Else
        EvaluateItem = result
    End If
End Function

'------------------------------------------------------------------------------
' Evaluate a mail item against rules in priority order.
' First matching rule wins. Return maNone when nothing matches.
'------------------------------------------------------------------------------

Public Function Evaluate(ByRef mail As Outlook.MailItem) As MailAction
    Dim result As MailAction

    result.ActionType = maNone
    result.FolderPath = vbNullString
    result.RuleName = vbNullString

    '--- Priority 1: Delete ---------------------------------------------------
    If MatchesDeleteRule(mail) Then
        result.ActionType = maDelete
        result.RuleName = "Delete"
        Evaluate = result
        Exit Function
    End If

    '--- Priority 2: Move (optionally mark read) ------------------------------
    If MatchesMoveRule(mail, result.FolderPath) Then
        If MatchesMarkReadRule(mail) Then
            result.ActionType = maMarkReadAndMove
            result.RuleName = "MarkReadAndMove"
        Else
            result.ActionType = maMove
            result.RuleName = "Move"
        End If
        Evaluate = result
        Exit Function
    End If

    '--- Priority 3: Mark read only -------------------------------------------
    If MatchesMarkReadRule(mail) Then
        result.ActionType = maMarkRead
        result.RuleName = "MarkRead"
        Evaluate = result
        Exit Function
    End If

    Evaluate = result
End Function

'==============================================================================
' Rule matchers — replace stubs with real sender / subject / domain checks.
' Keep each matcher focused on one concern so rules stay easy to extend.
'==============================================================================

Private Function MatchesDeleteRule(ByRef mail As Outlook.MailItem) As Boolean
    ' Example (disabled):
    ' If InStr(1, LCase$(mail.SenderEmailAddress), "noreply-spam@example.com", vbTextCompare) > 0 Then
    '     MatchesDeleteRule = True
    ' End If
    MatchesDeleteRule = False
End Function

Private Function MatchesMoveRule(ByRef mail As Outlook.MailItem, ByRef folderPath As String) As Boolean
    folderPath = vbNullString
    MatchesMoveRule = False

    ' ES Comms — username ceo.inc, or sender host prefix communications.es.
    If SenderUserNameIs(mail, "ceo.inc") Or SenderStartsWith(mail, "communications.es") Then
        folderPath = "\\BAE Comms\ES Comms"
        MatchesMoveRule = True
        Exit Function
    End If

    ' NHNA Comms — username sonashua.siteexecutive.
    If SenderUserNameIs(mail, "sonashua.siteexecutive") Then
        folderPath = "\\BAE Comms\NHNA Comms"
        MatchesMoveRule = True
        Exit Function
    End If

    ' ServiceNow notifications — match sender host prefix (not body URLs).
    If SenderStartsWith(mail, "servicenow.us") Then
        folderPath = "\\Tickets\IT Tickets"
        MatchesMoveRule = True
        Exit Function
    End If
End Function

Private Function MatchesMarkReadRule(ByRef mail As Outlook.MailItem) As Boolean
    ' Example (disabled):
    ' If InStr(1, LCase$(mail.SenderEmailAddress), "@notifications.example.com", vbTextCompare) > 0 Then
    '     MatchesMarkReadRule = True
    ' End If
    MatchesMarkReadRule = False
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
