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
' Priority: meeting responses → meeting requests → out of office
'           → sortComms → sortTickets → sortProductLines
'           → purchase request content rules
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

    If IsOutOfOfficeItem(item, msgClass) Then
        result.ActionType = maMove
        result.FolderPath = "Out of Office"
        result.RuleName = "OutOfOffice"
        EvaluateItem = result
        Exit Function
    End If

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

    ' Purchase Requests — subject/body mentions MPR ID or Manual Purchase Requisition.
    If MatchesPurchaseRequestRule(mail) Then
        folderPath = "\\Tickets\Purchase Requests"
        MatchesMoveRule = True
        Exit Function
    End If
End Function

Private Function MatchesPurchaseRequestRule(ByRef mail As Outlook.MailItem) As Boolean
    MatchesPurchaseRequestRule = MailContainsPhrase(mail, "MPR ID") Or _
                                 MailContainsPhrase(mail, "Manual Purchase Requisition")
End Function

' Broad Out of Office / automatic-reply detection (message class + subject/body cues).
Private Function IsOutOfOfficeItem(ByRef item As Object, ByVal msgClass As String) As Boolean
    Dim subject As String
    Dim body As String
    Dim text As String
    Dim autoSubmitted As String

    IsOutOfOfficeItem = False
    If item Is Nothing Then Exit Function

    If InStr(1, msgClass, "Oof", vbTextCompare) > 0 Or _
       InStr(1, msgClass, "OOF", vbTextCompare) > 0 Then
        IsOutOfOfficeItem = True
        Exit Function
    End If

    On Error Resume Next
    subject = item.Subject
    body = item.Body
    autoSubmitted = item.PropertyAccessor.GetProperty( _
        "http://schemas.microsoft.com/mapi/proptag/0x007D001F")
    If Len(autoSubmitted) = 0 Then
        autoSubmitted = item.PropertyAccessor.GetProperty( _
            "http://schemas.microsoft.com/mapi/proptag/0x007D001E")
    End If
    On Error GoTo 0

    If InStr(1, autoSubmitted, "Auto-Submitted:", vbTextCompare) > 0 Then
        If InStr(1, autoSubmitted, "auto-replied", vbTextCompare) > 0 Or _
           InStr(1, autoSubmitted, "auto-generated", vbTextCompare) > 0 Then
            If LooksLikeOutOfOfficeText(subject) Or LooksLikeOutOfOfficeText(Left$(body, 500)) Then
                IsOutOfOfficeItem = True
                Exit Function
            End If
        End If
    End If

    If LooksLikeOutOfOfficeText(subject) Then
        IsOutOfOfficeItem = True
        Exit Function
    End If

    ' Catch replies whose subject is generic but body leads with an OOO notice.
    text = Left$(body, 800)
    If LooksLikeOutOfOfficeText(text) Then
        IsOutOfOfficeItem = True
    End If
End Function

Private Function LooksLikeOutOfOfficeText(ByVal text As String) As Boolean
    Dim t As String
    t = LCase$(Trim$(text))
    If Len(t) = 0 Then Exit Function

    LooksLikeOutOfOfficeText = _
        (InStr(1, t, "out of office", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "out of the office", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "out-of-office", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "automatic reply", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "auto-reply", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "auto reply", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "autoreply", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "auto response", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "auto-response", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "away from the office", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "away from office", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "currently out of the office", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "currently out of office", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "i am currently out", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "i'm currently out", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "i will be out of the office", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "i will be out of office", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "i am out of the office", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "i am out of office", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "i'm out of the office", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "i'm out of office", vbBinaryCompare) > 0) Or _
        (Left$(t, 4) = "ooo:") Or _
        (Left$(t, 4) = "ooo ") Or _
        (Left$(t, 4) = "ooo-") Or _
        (InStr(1, t, "[ooo]", vbBinaryCompare) > 0) Or _
        (InStr(1, t, "(ooo)", vbBinaryCompare) > 0)
End Function

'------------------------------------------------------------------------------
' Shared helpers for building matchers
'------------------------------------------------------------------------------

Public Function SenderContains(ByRef mail As Outlook.MailItem, ByVal needle As String) As Boolean
    Dim addr As String
    addr = SenderAddress(mail)
    SenderContains = (InStr(1, addr, needle, vbTextCompare) > 0)
End Function

' True when subject or plain body contains `phrase` (case-insensitive).
Public Function MailContainsPhrase(ByRef mail As Outlook.MailItem, ByVal phrase As String) As Boolean
    Dim subject As String
    Dim body As String

    phrase = Trim$(phrase)
    If Len(phrase) = 0 Then Exit Function

    On Error Resume Next
    subject = mail.Subject
    body = mail.Body
    On Error GoTo 0

    MailContainsPhrase = (InStr(1, subject, phrase, vbTextCompare) > 0) Or _
                         (InStr(1, body, phrase, vbTextCompare) > 0)
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
