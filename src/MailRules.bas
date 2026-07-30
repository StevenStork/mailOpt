Attribute VB_Name = "MailRules"
'==============================================================================
' MailRules
' Define and evaluate filtering rules.
'
' Framework only: stub matchers are intentionally empty / always-false so
' nothing is moved, marked read, or deleted until you add real conditions.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Action types returned by Evaluate
'------------------------------------------------------------------------------

Public Enum MailActionType
    maNone = 0
    maMarkRead = 1
    maMove = 2
    maDelete = 3
    maMarkReadAndMove = 4
End Enum

Public Type MailAction
    ActionType As MailActionType
    FolderPath As String   ' Used by maMove / maMarkReadAndMove (e.g. "Newsletters")
    RuleName As String     ' For logging / debugging
End Type

'------------------------------------------------------------------------------
' Load / refresh rule configuration
' (Hook for future: read from a worksheet, JSON file, or registry.)
'------------------------------------------------------------------------------

Public Sub LoadRules()
    Logger.Log "MailRules loaded (stub — add real rules in Evaluate)."
End Sub

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
    ' Example (disabled):
    ' If InStr(1, LCase$(mail.Subject), "[newsletter]", vbTextCompare) > 0 Then
    '     folderPath = "Newsletters"
    '     MatchesMoveRule = True
    '     Exit Function
    ' End If
    folderPath = vbNullString
    MatchesMoveRule = False
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
