Attribute VB_Name = "SortRuleUI"
'==============================================================================
' SortRuleUI
' Entry points for the Add Sort Rule user form.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Macro: open the form for the currently opened / selected mail.
' Run via Alt+F8 → AddSortRuleFromCurrentMail
'------------------------------------------------------------------------------

Public Sub AddSortRuleFromCurrentMail()
    Dim mail As Outlook.MailItem

    Set mail = GetCurrentMailItem()
    If mail Is Nothing Then
        MsgBox "Open or select a mail message first.", vbExclamation, "Add Sort Rule"
        Exit Sub
    End If

    Load frmAddSortRule
    frmAddSortRule.LoadFromMail mail
    frmAddSortRule.Show vbModal
    Unload frmAddSortRule
End Sub

' Prefer the open Inspector window; otherwise the first Explorer selection.
Public Function GetCurrentMailItem() As Outlook.MailItem
    Dim item As Object
    Dim explorer As Outlook.Explorer
    Dim selection As Outlook.Selection

    Set GetCurrentMailItem = Nothing

    On Error Resume Next

    If Not Application.ActiveInspector Is Nothing Then
        Set item = Application.ActiveInspector.CurrentItem
        If TypeOf item Is Outlook.MailItem Then
            Set GetCurrentMailItem = item
            Exit Function
        End If
    End If

    Set explorer = Application.ActiveExplorer
    If explorer Is Nothing Then Exit Function

    Set selection = explorer.Selection
    If selection Is Nothing Then Exit Function
    If selection.Count < 1 Then Exit Function

    Set item = selection(1)
    If TypeOf item Is Outlook.MailItem Then
        Set GetCurrentMailItem = item
    End If

    On Error GoTo 0
End Function
