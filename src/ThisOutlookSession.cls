VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==============================================================================
' ThisOutlookSession
' Paste this code into the built-in ThisOutlookSession module in Outlook VBA
' (Alt+F11 > Project1 > Microsoft Outlook Objects > ThisOutlookSession).
'
' Wires Application-level events so mail is processed on startup and on arrival.
'==============================================================================
Option Explicit

Private Sub Application_Startup()
    On Error Resume Next
    MailProcessor.Initialize
    MailProcessor.ProcessInboxOnStartup
End Sub

' Fires when one or more new items arrive. EntryIDCollection is a comma-
' separated list of EntryIDs (may include non-mail items).
Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    On Error Resume Next
    MailProcessor.ProcessNewMail EntryIDCollection
End Sub

Private Sub Application_Quit()
    On Error Resume Next
    MailProcessor.Shutdown
End Sub
