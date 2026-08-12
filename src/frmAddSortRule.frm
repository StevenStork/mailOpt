VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddSortRule
   Caption         =   "Add Sort Rule"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6480
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddSortRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmAddSortRule
' Map the current mail's sender into a sort text file + destination folder.
'
' Controls are created at runtime (no .frx / designer controls required).
'==============================================================================
Option Explicit

Private Const fmStyleDropDownCombo As Long = 0
Private Const fmStyleDropDownList As Long = 2

' Runtime controls — WithEvents so Change/Click handlers work.
Private lblSender As MSForms.Label
Private WithEvents txtSender As MSForms.TextBox
Private lblName As MSForms.Label
Private WithEvents txtName As MSForms.TextBox
Private lblParentFolder As MSForms.Label
Private WithEvents cboParentFolder As MSForms.ComboBox
Private lblDestination As MSForms.Label
Private WithEvents cboDestination As MSForms.ComboBox
Private lblHint As MSForms.Label
Private WithEvents btnSave As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton

Private m_Building As Boolean

Public Sub LoadFromMail(ByRef mail As Outlook.MailItem)
    Dim senderAddr As String
    Dim senderName As String

    EnsureControls

    senderAddr = Trim$(MailRules.SenderAddress(mail))
    On Error Resume Next
    senderName = Trim$(mail.SenderName)
    On Error GoTo 0

    txtSender.Text = senderAddr
    txtName.Text = senderName
End Sub

Private Sub UserForm_Initialize()
    EnsureControls
    PopulateParentFolders
End Sub

Private Sub EnsureControls()
    If Not cboParentFolder Is Nothing Then Exit Sub

    m_Building = True

    Set lblSender = Me.Controls.Add("Forms.Label.1", "lblSender", True)
    lblSender.Caption = "Sender email"
    lblSender.Left = 12
    lblSender.Top = 12
    lblSender.Width = 90
    lblSender.Height = 18

    Set txtSender = Me.Controls.Add("Forms.TextBox.1", "txtSender", True)
    txtSender.Left = 108
    txtSender.Top = 10
    txtSender.Width = 198
    txtSender.Height = 20

    Set lblName = Me.Controls.Add("Forms.Label.1", "lblName", True)
    lblName.Caption = "Display name"
    lblName.Left = 12
    lblName.Top = 42
    lblName.Width = 90
    lblName.Height = 18

    Set txtName = Me.Controls.Add("Forms.TextBox.1", "txtName", True)
    txtName.Left = 108
    txtName.Top = 40
    txtName.Width = 198
    txtName.Height = 20

    Set lblParentFolder = Me.Controls.Add("Forms.Label.1", "lblParentFolder", True)
    lblParentFolder.Caption = "Parent folder"
    lblParentFolder.Left = 12
    lblParentFolder.Top = 72
    lblParentFolder.Width = 90
    lblParentFolder.Height = 18

    Set cboParentFolder = Me.Controls.Add("Forms.ComboBox.1", "cboParentFolder", True)
    cboParentFolder.Left = 108
    cboParentFolder.Top = 70
    cboParentFolder.Width = 198
    cboParentFolder.Height = 20
    cboParentFolder.Style = fmStyleDropDownList

    Set lblDestination = Me.Controls.Add("Forms.Label.1", "lblDestination", True)
    lblDestination.Caption = "Destination folder"
    lblDestination.Left = 12
    lblDestination.Top = 102
    lblDestination.Width = 90
    lblDestination.Height = 18

    Set cboDestination = Me.Controls.Add("Forms.ComboBox.1", "cboDestination", True)
    cboDestination.Left = 108
    cboDestination.Top = 100
    cboDestination.Width = 198
    cboDestination.Height = 20
    cboDestination.Style = fmStyleDropDownCombo
    cboDestination.Enabled = False

    Set lblHint = Me.Controls.Add("Forms.Label.1", "lblHint", True)
    lblHint.Caption = "Destination lists folders under the selected parent. You can also type a new folder name."
    lblHint.Left = 12
    lblHint.Top = 132
    lblHint.Width = 294
    lblHint.Height = 36

    Set btnSave = Me.Controls.Add("Forms.CommandButton.1", "btnSave", True)
    btnSave.Caption = "Save"
    btnSave.Left = 174
    btnSave.Top = 180
    btnSave.Width = 60
    btnSave.Height = 24
    btnSave.Default = True

    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    btnCancel.Caption = "Cancel"
    btnCancel.Left = 246
    btnCancel.Top = 180
    btnCancel.Width = 60
    btnCancel.Height = 24
    btnCancel.Cancel = True

    Me.Width = 340
    Me.Height = 250

    m_Building = False
End Sub

Private Sub PopulateParentFolders()
    Dim files As Variant
    Dim i As Long

    m_Building = True
    cboParentFolder.Clear
    files = SortRules.SortFileNames()
    For i = LBound(files) To UBound(files)
        cboParentFolder.AddItem SortRules.ParentFolderDisplayName(CStr(files(i)))
    Next i
    cboDestination.Clear
    cboDestination.Enabled = False
    m_Building = False
End Sub

Private Sub cboParentFolder_Change()
    If m_Building Then Exit Sub
    PopulateDestinationFolders
End Sub

Private Sub PopulateDestinationFolders()
    Dim fileName As String
    Dim parentPath As String
    Dim folders As Collection
    Dim i As Long

    cboDestination.Clear
    cboDestination.Enabled = False

    fileName = SortRules.SortFileNameFromParentDisplay(cboParentFolder.Text)
    If Len(fileName) = 0 Then Exit Sub

    parentPath = SortRules.ParentFolderForSortFile(fileName)
    If Len(parentPath) = 0 Then Exit Sub

    Set folders = FolderHelpers.ListChildFolderNames(parentPath)
    For i = 1 To folders.Count
        cboDestination.AddItem CStr(folders(i))
    Next i

    cboDestination.Enabled = True
    If folders.Count > 0 Then
        cboDestination.ListIndex = 0
    Else
        cboDestination.Text = vbNullString
    End If
End Sub

Private Sub btnSave_Click()
    Dim fileName As String
    Dim parentDisplay As String
    Dim parentPath As String
    Dim destPath As String
    Dim email As String
    Dim destination As String
    Dim displayName As String
    Dim answer As VbMsgBoxResult

    email = Trim$(txtSender.Text)
    displayName = Trim$(txtName.Text)
    parentDisplay = Trim$(cboParentFolder.Text)
    fileName = SortRules.SortFileNameFromParentDisplay(parentDisplay)
    destination = Trim$(cboDestination.Text)

    If Len(email) = 0 Then
        MsgBox "Sender email is required.", vbExclamation, Me.Caption
        txtSender.SetFocus
        Exit Sub
    End If

    If Len(fileName) = 0 Then
        MsgBox "Choose a parent folder.", vbExclamation, Me.Caption
        cboParentFolder.SetFocus
        Exit Sub
    End If

    If Len(destination) = 0 Then
        MsgBox "Choose or type a destination folder.", vbExclamation, Me.Caption
        cboDestination.SetFocus
        Exit Sub
    End If

    parentPath = SortRules.ParentFolderForSortFile(fileName)
    destPath = parentPath & "\" & destination

    If Not FolderHelpers.FolderExists(destPath) Then
        answer = MsgBox( _
            "The folder """ & destination & """ does not exist under " & parentDisplay & "." & vbCrLf & _
            "Create it?", _
            vbYesNo + vbQuestion, Me.Caption)
        If answer <> vbYes Then
            cboDestination.SetFocus
            Exit Sub
        End If

        If FolderHelpers.GetOrCreateFolder(destPath) Is Nothing Then
            MsgBox "Could not create the folder """ & destination & """.", vbCritical, Me.Caption
            Exit Sub
        End If
    End If

    If Not SortRules.UpsertSortRule(fileName, email, destination, displayName) Then
        MsgBox "Could not update the sort rules. Check that C:\mailOpt\sortRules\ is writable.", _
               vbCritical, Me.Caption
        Exit Sub
    End If

    MsgBox "Saved " & email & " -> " & destination & " under " & parentDisplay & ".", _
           vbInformation, Me.Caption
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub
