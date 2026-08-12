VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddSortRule
   Caption         =   "Add Sort Rule"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6480
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblSender
      Caption         =   "Sender email"
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1800
   End
   Begin MSForms.TextBox txtSender
      Height          =   300
      Left            =   2160
      TabIndex        =   1
      Top             =   180
      Width           =   3960
   End
   Begin MSForms.Label lblName
      Caption         =   "Display name"
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1800
   End
   Begin MSForms.TextBox txtName
      Height          =   300
      Left            =   2160
      TabIndex        =   3
      Top             =   660
      Width           =   3960
   End
   Begin MSForms.Label lblSortFile
      Caption         =   "Sort text file"
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1800
   End
   Begin MSForms.ComboBox cboSortFile
      Height          =   300
      Left            =   2160
      Style           =   2  'fmStyleDropDownList
      TabIndex        =   5
      Top             =   1140
      Width           =   3960
   End
   Begin MSForms.Label lblDestination
      Caption         =   "Destination folder"
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1800
   End
   Begin MSForms.ComboBox cboDestination
      Height          =   300
      Left            =   2160
      Style           =   0  'fmStyleDropDownCombo
      TabIndex        =   7
      Top             =   1620
      Width           =   3960
   End
   Begin MSForms.Label lblHint
      Caption         =   "Destination lists folders under the parent for the selected sort file. You can also type a new folder name."
      Height          =   600
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   5880
   End
   Begin MSForms.CommandButton btnSave
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   360
      Left            =   3480
      TabIndex        =   9
      Top             =   3000
      Width           =   1200
   End
   Begin MSForms.CommandButton btnCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   4920
      TabIndex        =   10
      Top             =   3000
      Width           =   1200
   End
End
Attribute VB_Name = "frmAddSortRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmAddSortRule
' Map the current mail's sender into a sort text file + destination folder.
'==============================================================================
Option Explicit

Public Sub LoadFromMail(ByRef mail As Outlook.MailItem)
    Dim senderAddr As String
    Dim senderName As String

    senderAddr = Trim$(MailRules.SenderAddress(mail))
    On Error Resume Next
    senderName = Trim$(mail.SenderName)
    On Error GoTo 0

    Me.txtSender.Text = senderAddr
    Me.txtName.Text = senderName
End Sub

Private Sub UserForm_Initialize()
    Dim files As Variant
    Dim i As Long

    files = SortRules.SortFileNames()
    Me.cboSortFile.Clear
    For i = LBound(files) To UBound(files)
        Me.cboSortFile.AddItem SortRules.SortFileDisplayLabel(CStr(files(i)))
    Next i

    Me.cboDestination.Clear
    Me.cboDestination.Enabled = False
End Sub

Private Sub cboSortFile_Change()
    PopulateDestinationFolders
End Sub

Private Sub PopulateDestinationFolders()
    Dim fileName As String
    Dim parentPath As String
    Dim folders As Collection
    Dim i As Long

    Me.cboDestination.Clear
    Me.cboDestination.Enabled = False

    fileName = SortRules.SortFileNameFromDisplayLabel(Me.cboSortFile.Text)
    If Len(fileName) = 0 Then Exit Sub

    parentPath = SortRules.ParentFolderForSortFile(fileName)
    If Len(parentPath) = 0 Then Exit Sub

    Set folders = FolderHelpers.ListChildFolderNames(parentPath)
    For i = 1 To folders.Count
        Me.cboDestination.AddItem CStr(folders(i))
    Next i

    Me.cboDestination.Enabled = True
    If folders.Count > 0 Then
        Me.cboDestination.ListIndex = 0
    Else
        Me.cboDestination.Text = vbNullString
    End If
End Sub

Private Sub btnSave_Click()
    Dim fileName As String
    Dim email As String
    Dim destination As String
    Dim displayName As String

    email = Trim$(Me.txtSender.Text)
    displayName = Trim$(Me.txtName.Text)
    fileName = SortRules.SortFileNameFromDisplayLabel(Me.cboSortFile.Text)
    destination = Trim$(Me.cboDestination.Text)

    If Len(email) = 0 Then
        MsgBox "Sender email is required.", vbExclamation, Me.Caption
        Me.txtSender.SetFocus
        Exit Sub
    End If

    If Len(fileName) = 0 Then
        MsgBox "Choose which sort text file this sender belongs to.", vbExclamation, Me.Caption
        Me.cboSortFile.SetFocus
        Exit Sub
    End If

    If Len(destination) = 0 Then
        MsgBox "Choose or type a destination folder.", vbExclamation, Me.Caption
        Me.cboDestination.SetFocus
        Exit Sub
    End If

    If Not SortRules.UpsertSortRule(fileName, email, destination, displayName) Then
        MsgBox "Could not update the sort file. Check that C:\mailOpt\sortRules\ is writable.", _
               vbCritical, Me.Caption
        Exit Sub
    End If

    MsgBox "Saved " & email & " → " & destination & " in " & fileName & ".", _
           vbInformation, Me.Caption
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub
