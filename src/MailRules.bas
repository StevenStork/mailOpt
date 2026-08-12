Attribute VB_Name = "MailRules"
'==============================================================================
' MailRules
' Rule evaluation, sort-file routing, and Add Sort Rule UI entry points.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Action types
'------------------------------------------------------------------------------

Public Enum MailActionType
    maNone = 0
    maMove = 1
    maMarkReadAndMove = 2
    maFlag = 3
End Enum

Public Type MailAction
    ActionType As MailActionType
    FolderPath As String
    RuleName As String
End Type

Private Const MSG_MEETING_ACCEPTED As String = "IPM.Schedule.Meeting.Resp.Pos"
Private Const MSG_MEETING_TENTATIVE As String = "IPM.Schedule.Meeting.Resp.Tent"
Private Const MSG_MEETING_DECLINED As String = "IPM.Schedule.Meeting.Resp.Neg"
Private Const MSG_MEETING_REQUEST As String = "IPM.Schedule.Meeting.Request"

Private Const SORT_RULES_FOLDER As String = "C:\mailOpt\sortRules\"
Private Const SORT_CATALOG_COUNT As Long = 3

' Parallel catalog: file base name, parent Outlook path, loaded rules dictionary.
Private m_SortFiles(1 To SORT_CATALOG_COUNT) As String
Private m_SortParents(1 To SORT_CATALOG_COUNT) As String
Private m_SortRules(1 To SORT_CATALOG_COUNT) As Object
Private m_CatalogReady As Boolean

' Cached senders for the mail currently being evaluated.
Private m_SortSenderMailId As String
Private m_SortRootAddr As String
Private m_SortCurrentAddr As String

'------------------------------------------------------------------------------
' Load
'------------------------------------------------------------------------------

Public Sub LoadRules()
    Dim i As Long
    EnsureCatalog
    For i = 1 To SORT_CATALOG_COUNT
        Set m_SortRules(i) = LoadSortRulesFile(m_SortFiles(i))
    Next i
End Sub

Private Sub EnsureCatalog()
    If m_CatalogReady Then Exit Sub
    m_SortFiles(1) = "sortComms"
    m_SortParents(1) = "\\BAE Comms"
    m_SortFiles(2) = "sortTickets"
    m_SortParents(2) = "\\Tickets"
    m_SortFiles(3) = "sortProductLines"
    m_SortParents(3) = "\\Program Groups"
    m_CatalogReady = True
End Sub

'------------------------------------------------------------------------------
' Evaluate
' Priority: meeting responses/requests → each sort file (root sender, then
' current sender) → none
'------------------------------------------------------------------------------

Public Function EvaluateItem(ByRef item As Object) As MailAction
    Dim result As MailAction
    Dim msgClass As String
    Dim folderPath As String

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

    Select Case msgClass
        Case MSG_MEETING_ACCEPTED
            result.ActionType = maMarkReadAndMove
            result.FolderPath = "Accepted Invites"
            result.RuleName = "MeetingAccepted"
            EvaluateItem = result
            Exit Function
        Case MSG_MEETING_TENTATIVE
            result.ActionType = maMove
            result.FolderPath = "Tentative Invites"
            result.RuleName = "MeetingTentative"
            EvaluateItem = result
            Exit Function
        Case MSG_MEETING_DECLINED
            result.ActionType = maFlag
            result.RuleName = "MeetingDeclined"
            EvaluateItem = result
            Exit Function
        Case MSG_MEETING_REQUEST
            result.ActionType = maMove
            result.FolderPath = "Meeting Requests"
            result.RuleName = "MeetingRequest"
            EvaluateItem = result
            Exit Function
    End Select

    If TypeOf item Is Outlook.MailItem Then
        If MatchesMoveRule(item, folderPath) Then
            result.ActionType = maMove
            result.FolderPath = folderPath
            result.RuleName = "Move"
        End If
    End If

    EvaluateItem = result
End Function

' For each sort file: try conversation-root sender, then current sender.
Private Function MatchesMoveRule(ByRef mail As Outlook.MailItem, ByRef folderPath As String) As Boolean
    Dim rootAddr As String
    Dim currentAddr As String
    Dim i As Long

    folderPath = vbNullString
    MatchesMoveRule = False
    EnsureCatalog

    GetSortSenderAddresses mail, rootAddr, currentAddr

    For i = 1 To SORT_CATALOG_COUNT
        If m_SortRules(i) Is Nothing Then GoTo NextCatalog
        If m_SortRules(i).Count = 0 Then GoTo NextCatalog

        If MatchAddressAgainstRules(rootAddr, m_SortRules(i), m_SortParents(i), folderPath) Then
            MatchesMoveRule = True
            Exit Function
        End If

        If StrComp(currentAddr, rootAddr, vbBinaryCompare) <> 0 Then
            If MatchAddressAgainstRules(currentAddr, m_SortRules(i), m_SortParents(i), folderPath) Then
                MatchesMoveRule = True
                Exit Function
            End If
        End If
NextCatalog:
    Next i
End Function

Private Function MatchAddressAgainstRules(ByVal addr As String, ByRef rules As Object, _
    ByVal parentFolder As String, ByRef folderPath As String) As Boolean

    Dim userName As String
    Dim dest As String

    MatchAddressAgainstRules = False
    folderPath = vbNullString
    addr = LCase$(Trim$(addr))
    If Len(addr) = 0 Then Exit Function

    If rules.Exists(addr) Then
        dest = rules(addr)
        folderPath = parentFolder & "\" & dest
        MatchAddressAgainstRules = True
        Exit Function
    End If

    userName = SenderLocalPart(addr)
    If Len(userName) > 0 And rules.Exists(userName) Then
        dest = rules(userName)
        folderPath = parentFolder & "\" & dest
        MatchAddressAgainstRules = True
    End If
End Function

'------------------------------------------------------------------------------
' Sender helpers
'------------------------------------------------------------------------------

Public Function SenderAddress(ByRef mail As Outlook.MailItem) As String
    Dim ae As Outlook.AddressEntry
    Dim exch As Outlook.ExchangeUser

    On Error Resume Next
    If mail.SenderEmailType = "EX" Then
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

' Lowercase root sender (when available) and current sender; cached per EntryID.
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

Private Function ConversationRootMail(ByRef mail As Outlook.MailItem) As Outlook.MailItem
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
        If TypeOf item Is Outlook.MailItem Then Set candidate = item
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
End Function

Private Function SenderLocalPart(ByVal addr As String) As String
    Dim atPos As Long
    atPos = InStr(1, addr, "@")
    If atPos = 0 Then
        SenderLocalPart = LCase$(addr)
    Else
        SenderLocalPart = LCase$(Left$(addr, atPos - 1))
    End If
End Function

'------------------------------------------------------------------------------
' Sort catalog (for form + upsert)
'------------------------------------------------------------------------------

Public Function SortFileNames() As Variant
    EnsureCatalog
    SortFileNames = Array(m_SortFiles(1), m_SortFiles(2), m_SortFiles(3))
End Function

Public Function ParentFolderForSortFile(ByVal fileName As String) As String
    Dim i As Long
    EnsureCatalog
    fileName = Trim$(fileName)
    For i = 1 To SORT_CATALOG_COUNT
        If StrComp(m_SortFiles(i), fileName, vbTextCompare) = 0 Then
            ParentFolderForSortFile = m_SortParents(i)
            Exit Function
        End If
    Next i
End Function

' "\\BAE Comms" → "BAE Comms"
Public Function ParentFolderDisplayName(ByVal fileName As String) As String
    ParentFolderDisplayName = StripRootPrefix(ParentFolderForSortFile(fileName))
End Function

Public Function SortFileNameFromParentDisplay(ByVal displayName As String) As String
    Dim i As Long
    EnsureCatalog
    displayName = StripRootPrefix(Trim$(displayName))
    If Len(displayName) = 0 Then Exit Function

    For i = 1 To SORT_CATALOG_COUNT
        If StrComp(StripRootPrefix(m_SortParents(i)), displayName, vbTextCompare) = 0 Then
            SortFileNameFromParentDisplay = m_SortFiles(i)
            Exit Function
        End If
    Next i
End Function

Private Function StripRootPrefix(ByVal folderPath As String) As String
    If Left$(folderPath, 2) = "\\" Then
        StripRootPrefix = Mid$(folderPath, 3)
    Else
        StripRootPrefix = folderPath
    End If
End Function

'------------------------------------------------------------------------------
' Sort file upsert
'------------------------------------------------------------------------------

Public Function UpsertSortRule(ByVal fileName As String, ByVal email As String, _
    ByVal destination As String, ByVal displayName As String) As Boolean

    Dim filePath As String
    Dim lines() As String
    Dim i As Long
    Dim cols() As String
    Dim emailKey As String
    Dim line As String
    Dim found As Boolean
    Dim outText As String
    Dim newLine As String

    UpsertSortRule = False

    email = Trim$(email)
    destination = Trim$(destination)
    displayName = Trim$(displayName)
    fileName = Trim$(fileName)

    If Len(email) = 0 Or Len(destination) = 0 Or Len(fileName) = 0 Then Exit Function
    If Len(ParentFolderForSortFile(fileName)) = 0 Then Exit Function

    filePath = EnsureSortRulesFilePath(fileName)
    If Len(filePath) = 0 Then Exit Function

    emailKey = LCase$(email)
    newLine = email & vbTab & destination & vbTab & displayName

    lines = ReadAllLines(filePath)
    found = False
    outText = vbNullString

    If UBound(lines) >= LBound(lines) Then
        For i = LBound(lines) To UBound(lines)
            line = lines(i)
            If Len(Trim$(line)) = 0 Then GoTo NextUpsertLine

            If IsHeaderLine(line) Then
                outText = outText & line & vbCrLf
                GoTo NextUpsertLine
            End If

            cols = SplitSortLine(line)
            If UBound(cols) - LBound(cols) + 1 >= 1 Then
                If StrComp(LCase$(Trim$(cols(LBound(cols)))), emailKey, vbBinaryCompare) = 0 Then
                    outText = outText & newLine & vbCrLf
                    found = True
                    GoTo NextUpsertLine
                End If
            End If

            outText = outText & line & vbCrLf
NextUpsertLine:
        Next i
    End If

    If Not found Then
        If Len(outText) = 0 Then
            outText = "Email" & vbTab & "Destination" & vbTab & "Name" & vbCrLf
        End If
        outText = outText & newLine & vbCrLf
    End If

    If Not WriteTextFile(filePath, outText) Then Exit Function

    LoadRules
    UpsertSortRule = True
End Function

Private Function EnsureSortRulesFilePath(ByVal fileName As String) As String
    Dim base As String
    Dim candidate As String
    Dim folderPath As String

    base = SORT_RULES_FOLDER
    If Right$(base, 1) <> "\" Then base = base & "\"
    folderPath = Left$(base, Len(base) - 1)

    If Not EnsureFolderTree(folderPath) Then Exit Function

    candidate = ResolveSortRulesFilePath(fileName)
    If Len(candidate) > 0 Then
        EnsureSortRulesFilePath = candidate
        Exit Function
    End If

    candidate = base & fileName & ".txt"
    If Not WriteTextFile(candidate, "Email" & vbTab & "Destination" & vbTab & "Name" & vbCrLf) Then Exit Function
    EnsureSortRulesFilePath = candidate
End Function

Private Function EnsureFolderTree(ByVal folderPath As String) As Boolean
    Dim fso As Object
    Dim parts() As String
    Dim current As String
    Dim i As Long

    EnsureFolderTree = False
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function

    If fso.FolderExists(folderPath) Then
        EnsureFolderTree = True
        Exit Function
    End If

    parts = Split(folderPath, "\")
    current = parts(0)
    For i = 1 To UBound(parts)
        If Len(parts(i)) = 0 Then GoTo NextFolderPart
        current = current & "\" & parts(i)
        If Not fso.FolderExists(current) Then fso.CreateFolder current
        If Not fso.FolderExists(current) Then Exit Function
NextFolderPart:
    Next i

    EnsureFolderTree = fso.FolderExists(folderPath)
End Function

Private Function WriteTextFile(ByVal filePath As String, ByVal content As String) As Boolean
    Dim fso As Object
    Dim ts As Object

    WriteTextFile = False
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function

    Set ts = fso.CreateTextFile(filePath, True, False)
    If ts Is Nothing Then Exit Function
    ts.Write content
    ts.Close
    WriteTextFile = (Err.Number = 0)
End Function

Private Function LoadSortRulesFile(ByVal fileName As String) As Object
    Dim dict As Object
    Dim filePath As String
    Dim lines() As String
    Dim i As Long
    Dim cols() As String
    Dim emailKey As String
    Dim destination As String
    Dim line As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' TextCompare

    filePath = ResolveSortRulesFilePath(fileName)
    If Len(filePath) = 0 Then
        Set LoadSortRulesFile = dict
        Exit Function
    End If

    lines = ReadAllLines(filePath)
    If UBound(lines) < LBound(lines) Then
        Set LoadSortRulesFile = dict
        Exit Function
    End If

    For i = LBound(lines) To UBound(lines)
        line = Trim$(lines(i))
        If Len(line) = 0 Then GoTo NextLine
        If IsHeaderLine(line) Then GoTo NextLine

        cols = SplitSortLine(line)
        If UBound(cols) - LBound(cols) + 1 < 2 Then GoTo NextLine

        emailKey = LCase$(Trim$(cols(LBound(cols))))
        destination = Trim$(cols(LBound(cols) + 1))

        If Len(emailKey) = 0 Or Len(destination) = 0 Then GoTo NextLine
        If Not dict.Exists(emailKey) Then dict.Add emailKey, destination
NextLine:
    Next i

    Set LoadSortRulesFile = dict
End Function

Private Function ResolveSortRulesFilePath(ByVal fileName As String) As String
    Dim base As String
    Dim candidate As String

    base = SORT_RULES_FOLDER
    If Right$(base, 1) <> "\" Then base = base & "\"

    candidate = base & fileName
    If Len(Dir$(candidate)) > 0 Then
        ResolveSortRulesFilePath = candidate
        Exit Function
    End If

    candidate = base & fileName & ".txt"
    If Len(Dir$(candidate)) > 0 Then
        ResolveSortRulesFilePath = candidate
        Exit Function
    End If
End Function

Private Function ReadAllLines(ByVal filePath As String) As String()
    Dim fileNum As Integer
    Dim content As String

    fileNum = FreeFile
    Open filePath For Input As #fileNum
    content = Input$(LOF(fileNum), fileNum)
    Close #fileNum

    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)
    ReadAllLines = Split(content, vbLf)
End Function

Private Function IsHeaderLine(ByVal line As String) As Boolean
    IsHeaderLine = (InStr(1, line, "Email", vbTextCompare) > 0 And _
                    InStr(1, line, "Destination", vbTextCompare) > 0)
End Function

Private Function SplitSortLine(ByVal line As String) As String()
    If InStr(line, vbTab) > 0 Then
        SplitSortLine = Split(line, vbTab)
    ElseIf InStr(line, ",") > 0 Then
        SplitSortLine = Split(line, ",")
    Else
        SplitSortLine = Split(line, vbTab)
    End If
End Function

'------------------------------------------------------------------------------
' Add Sort Rule form
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
    If TypeOf item Is Outlook.MailItem Then Set GetCurrentMailItem = item
End Function
