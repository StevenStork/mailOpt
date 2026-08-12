Attribute VB_Name = "SortRules"
'==============================================================================
' SortRules
' Load sender → folder mappings from text files (Email, Destination, Name).
'==============================================================================
Option Explicit

Private Const SORT_RULES_FOLDER As String = "C:\mailOpt\sortRules\"

Private Const SORT_COMMS_FILE As String = "sortComms"
Private Const SORT_TICKETS_FILE As String = "sortTickets"
Private Const SORT_PRODUCT_LINES_FILE As String = "sortProductLines"

Private Const BAE_COMMS_PARENT As String = "\\BAE Comms"
Private Const TICKETS_PARENT As String = "\\Tickets"
Private Const PROGRAM_GROUPS_PARENT As String = "\\Program Groups"

Private m_CommsRules As Object
Private m_TicketsRules As Object
Private m_ProductLineRules As Object

'------------------------------------------------------------------------------
' Load / reload
'------------------------------------------------------------------------------

Public Sub LoadAllSortRules()
    Set m_CommsRules = LoadSortRulesFile(SORT_COMMS_FILE)
    Set m_TicketsRules = LoadSortRulesFile(SORT_TICKETS_FILE)
    Set m_ProductLineRules = LoadSortRulesFile(SORT_PRODUCT_LINES_FILE)
End Sub

Public Sub ReloadAllSortRules()
    LoadAllSortRules
End Sub

'------------------------------------------------------------------------------
' Form / UI helpers — sort file catalog and write-back
'------------------------------------------------------------------------------

' Base names: sortComms, sortTickets, sortProductLines
Public Function SortFileNames() As Variant
    SortFileNames = Array(SORT_COMMS_FILE, SORT_TICKETS_FILE, SORT_PRODUCT_LINES_FILE)
End Function

Public Function ParentFolderForSortFile(ByVal fileName As String) As String
    Select Case LCase$(Trim$(fileName))
        Case LCase$(SORT_COMMS_FILE)
            ParentFolderForSortFile = BAE_COMMS_PARENT
        Case LCase$(SORT_TICKETS_FILE)
            ParentFolderForSortFile = TICKETS_PARENT
        Case LCase$(SORT_PRODUCT_LINES_FILE)
            ParentFolderForSortFile = PROGRAM_GROUPS_PARENT
        Case Else
            ParentFolderForSortFile = vbNullString
    End Select
End Function

' Parent folder name for UI dropdowns (e.g. "BAE Comms"), no \\ prefix.
Public Function ParentFolderDisplayName(ByVal fileName As String) As String
    ParentFolderDisplayName = StripRootPrefix(ParentFolderForSortFile(fileName))
End Function

' Map a parent display name ("BAE Comms" or "\\BAE Comms") back to sortComms / etc.
Public Function SortFileNameFromParentDisplay(ByVal displayName As String) As String
    Dim files As Variant
    Dim i As Long
    Dim candidate As String

    displayName = StripRootPrefix(Trim$(displayName))
    If Len(displayName) = 0 Then Exit Function

    files = SortFileNames()
    For i = LBound(files) To UBound(files)
        candidate = ParentFolderDisplayName(CStr(files(i)))
        If StrComp(candidate, displayName, vbTextCompare) = 0 Then
            SortFileNameFromParentDisplay = CStr(files(i))
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

' Add or update a row in the chosen sort file, then reload in-memory rules.
' Returns True on success.
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

            ' Preserve blank trailing lines lightly: skip empty only if we have content after.
            If Len(Trim$(line)) = 0 Then
                ' Keep a single blank line only when it is not the last content row.
                GoTo NextUpsertLine
            End If

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

    ReloadAllSortRules
    UpsertSortRule = True
End Function

' Resolve path, creating the rules folder and an empty headered file when missing.
Private Function EnsureSortRulesFilePath(ByVal fileName As String) As String
    Dim base As String
    Dim candidate As String
    Dim folderPath As String

    base = SORT_RULES_FOLDER
    If Right$(base, 1) <> "\" Then base = base & "\"
    folderPath = Left$(base, Len(base) - 1)

    If Not EnsureFolderTree(folderPath) Then
        EnsureSortRulesFilePath = vbNullString
        Exit Function
    End If

    candidate = ResolveSortRulesFilePath(fileName)
    If Len(candidate) > 0 Then
        EnsureSortRulesFilePath = candidate
        Exit Function
    End If

    candidate = base & fileName & ".txt"
    If Not WriteTextFile(candidate, "Email" & vbTab & "Destination" & vbTab & "Name" & vbCrLf) Then
        EnsureSortRulesFilePath = vbNullString
        Exit Function
    End If

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
    On Error GoTo 0
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
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Match — sender Email column → parent\<Destination>
'------------------------------------------------------------------------------

Public Function MatchCommsRule(ByRef mail As Outlook.MailItem, ByRef folderPath As String) As Boolean
    MatchCommsRule = MatchLoadedRules(mail, m_CommsRules, BAE_COMMS_PARENT, folderPath)
End Function

Public Function MatchTicketsRule(ByRef mail As Outlook.MailItem, ByRef folderPath As String) As Boolean
    MatchTicketsRule = MatchLoadedRules(mail, m_TicketsRules, TICKETS_PARENT, folderPath)
End Function

Public Function MatchProductLineRule(ByRef mail As Outlook.MailItem, ByRef folderPath As String) As Boolean
    MatchProductLineRule = MatchLoadedRules(mail, m_ProductLineRules, PROGRAM_GROUPS_PARENT, folderPath)
End Function

Private Function MatchLoadedRules(ByRef mail As Outlook.MailItem, ByRef rules As Object, _
    ByVal parentFolder As String, ByRef folderPath As String) As Boolean

    Dim addr As String
    Dim userName As String
    Dim dest As String

    folderPath = vbNullString
    MatchLoadedRules = False

    If rules Is Nothing Then Exit Function
    If rules.Count = 0 Then Exit Function

    addr = LCase$(Trim$(MailRules.SenderAddress(mail)))
    userName = LCase$(SenderLocalPart(addr))

    If Len(addr) > 0 Then
        If rules.Exists(addr) Then
            dest = rules(addr)
            folderPath = parentFolder & "\" & dest
            MatchLoadedRules = True
            Exit Function
        End If
    End If

    If Len(userName) > 0 Then
        If rules.Exists(userName) Then
            dest = rules(userName)
            folderPath = parentFolder & "\" & dest
            MatchLoadedRules = True
            Exit Function
        End If
    End If
End Function

'------------------------------------------------------------------------------
' File parsing
'------------------------------------------------------------------------------

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
    dict.CompareMode = 1 ' TextCompare (case-insensitive keys)

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

    ResolveSortRulesFilePath = vbNullString
End Function

Private Function ReadAllLines(ByVal filePath As String) As String()
    Dim fileNum As Integer
    Dim content As String
    Dim lines() As String

    fileNum = FreeFile
    Open filePath For Input As #fileNum
    content = Input$(LOF(fileNum), fileNum)
    Close #fileNum

    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)
    lines = Split(content, vbLf)

    ReadAllLines = lines
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

Private Function SenderLocalPart(ByVal addr As String) As String
    Dim atPos As Long
    atPos = InStr(1, addr, "@")
    If atPos = 0 Then
        SenderLocalPart = addr
    Else
        SenderLocalPart = Left$(addr, atPos - 1)
    End If
End Function
