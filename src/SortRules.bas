Attribute VB_Name = "SortRules"
'==============================================================================
' SortRules
' Load sender → folder mappings from text files (Email, Destination, Name).
'==============================================================================
Option Explicit

Private Const SORT_RULES_FOLDER As String = "C:\mailOpt\sortRules\"

Private Const SORT_PRODUCT_LINES_FILE As String = "sortProductLines"
Private Const PROGRAM_GROUPS_PARENT As String = "\\Program Groups"

Private m_ProductLineRules As Object

'------------------------------------------------------------------------------
' Load / reload
'------------------------------------------------------------------------------

Public Sub LoadProductLineRules()
    Set m_ProductLineRules = LoadSortRulesFile(SORT_PRODUCT_LINES_FILE)
End Sub

Public Sub ReloadAllSortRules()
    LoadProductLineRules
End Sub

'------------------------------------------------------------------------------
' Match — sender Email column → Program Groups\<Destination>
'------------------------------------------------------------------------------

Public Function MatchProductLineRule(ByRef mail As Outlook.MailItem, ByRef folderPath As String) As Boolean
    Dim addr As String
    Dim userName As String
    Dim dest As String

    folderPath = vbNullString
    MatchProductLineRule = False

    If m_ProductLineRules Is Nothing Then Exit Function
    If m_ProductLineRules.Count = 0 Then Exit Function

    addr = LCase$(Trim$(MailRules.SenderAddress(mail)))
    userName = LCase$(SenderLocalPart(addr))

    If Len(addr) > 0 Then
        If m_ProductLineRules.Exists(addr) Then
            dest = m_ProductLineRules(addr)
            folderPath = PROGRAM_GROUPS_PARENT & "\" & dest
            MatchProductLineRule = True
            Exit Function
        End If
    End If

    If Len(userName) > 0 Then
        If m_ProductLineRules.Exists(userName) Then
            dest = m_ProductLineRules(userName)
            folderPath = PROGRAM_GROUPS_PARENT & "\" & dest
            MatchProductLineRule = True
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
