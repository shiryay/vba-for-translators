Attribute VB_Name = "UtilityModule"
Sub KillHiddenText()
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Font.Hidden = True
    Selection.Find.replacement.ClearFormatting
    With Selection.Find
        .Text = "^?"
        .replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub PasteAsIs()
'
' Paste retaining the original formatting
'
'
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
End Sub

Public Function convert_date(ByVal dt As String) As String
'
' Convert date from EU to US and vice versa
'

Dim split_date() As String
Dim sep As String

If InStr(dt, ".") Then sep = "."
If InStr(dt, "/") Then sep = "/"
If InStr(dt, "-") Then sep = "-"

split_date = Split(dt, sep)
convert_date = split_date(1) + sep + split_date(0) + sep + split_date(2)

End Function

Sub ConvertDates()
Dim dates() As String
Dim replacement As String
Dim regExp As Object
Dim logFile As String

logFile = ActiveDocument.Path + "\changes.txt"
Set regExp = CreateObject("vbscript.regexp")
Open logFile For Output As #1

With regExp
    .Pattern = "(\d|\d\d)([\./-])(\d|\d\d)[\./-](\d\d\d\d|\d\d)\b"
    .Global = True
    If .Test(ActiveDocument.Content) Then
        Set regExp_Matches = .Execute(ActiveDocument.Content)
    End If
End With

For Each Match In regExp_Matches
    Selection.Find.ClearFormatting
    Selection.Find.replacement.ClearFormatting
    s = Match & " -> " & convert_date(Match) & Chr(13) & Chr(10)
    Print #1, s
    With Selection.Find
        .Text = Match
        .replacement.Text = convert_date(Match)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
Next
Close #1

End Sub

Public Function URL_Encode(ByRef txt As String) As String
'
' Taken from https://excelvba.ru/code/URLEncode
'
    Dim buffer As String, i As Long, c As Long, n As Long
    buffer = String$(Len(txt) * 12, "%")
    For i = 1 To Len(txt)
        c = AscW(Mid$(txt, i, 1)) And 65535
 
        Select Case c
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95  ' Unescaped 0-9A-Za-z-._ '
                n = n + 1
                Mid$(buffer, n) = ChrW(c)
            Case Is <= 127            ' Escaped UTF-8 1 bytes U+0000 to U+007F '
                n = n + 3
                Mid$(buffer, n - 1) = Right$(Hex$(256 + c), 2)
            Case Is <= 2047           ' Escaped UTF-8 2 bytes U+0080 to U+07FF '
                n = n + 6
                Mid$(buffer, n - 4) = Hex$(192 + (c \ 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
            Case 55296 To 57343       ' Escaped UTF-8 4 bytes U+010000 to U+10FFFF '
                i = i + 1
                c = 65536 + (c Mod 1024) * 1024 + (AscW(Mid$(txt, i, 1)) And 1023)
                n = n + 12
                Mid$(buffer, n - 10) = Hex$(240 + (c \ 262144))
                Mid$(buffer, n - 7) = Hex$(128 + ((c \ 4096) Mod 64))
                Mid$(buffer, n - 4) = Hex$(128 + ((c \ 64) Mod 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
            Case Else                 ' Escaped UTF-8 3 bytes U+0800 to U+FFFF '
                n = n + 9
                Mid$(buffer, n - 7) = Hex$(224 + (c \ 4096))
                Mid$(buffer, n - 4) = Hex$(128 + ((c \ 64) Mod 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
        End Select
    Next
    URL_Encode = Left$(buffer, n)
End Function
