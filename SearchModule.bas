Attribute VB_Name = "SearchModule"
Sub Google()
    search_flag = "Google"
    Search (search_flag)
End Sub

Sub LingueeDe()
    search_flag = "LingueeDeEn"
    Search (search_flag)
End Sub

Sub LingueeRu()
    search_flag = "LingueeRuEn"
    Search (search_flag)
End Sub

Sub LingueeEs()
    search_flag = "LingueeEsEn"
    Search (search_flag)
End Sub

Sub LingueeFr()
    search_flag = "LingueeFrEn"
    Search (search_flag)
End Sub

Sub GoogleTranslate()
    search_flag = "GoogleTr"
    Search (search_flag)
End Sub

Sub SearchProz()
    search_flag = "Proz"
    Search (search_flag)
End Sub

Sub SearchInsurinfo()
    search_flag = "Insur"
    Search (search_flag)
End Sub

Sub SearchColloc()
    search_flag = "Colloc"
    Search (search_flag)
End Sub

Sub SearchMultitran()
    search_flag = "Multitran"
    Search (search_flag)
End Sub

Sub Abkuerzungen()
    search_flag = "Abkuerzungen"
    Search (search_flag)
End Sub

Sub Acronymfinder()
    search_flag = "Acronymfinder"
    Search (search_flag)
End Sub

Private Function selected()
   If WordBasic.GetSelStartPos() = WordBasic.GetSelEndPos() Then
        selected = 0
     Else
        selected = 1
     End If
End Function

Sub mul()
 Dim RetVal
 If selected = 0 Then
     WordBasic.SelectCurWord
 End If

 If selected = 1 Then
     WordBasic.EditCopy
 End If

RetVal = Shell("D:\mt\network\multitran.exe", 1)

End Sub

Public Function Search(ByVal flag As String)
    Dim urls
    Dim arg As String, url As String
    
    Set urls = CreateObject("Scripting.Dictionary")
    
    urls.Add "Google", "https://www.google.ru/search?q=%22{query}%22"
    urls.Add "GoogleTr", "https://translate.google.ru/?sl=auto&tl=en&text={query}&op=translate&hl=en"
    urls.Add "LingueeDeEn", "https://www.linguee.de/deutsch-englisch/search?source=auto&query=%22{query}%22"
    urls.Add "LingueeRuEn", "https://www.linguee.ru/russian-english/search?source=auto&query={query}"
    urls.Add "LingueeEsEn", "https://www.linguee.com/english-spanish/search?source=spanish&query={query}"
    urls.Add "LingueeFrEn", "https://www.linguee.fr/francais-anglais/search?source=auto&query={query}"
    urls.Add "Proz", "https://www.proz.com/search/?term={query}&from=rus&to=eng&es=1"
    urls.Add "Insur", "https://www.insur-info.ru/dictionary/search/?q={query}&btnFind=%C8%F1%EA%E0%F2%FC%21&q_far"
    urls.Add "Colloc", "http://www.ozdic.com/collocation-dictionary/{query}"
    urls.Add "Multitran", "https://www.multitran.com/c/m.exe?CL=1&s={query}&l1=1&l2=2"
    urls.Add "Abkuerzungen", "http://abkuerzungen.de/result.php?searchterm={query}&language=de"
    urls.Add "Acronymfinder", "https://www.acronymfinder.com/{query}.html"

    If selected = 1 Then
        arg = Replace(Selection.Text, vbNewLine, "", , , vbTextCompare) 'new line stripping
        arg = Replace(Selection.Text, "/", "%2F", , , vbTextCompare) 'replacing forward slash to make query address-bar-friendly
        arg = RTrim(arg)
    End If
    If selected = 0 Then
        arg = InputBox("Enter query")
    End If
    If arg = "" Then
        Exit Function
    End If
    
    url = Replace(urls(flag), "{query}", arg, , , vbTextCompare)
    url = UtilityModule.URL_Encode(url)
    
    On Error Resume Next
    ActiveDocument.FollowHyperlink Address:=url

End Function

