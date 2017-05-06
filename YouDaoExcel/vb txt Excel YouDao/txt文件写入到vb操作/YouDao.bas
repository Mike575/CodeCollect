Attribute VB_Name = "Module1"
Private Type Character
    word As String
    trans As String
    phonetic As String
    tags As String
    'progress As Integer
End Type

'汇出有道xml格式单词库文件
Sub xmlVocabulary()

    Dim newChar As Character
    Dim R As Range
    Dim Row As Range
    Dim strOutput As String
    Dim arrBytes() As Byte

    newChar.tags = ActiveSheet.Name
    ActiveSheet.Names.Add Name:="NewWord", RefersTo:="=OFFSET($A$1,0,0,COUNTA($A:$A))"
    Set R = ActiveSheet.Names("NewWord").RefersToRange
    
    strOutput = "<wordbook>"
    For Each Row In R.Rows
        newChar.word = Trim(Row(1))
        Call searchWord(newChar.word, newChar.trans, newChar.phonetic)
        strOutput = strOutput & vbCrLf & "<item>"
        strOutput = strOutput & vbCrLf & "<word>" & newChar.word & "</word>"
        strOutput = strOutput & vbCrLf & "<trans><![CDATA[" & newChar.trans & "]]></trans>"
        strOutput = strOutput & vbCrLf & "<phonetic><![CDATA[" & newChar.phonetic & "]]></phonetic>"
        strOutput = strOutput & vbCrLf & "<tags>" & newChar.tags & "</tags>"
        strOutput = strOutput & vbCrLf & "<progress>1</progress>"
        strOutput = strOutput & vbCrLf & "</item>"
    Next Row
    strOutput = strOutput & vbCrLf & "</wordbook>"
    
    arrBytes = ChrW(&HFEFF) & strOutput     '写入unicode文字码
    
    Open Application.ActiveWorkbook.Path & "\" & newChar.tags & ".xml" For Binary As #1      '建立xml格式档案
    Put #1, , arrBytes
    
    Close #1

End Sub

'单词音译写入Excel档
Sub xlsmVocabulary()

    Dim newChar As Character
    Dim R As Range
    Dim Row As Range
    Dim rr As Integer

    strTags = ActiveSheet.Name
    ActiveSheet.Names.Add Name:="NewWord", RefersTo:="=OFFSET($B$1,0,0,COUNTA($B:$B))"
    Set R = ActiveSheet.Names("NewWord").RefersToRange
    
    rr = 0
    
    For Each Row In R.Rows
        rr = rr + 1
        newChar.word = Trim(Row(1))
        
        Call searchWord(newChar.word, newChar.trans, newChar.phonetic)
    
        Worksheets(strTags).Cells(rr, 13).Value = TrimEnglish(newChar.trans)      '撷取翻译
        
    Next Row
    
End Sub

Function TrimEnglish(trans As String) As String
    Dim SenArray() As String
    Dim iArray As Integer
    Dim TrimEnglishpre As String
    
    
    SenArray = Split(trans)
    iArray = 0
    TrimEnglish = ""
    TrimEnglishpre = ""
    
    Do While iArray <= UBound(SenArray)
        
        TrimEnglishpre = TrimEnglishpre + SenArray(iArray)
        
        iArray = iArray + 1
    Loop
    
    i = 0
    Do While i < Len(TrimEnglishpre)
        i = i + 1
        temps = Mid(TrimEnglishpre, i, 1)
        If i <= 2 Then
            If (Asc(temps) < 0 And temps <> "" And temps <> " " And temps <> "【" And temps <> "】" And temps <> "有" And temps <> "机" And temps <> "学" And temps <> "无" And temps <> "" And temps <> " " And temps <> " 亦" And temps <> " 作") Or temps = " 1" Or temps = " 2" Or temps = " 3" Or temps = " 4" Or temps = " 5" Or temps = " 6" Or temps = " 7" Or temps = " 8" Or temps = " 9" Or temps = " 0" Or Asc(temps) = 48 Or Asc(temps) = 49 Or Asc(temps) = 50 Or Asc(temps) = 51 Or Asc(temps) = 52 Or Asc(temps) = 53 Or Asc(temps) = 54 Or Asc(temps) = 55 Or Asc(temps) = 56 Or Asc(temps) = 57 Then
                TrimEnglish = TrimEnglish + temps
            End If
        Else
            If (Asc(temps) < 0 And temps <> "" And temps <> " " And temps <> "【" And temps <> "】" And temps <> "有" And temps <> "机" And temps <> "学" And temps <> "无" And temps <> "" And temps <> " " And temps <> "化" And temps <> " 亦" And temps <> " 作") Or temps = " 1" Or temps = " 2" Or temps = " 3" Or temps = " 4" Or temps = " 5" Or temps = " 6" Or temps = " 7" Or temps = " 8" Or temps = " 9" Or temps = " 0" Or Asc(temps) = 48 Or Asc(temps) = 49 Or Asc(temps) = 50 Or Asc(temps) = 51 Or Asc(temps) = 52 Or Asc(temps) = 53 Or Asc(temps) = 54 Or Asc(temps) = 55 Or Asc(temps) = 56 Or Asc(temps) = 57 Then
                TrimEnglish = TrimEnglish + temps
            End If
        End If
        
    Loop
    
        TrimEnglishpre = ""
        TrimEnglishpre = TrimEnglish
        TrimEnglish = ""
        
        i = 0
    Do While i < Len(TrimEnglishpre)
        i = i + 1
        temps = Mid(TrimEnglishpre, i, 1)
        If i <= 2 Then
               If (Asc(temps) < 0 And temps <> "" And temps <> " " And temps <> "【" And temps <> "】" And temps <> "有" And temps <> "机" And temps <> "学" And temps <> "无" And temps <> "" And temps <> " " And temps <> " 亦" And temps <> " 作") Or temps = " 1" Or temps = " 2" Or temps = " 3" Or temps = " 4" Or temps = " 5" Or temps = " 6" Or temps = " 7" Or temps = " 8" Or temps = " 9" Or temps = " 0" Or Asc(temps) = 48 Or Asc(temps) = 49 Or Asc(temps) = 50 Or Asc(temps) = 51 Or Asc(temps) = 52 Or Asc(temps) = 53 Or Asc(temps) = 54 Or Asc(temps) = 55 Or Asc(temps) = 56 Or Asc(temps) = 57 Or temps = "-" Then
                TrimEnglish = TrimEnglish + temps
                End If
        Else
            If (Asc(temps) < 0 And temps <> "" And temps <> " " And temps <> "【" And temps <> "】" And temps <> "有" And temps <> "机" And temps <> "学" And temps <> "无" And temps <> "" And temps <> " " And temps <> "化" And temps <> " 亦" And temps <> " 作") Or temps = " 1" Or temps = " 2" Or temps = " 3" Or temps = " 4" Or temps = " 5" Or temps = " 6" Or temps = " 7" Or temps = " 8" Or temps = " 9" Or temps = " 0" Or Asc(temps) = 48 Or Asc(temps) = 49 Or Asc(temps) = 50 Or Asc(temps) = 51 Or Asc(temps) = 52 Or Asc(temps) = 53 Or Asc(temps) = 54 Or Asc(temps) = 55 Or Asc(temps) = 56 Or Asc(temps) = 57 Then
                TrimEnglish = TrimEnglish + temps
            End If
        End If
        
    Loop
    

End Function


    
    Sub searchWord(tmpWord As String, tmpTrans As String, tmpPhonetic As String)
    'http://dict.youdao.com/search?q=单词&keyfrom=dict.index
        Dim XH As Object
        Dim s() As String
        Dim str_tmp As String
        Dim str_base As String
        
        tmpTrans = ""
        tmpPhonetic = ""

        '开启网页
        Set XH = CreateObject("Microsoft.XMLHTTP")
        On Error Resume Next
        XH.Open "get", "http://dict.youdao.com/search?q=" & tmpWord & "&keyfrom=dict.index", False
        XH.send
        On Error Resume Next
        str_base = XH.responseText
        XH.Close
        Set XH = Nothing
        
        str_base = Split(Split(XH.responseText, "<div id=""webTrans"" class=""trans-wrapper trans-tab"">")(0), "<span class=""keyword"">")(1)

        '撷取音标
        If UBound(Split(str_base, "<span class=""pronounce"">美")) = 1 Then
        '美式音标
            tmpPhonetic = Split((Split(Split(str_base, "<span class=""pronounce"">美")(1), "<span class=""phonetic"">")(1)), "</span>")(0)
            On Error Resume Next
        Else
            tmpPhonetic = Split((Split(str_base, "<span class=""phonetic"">")(1)), "</span>")(0)
            On Error Resume Next
        End If

        '撷取中文翻译
        str_tmp = Split((Split(str_base, "<div class=""trans-container"">")(1)), "</div>")(0)
        str_tmp = Split((Split(str_tmp, "<ul>")(1)), "</ul>")(0)
        s = Split(str_tmp, "<li>")
        tmpTrans = Split(s(LBound(s) + 1), "</li")(0)
        For i = LBound(s) + 2 To UBound(s)
            tmpTrans = tmpTrans & Chr(10) & Split(s(i), "</li")(0)
        Next

    End Sub

