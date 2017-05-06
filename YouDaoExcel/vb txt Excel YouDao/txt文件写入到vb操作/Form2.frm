VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XlsObj As Excel.Application  'Excel对象
Dim XlsBook As Excel.Workbook  '
Dim XlsSheet As Excel.Worksheet  '工作表


Private Sub Form_Load()
    Dim a() As String
    Dim StrOrigin As String
    Dim StrModify As String
    Dim MyArray, Mystring
    Dim EnglishName() As String
    Dim Alias() As String
    Dim Index() As String
    Dim iAlias As Integer
    Dim iArray As Integer
    
    Dim XlsObj As Excel.Application
    '创建有一个工作表的工作簿
    Set XlsObj = CreateObject("Excel.Application")

    XlsObj.Visible = True
    XlsObj.SheetsInNewWorkbook = 1
     Set XlsBook = XlsObj.Workbooks.Open("C:\Users\57584\Desktop\vb txt Excel YouDao\pureModify.xls")
     
    '设置活动工作表
    '或者　Set XlsSheet = XlsObj.Worksheets(1)　代表第1个Sheet
    Set XlsSheet = XlsObj.Worksheets("sheet1")
    
    
    
    
    iAlias = 0
    iArray = 1
    Open "C:\Users\57584\Desktop\vb txt Excel YouDao\txt文件写入到vb操作\pure.txt" For Input As #11
    
    
    Do While Not EOF(11)                                ' do while not the end of essay
    i = i + 1
    ReDim Preserve a(i)
    Line Input #11, a(i)                                'a(i) include content of each line
    Loop
    Close #11
    
    Open "App.Path & aaa.txt" For Output As #12
    
    i = 1
    Do While i < (UBound(a) - 2)                        'check the whole essay ,both of which start from corner mark 1
        i = i + 3
        iAlias = iAlias + 1
        
        StrOrigin = a(i)
        Mystring = Replace(StrOrigin, """", " ")         'delete\"
        MyArray = Split(Mystring, " ", -1, 1)           'split Mystring
        
        ReDim Preserve Alias(iAlias)
        ReDim Preserve EnglishName(iAlias)
        
        Alias(iAlias) = MyArray(0)                      'get Alias
        
        iArray = 1
        Do While iArray < UBound(MyArray)               'seek for Englishname
            iArray = iArray + 1
            If MyArray(iArray) <> " " And MyArray(iArray) <> "" And MyArray(iArray) <> 0 Then
            EnglishName(iAlias) = MyArray(iArray)
                 
            End If
        Loop
    Loop
    
    Dim iMid As Integer
    iAlias = 0
    iArray = 1
    Mystring = ""
    i = 1 + 2
    Do While i < (UBound(a) - 2)                        'check the whole essay ,start from corner mark 1
        i = i + 3                                       'input every third line
        iAlias = iAlias + 1
        
        StrOrigin = Trim(a(i))                          'delete blank
  
        iMid = 0
        Do While iMid < Len(StrOrigin)
        iMid = iMid + 1
        temps = Mid(StrOrigin, iMid, 1)                        'Module数字有问题

            If temps = " " Or temps = "1" Or temps = "2" Or temps = "3" Or temps = "4" Or temps = "5" Or temps = "6" Or temps = "7" Or temps = "8" Or temps = "9" Or temps = "0" Or Asc(temps) = 48 Or Asc(temps) = 49 Or Asc(temps) = 50 Or Asc(temps) = 51 Or Asc(temps) = 52 Or Asc(temps) = 53 Or Asc(temps) = 54 Or Asc(temps) = 55 Or Asc(temps) = 56 Or Asc(temps) = 57 Or temps = "-" Then
                Mystring = Mystring + temps
            ElseIf temps = """" Then
                Mystring = Mystirng + ""
            ElseIf temps = "*" Then
                Mystring = Mystirng + "none"
            Else
                Exit Do
            End If
                
        Loop
        
        MyArray = Split(Trim(Mystring), " ", -1, 1)           'split Mystring
        Mystring = ""
        ReDim Preserve Index(iAlias)
        Index(iAlias) = MyArray(0)                      'get Index
    Loop
    
    
    Print #12, EnglishName(26)
    Close #12
    
    

   
    
    iAlias = 0
    i = 0
    Do While iAlias < UBound(Alias)
    iAlias = iAlias + 1
    i = i + 1
 '   XlsSheet.Cells(i, 1) = Alias(iAlias)
 '   XlsSheet.Cells(i, 2) = EnglishName(iAlias)
    XlsSheet.Cells(i, 6) = Index(iAlias)
    Loop
    
 '   XlsBook.SaveAs FileName:=App.Path & "pureModify.xls"
    
   ' Call xlsmVocabulary
    
End Sub
