VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Code"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'write text
    
    '创建有一个工作表的工作簿
    Set XlsObj = CreateObject("Excel.Application")

    XlsObj.Visible = True
    XlsObj.SheetsInNewWorkbook = 1
     Set XlsBook = XlsObj.Workbooks.Open("C:\Users\57584\Documents\GitHub\YouDaoExcel\vb txt Excel YouDao\pureModify.xls")
     
    '设置活动工作表
    '或者　Set XlsSheet = XlsObj.Worksheets(1)　代表第1个Sheet
    Set XlsSheet = XlsObj.Worksheets("sheet1")
    
    Open "C:\Users\57584\Documents\GitHub\YouDaoExcel\vb txt Excel YouDao\txt文件写入到vb操作\code.txt" For Output As #2
    
    Dim i As Integer
    Dim strRow As String
    Dim strIndex As String
    Dim strChineseName As String
    Dim strEnglishName As String
    Dim strChemicalFormula As String
    i = 0
    strRow = ""
    strIndex = ""
    strChineseName = ""
    strEnglishName = ""
    strChemicalFormula = ""
    
    Do While i < 2114
    i = i + 1
    strRow = "Compone" & "(" & i & ")" & "." & "Row" & "=" & i & vbCrLf
    strIndex = "Compone" & "(" & i & ")" & "." & "Index" & "=" & """" & XlsSheet.Cells(i, 3).Text & """" & vbCrLf
    strChineseName = "Compone" & "(" & i & ")" & "." & "ChineseName" & "=" & """" & XlsSheet.Cells(i, 4).Text & """" & vbCrLf
    strEnglishName = "Compone" & "(" & i & ")" & "." & "EnglishName" & "=" & """" & XlsSheet.Cells(i, 2).Text & """" & vbCrLf
    strChemicalFormula = "Compone" & "(" & i & ")" & "." & "ChemicalFormula" & "=" & """" & XlsSheet.Cells(i, 1).Text & """" & vbCrLf
    Print #2, strRow
    Print #2, strIndex
    Print #2, strChineseName
    Print #2, strEnglishName
    Print #2, strChemicalFormula
    Loop
    Close #2
    
End Sub
