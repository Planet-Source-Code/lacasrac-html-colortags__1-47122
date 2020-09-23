VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "HTML TAG Colors"
   ClientHeight    =   4710
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4710
   ScaleWidth      =   5925
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox RTB 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4048
      _Version        =   393217
      BackColor       =   12632256
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"colortag.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------
'---------------------------------------------
'Very simple how to use the RichTextBox ctrl
'ColorTags By Kozari Laszlo 2003/jul/21
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------
'http://lacasrac.srv.hu
'---------------------------------------------
'VOTE on PlanetSource if u like it!!
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

Private Sub Color_Tags(RTB1 As RichTextBox)

'---------------------------------------------
'Create colored tags
'---------------------------------------------

Dim Txt         As String: Txt = RTB1.Text
Dim Tag_start   As Long
Dim Tag_end     As Long
Dim Rgb1        As Long


'---------------------------------------------
'For the 20 tags
For a = 1 To 80
Tag_end = 1: Tag_start = 0

Do
'---------------------------------------------
'20-20-20-20 =80 tag/4=20
'Tags here...
str1 = Choose(a, _
        "<META", "<!--", "<TR", "<TD", "<TABLE", _
        "<FONT", "<SCRIPT", "<TITLE", "<HEAD", "<HTML", _
        "<LINK", "<BODY", "<SELECT", "<OPTION", "<INPUT", _
        "<B", "<A", "<IMG", "<CENTER", "<P", _
        "</META", "<!--", "</TR", "</TD", "</TABLE", _
        "</FONT", "</SCRIPT", "</TITLE", "</HEAD", "</HTML", _
        "</LINK", "</BODY", "</SELECT", "</OPTION", "</INPUT", _
        "</B", "</A", "</IMG", "</CENTER", "</P", _
        "<meta", "<!--", "<tr", "<td", "<table", _
        "<font", "<script", "<title", "<head", "<html", _
        "<link", "<body", "<select", "<option", "<input", _
        "<b", "<a", "<img", "<center", "<p", _
        "</meta", "<!--", "</tr", "</td", "</table", _
        "</font", "</script", "</title", "</head", "</html", _
        "</link", "</body", "</select", "</option", "</input", _
        "</b", "</a", "</img", "</center", "</p")

'---------------------------------------------
b = 1
str2 = Choose(b, ">")

'---------------------------------------------
'(1)
On Error Resume Next
    Tag_start = InStr(Tag_end, Txt, str1)
    If Tag_start = 0 Then Exit Do
    Tag_end = InStr(Tag_start, Txt, str2)
    If Tag_end = 0 Then Exit Do
    RTB1.SelStart = Tag_start - 2
    RTB1.SelLength = Tag_end - Tag_start + 2
    
'---------------------------------------------
'And the colors of the tags
Rgb1 = Choose(a, _
                 RGB(100, 0, 0), RGB(200, 0, 0), RGB(0, 100, 0), RGB(0, 200, 0), _
                 RGB(0, 0, 100), RGB(0, 0, 200), RGB(0, 255, 0), RGB(0, 255, 200), _
                 RGB(230, 255, 20), RGB(200, 255, 120), RGB(200, 255, 220), RGB(20, 205, 20), _
                 RGB(100, 155, 20), RGB(100, 55, 20), RGB(200, 155, 120), RGB(20, 1, 255), _
                 RGB(10, 15, 20), RGB(100, 5, 220), RGB(100, 15, 120), RGB(20, 100, 255), _
                 RGB(100, 0, 0), RGB(200, 0, 0), RGB(0, 100, 0), RGB(0, 200, 0), _
                 RGB(0, 0, 100), RGB(0, 0, 200), RGB(0, 255, 0), RGB(0, 255, 200), _
                 RGB(230, 255, 20), RGB(200, 255, 120), RGB(200, 255, 220), RGB(20, 205, 20), _
                 RGB(100, 155, 20), RGB(100, 55, 20), RGB(200, 155, 120), RGB(20, 1, 255), _
                 RGB(10, 15, 20), RGB(100, 5, 220), RGB(100, 15, 120), RGB(20, 100, 255), _
                 RGB(100, 0, 0), RGB(200, 0, 0), RGB(0, 100, 0), RGB(0, 200, 0), _
                 RGB(0, 0, 100), RGB(0, 0, 200), RGB(0, 255, 0), RGB(0, 255, 200), _
                 RGB(230, 255, 20), RGB(200, 255, 120), RGB(200, 255, 220), RGB(20, 205, 20), _
                 RGB(100, 155, 20), RGB(100, 55, 20), RGB(200, 155, 120), RGB(20, 1, 255), _
                 RGB(10, 15, 20), RGB(100, 5, 220), RGB(100, 15, 120), RGB(20, 100, 255), _
                 RGB(100, 0, 0), RGB(200, 0, 0), RGB(0, 100, 0), RGB(0, 200, 0), _
                 RGB(0, 0, 100), RGB(0, 0, 200), RGB(0, 255, 0), RGB(0, 255, 200), _
                 RGB(230, 255, 20), RGB(200, 255, 120), RGB(200, 255, 220), RGB(20, 205, 20), _
                 RGB(100, 155, 20), RGB(100, 55, 20), RGB(200, 155, 120), RGB(20, 1, 255), _
                 RGB(10, 15, 20), RGB(100, 5, 220), RGB(100, 15, 120), RGB(20, 100, 255))
    
    RTB1.SelColor = Rgb1
'---------------------------------------------
Loop
Next a

'---------------------------------------------
'---------------------------------------------
   
End Sub

Private Sub Form_Load()

'---------------------------------------------
'Load the htm file
'---------------------------------------------

Dim FF As Integer
Dim Txt As String
    
FF = FreeFile
Open App.Path + "\index_000.htm" For Input As FF
    Txt = Input$(LOF(FF), FF):  RTB.Text = Txt
Close FF

'---------------------------------------------
'Call the algorithm..
Call Color_Tags(RTB)
'---------------------------------------------

End Sub
Private Sub Form_Resize()
    RTB.Move 0, 0, ScaleWidth, ScaleHeight
End Sub


