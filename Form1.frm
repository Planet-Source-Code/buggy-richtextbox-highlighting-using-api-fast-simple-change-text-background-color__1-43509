VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "highlight selected test"
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   2880
      Width           =   2175
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4683
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'quick richtext rtf highlighting demo hack using api for psc
'bugbyter 02-2003

'ripped from vbaccelerator
'http://vbaccelerator.nuwebhost.com/codelib/richedit/richedit.htm
'watch their source code for PARAFORMAT2 (paragraph settings) and all the constants!
'...there is so much more!
'this is just a quick hack as i found its missing on PSC (and the whole web), don't have time for more...

'now found more on psc, too:
'http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=38434&lngWId=1

Option Explicit

Private Sub Command1_Click()
    Dim RTFformat As CHARFORMAT2
    With RTFformat
        .cbSize = Len(RTFformat)
        .dwMask = CFM_BACKCOLOR
        .crBackColor = vbYellow
    End With
    SendMessage RichTextBox1.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, RTFformat
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To 30
        RichTextBox1.Text = RichTextBox1.Text & "This is a highlighting demo: select some text and click the button." & vbCrLf
    Next
End Sub
