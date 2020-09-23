VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Searchtxt 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3735
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Search For String"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Search For Hex"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Val As Long, Temp As String

If Option1.Value = True Then
    If Searchtxt.Text = "" Then Exit Sub
    Me.Caption = "Searching..."
    Val = Form1.HexSearch(Searchtxt.Text, HexSearchVal)
    Me.Caption = "Search"
    If Val = -1 Then Exit Sub
    HexSearchVal = Val + 1
    CurrentPos = Val
    Form1.Edit.Visible = False
    Form1.SortHex
    Temp = Hex(HexDisplayed(1))
    If Len(Temp) = 1 Then Temp = "0" & Temp
    Form1.Edit.Left = 0
    Form1.Edit.Top = 0
    Form1.Edit.Text = Temp
    Form1.Edit.Visible = True
Else
    If Searchtxt.Text = "" Then Exit Sub
    Me.Caption = "Searching..."
    Val = Form1.SearchChars(Searchtxt.Text, CharSearchVal)
    Me.Caption = "Search"
    If Val = -1 Then Exit Sub
    CharSearchVal = Val + 1
    CurrentPos = Val
    Form1.Edit.Visible = False
    Form1.SortHex
    Temp = Hex(HexDisplayed(1))
    If Len(Temp) = 1 Then Temp = "0" & Temp
    Form1.Edit.Left = 0
    Form1.Edit.Top = 0
    Form1.Edit.Visible = True
    Form1.Edit.Text = Temp
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Searchtxt.MaxLength = 2
AlwaysOnTop Me, True
End Sub

Private Sub Option1_Click()
Searchtxt.MaxLength = 2
Searchtxt.Text = ""
End Sub

Private Sub Option2_Click()
Searchtxt.MaxLength = 8
Searchtxt.Text = ""
End Sub

Private Sub Searchtxt_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim Character As String
If Option1.Value = True Then
Character = Chr(KeyAscii)
KeyAscii = Asc(UCase(Character))
If Chr(KeyAscii) <> vbBack Then
    If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 70) Then
        DoEvents
    Else
        KeyAscii = 0
    End If
End If
HexSearchVal = 1
End If
CharSearchVal = 1
End Sub
