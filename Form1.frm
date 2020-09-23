VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hex Editor Pro"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "Rockwell"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1262
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":24A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":29E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3156
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "close"
            Object.ToolTipText     =   "Close"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sep1"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Edit"
            Object.ToolTipText     =   "Edit Mode"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sep2"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add"
            Object.ToolTipText     =   "Add Bytes"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Remove"
            Object.ToolTipText     =   "Remove Byte"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Insert"
            Object.ToolTipText     =   "Insert Byte"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sep3"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "search"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "goto"
            Object.ToolTipText     =   "Goto Byte"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin VB.Frame frame 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   9615
      Begin VB.Frame Frame1 
         Caption         =   "Converstions"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   7440
         TabIndex        =   20
         Top             =   3840
         Width           =   2055
         Begin VB.TextBox asciidisp 
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
            Left            =   840
            MaxLength       =   3
            TabIndex        =   24
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox hexdisp 
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
            Left            =   840
            MaxLength       =   2
            TabIndex        =   23
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox chardisp 
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
            Left            =   840
            MaxLength       =   1
            TabIndex        =   22
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox binarytxt 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   840
            MaxLength       =   8
            TabIndex        =   21
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Hex:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Ascii:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Char:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Binary:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   2640
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3135
         Left            =   120
         TabIndex        =   15
         Top             =   3840
         Width           =   1815
         Begin VB.CommandButton cmdremove 
            Caption         =   "Remove Byte"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   "Remove Byte"
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton CmdInsert 
            Caption         =   "Insert Byte"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   17
            ToolTipText     =   "Insert Byte"
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton AddBytes 
            Caption         =   "Add Bytes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   16
            ToolTipText     =   "Add Bytes"
            Top             =   2160
            Width           =   1575
         End
      End
      Begin VB.PictureBox DispTxt 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   2040
         ScaleHeight     =   3075
         ScaleWidth      =   5235
         TabIndex        =   14
         Top             =   3840
         Width           =   5295
         Begin VB.Label Showtxt 
            BackColor       =   &H00800000&
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   300
            Left            =   1920
            TabIndex        =   19
            Top             =   600
            Width           =   1020
         End
      End
      Begin VB.PictureBox HexDisplay 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3045
         Left            =   2040
         ScaleHeight     =   2985
         ScaleWidth      =   5235
         TabIndex        =   9
         Top             =   720
         Width           =   5295
         Begin VB.TextBox Edit 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   720
            MaxLength       =   2
            TabIndex        =   10
            Top             =   960
            Width           =   375
         End
      End
      Begin VB.CommandButton VTop 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         Picture         =   "Form1.frx":326E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Goto top"
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton Up10 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         Picture         =   "Form1.frx":3550
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Up 10 lines"
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton Up1 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         Picture         =   "Form1.frx":37C2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Up 1 line"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton Down1 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         Picture         =   "Form1.frx":3A34
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Down 1 line"
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton Down10 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         Picture         =   "Form1.frx":3CA6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Down 10 lines"
         Top             =   2880
         Width           =   2055
      End
      Begin VB.CommandButton Bottom 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         Picture         =   "Form1.frx":3F18
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Goto bottom"
         Top             =   3360
         Width           =   2055
      End
      Begin VB.PictureBox Position 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3060
         Left            =   120
         ScaleHeight     =   3030
         ScaleWidth      =   1785
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.PictureBox ColSet 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   5265
         TabIndex        =   1
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label ByteNo 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Size 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   7440
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Menu filemnu 
      Caption         =   "&File"
      Begin VB.Menu openmnu 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu savemnu 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu closemnu 
         Caption         =   "C&lose"
         Shortcut        =   ^L
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu exitmnu 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu editmnu 
      Caption         =   "&Edit"
      Begin VB.Menu editmodemnu 
         Caption         =   "E&dit Mode"
         Checked         =   -1  'True
         Shortcut        =   ^D
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu searchmnu 
         Caption         =   "S&earch"
         Shortcut        =   ^E
      End
      Begin VB.Menu bytemnu 
         Caption         =   "&Goto byte"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu tools 
      Caption         =   "&Tools"
      Begin VB.Menu insertbytes 
         Caption         =   "&Insert Byte"
         Shortcut        =   ^I
      End
      Begin VB.Menu rembyte 
         Caption         =   "&Remove Byte"
         Shortcut        =   ^R
      End
      Begin VB.Menu addbyte 
         Caption         =   "&Add Bytes"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu popup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu edmode 
         Caption         =   "E&dit Mode"
         Checked         =   -1  'True
      End
      Begin VB.Menu c 
         Caption         =   "-"
      End
      Begin VB.Menu insertb 
         Caption         =   "&Insert Byte"
      End
      Begin VB.Menu removeb 
         Caption         =   "&Remove Byte"
      End
      Begin VB.Menu addb 
         Caption         =   "&Add Bytes"
      End
      Begin VB.Menu e 
         Caption         =   "-"
      End
      Begin VB.Menu gos 
         Caption         =   "S&earch"
      End
      Begin VB.Menu gob 
         Caption         =   "&Goto byte"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SetMode As Boolean

Private Sub addb_Click()
AddBytes_Click
End Sub

Private Sub addbyte_Click()
AddBytes_Click
End Sub

Private Sub AddBytes_Click()
On Error Resume Next
If MsgBox("Are you sure you want to add bytes to the end of the file?", vbYesNo) = vbNo Then Exit Sub
Form4.Show
End Sub

Private Sub asciidisp_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If asciidisp > 255 Then
    hexdisp.Text = ""
    chardisp.Text = ""
    binarytxt.Text = ""
Else
    hexdisp.Text = Hex(asciidisp.Text)
    chardisp.Text = Chr(asciidisp)
    binarytxt.Text = GetBinary(hexdisp.Text)
End If
End Sub

Private Sub binarytxt_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) <> vbBack Then
    If (KeyAscii >= 48 And KeyAscii <= 49) Then
        DoEvents
    Else
        KeyAscii = 0
    End If
End If
End Sub

Private Sub binarytxt_KeyUp(KeyCode As Integer, Shift As Integer)
Dim Length As Integer, Counter As Integer, Total As Integer, no As Integer

Length = Len(binarytxt.Text)
no = 1
Total = 0
For Counter = 0 To Length - 1
If Mid(binarytxt.Text, Length - Counter, 1) = 1 Then Total = Total + no
no = no * 2
Next Counter
asciidisp.Text = Total
chardisp.Text = Chr(Total)
hexdisp.Text = Hex(Total)
End Sub

Private Sub Bottom_Click()
Dim EndBit As Integer, SetLen As String
ByteNo.Caption = ""
SetLen = SizeOfFile
Edit.Visible = False
Showtxt.Visible = False
EndBit = Mid(SetLen, Len(SetLen), 1)
CurrentPos = SizeOfFile - EndBit
SortHex
End Sub

Private Sub bytemnu_Click()
If Fileopen = False Then Exit Sub
Form2.Show
End Sub

Private Sub chardisp_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
asciidisp.Text = Asc(chardisp.Text)
hexdisp.Text = Hex(asciidisp.Text)
binarytxt.Text = GetBinary(hexdisp.Text)
End Sub

Private Sub closemnu_Click()
Dim Counter As Integer
If Fileopen = False Then Exit Sub

If MsgBox("Are you sure you want to close this file?", vbYesNo) = vbYes Then
ReDim ByteArray(0) As Byte

For Counter = 1 To 100
    HexDisplayed(Counter) = 100
Next Counter

Attributes (False)
Me.Caption = "Hex Editor Pro"
FileName = ""
sizeofile = 0
CurrentPos = 0
DispTxt.Cls
HexDisplay.Cls
ByteNo.Caption = ""
Size.Caption = ""
End If
End Sub

Private Sub CmdInsert_Click()
On Error Resume Next
If Edit.Visible = False Then MsgBox "You must select a byte first", vbExclamation: Exit Sub
If ByteNo.Caption = "" Then Exit Sub
If MsgBox("Are you sure you want to add a byte here?", vbYesNo) = vbNo Then Exit Sub
insertbyte (ByteNo.Caption)
Edit.Text = "00"
SortHex
ByteNo = ""
End Sub

Private Sub cmdremove_Click()
On Error Resume Next
If Edit.Visible = False Then MsgBox "You must select a byte first", vbExclamation: Exit Sub
If ByteNo.Caption = "" Then Exit Sub
If MsgBox("Are you sure you want to remove this byte?", vbYesNo) = vbNo Then Exit Sub
RemoveByte (ByteNo.Caption)
Edit.Visible = False
Showtxt.Visible = False
SortHex
ByteNo = ""
End Sub

Private Sub DispTxt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim Col, Row, no As Integer, HexSet As String, SAlpha As String, SetTemp As Integer

If Button = 2 And Fileopen = True Then
PopupMenu popup, , DispTxt.Left + X + 30, DispTxt.Top + Y + 380
Exit Sub
End If

Edit.Left = Int((X / HexDisplay.Width) * 10) * (HexDisplay.Width / 10)
Edit.Top = Int((Y / HexDisplay.Height) * 10) * (HexDisplay.Height / 10)

Showtxt.Left = Int((X / HexDisplay.Width) * 10) * (HexDisplay.Width / 10)
Showtxt.Top = Int((Y / HexDisplay.Height) * 10) * (HexDisplay.Height / 10)


Col = Int((X / HexDisplay.Width) * 10) + 1
Row = Int((Y / HexDisplay.Height) * 10) + 1

SetRow = Row
SetCol = Col

no = ((Row - 1) * 10) + Col


SetTemp = HexDisplayed(no)
If SetTemp = 0 Or SetTemp = 13 Or SetTemp = 10 Then
    SAlpha = " "
Else
    SAlpha = Chr(SetTemp)
End If

Showtxt.Caption = SAlpha

If Fileopen = True Then
    ByteNo.Caption = CurrentPos + (no - 1)
End If

HexSet = Hex(HexDisplayed(no))
If Len(HexSet) = 1 Then HexSet = "0" & HexSet

Edit.Visible = True
Showtxt.Visible = True

If Fileopen = False Then Edit.Visible = False: Showtxt.Visible = False
If HexSet <> "100" Then Edit.Text = HexSet Else Edit.Text = ""
End Sub

Private Sub Down1_Click()
If CurrentPos > SizeOfFile - 10 Then Exit Sub
ByteNo.Caption = ""
Edit.Visible = False
Showtxt.Visible = False
CurrentPos = CurrentPos + 10
SortHex
End Sub

Private Sub Down10_Click()
If CurrentPos > SizeOfFile - 100 Then Exit Sub
ByteNo.Caption = ""
Edit.Visible = False
Showtxt.Visible = False
CurrentPos = CurrentPos + 100
SortHex
End Sub

Private Sub Edit_KeyPress(KeyAscii As Integer)
Dim Character As String

If ByteNo.Caption > SizeOfFile Then Exit Sub
Character = Chr(KeyAscii)
KeyAscii = Asc(UCase(Character))
If Chr(KeyAscii) <> vbBack Then
    If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 70) Then
        DoEvents
    Else
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Edit_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If ByteNo.Caption > SizeOfFile Then Exit Sub
no = ((SetRow - 1) * 10) + SetCol
ByteArray((CurrentPos - 1) + no) = HexToDec(Edit.Text)
HexDisplayed(no) = HexToDec(Edit.Text)
SortHex
Showtxt.Caption = Chr(HexToDec(Edit.Text))
End Sub

Private Sub editmodemnu_Click()
Dim SetTemp As Integer, SAlpha As String
If Selected = False Then
    ByteNo.Caption = ""
    editmodemnu.Checked = True
    edmode.Checked = True
    Selected = True
    Edit.BackColor = vbYellow
    Edit.ForeColor = vbBlack
    Edit.Locked = False
    Edit.Left = 0
    Edit.Top = 0
    Edit.Text = Hex(HexDisplayed(1))
    Showtxt.Left = 0
    Showtxt.Top = 0
    SetTemp = HexDisplayed(1)
    If SetTemp = 0 Or SetTemp = 13 Or SetTemp = 10 Then
        SAlpha = " "
    Else
        SAlpha = Chr(SetTemp)
    End If
    Showtxt.Caption = SAlpha
Else
    ByteNo.Caption = ""
    editmodemnu.Checked = False
    edmode.Checked = False
    Selected = False
    Edit.Locked = True
    Edit.BackColor = &H800000
    Edit.ForeColor = vbWhite
    Edit.Left = 0
    Edit.Top = 0
    Edit.Text = Hex(HexDisplayed(1))
    Showtxt.Left = 0
    Showtxt.Top = 0
    SetTemp = HexDisplayed(1)
    If SetTemp = 0 Or SetTemp = 13 Or SetTemp = 10 Then
        SAlpha = " "
    Else
        SAlpha = Chr(SetTemp)
    End If
    Showtxt.Caption = SAlpha
End If
End Sub

Private Sub edmode_Click()
    editmodemnu_Click
End Sub

Private Sub exitmnu_Click()
Unload Me
End
End Sub

Private Sub Form_Load()
On Error Resume Next
CmdEdit.Caption = "Edit Mode"
Edit.Locked = True
Edit.BackColor = &H800000
Edit.ForeColor = vbWhite
editmodemnu.Checked = False
edmode.Checked = False
Edit.Width = HexDisplay.Width / 10
Edit.Height = HexDisplay.Height / 10
Showtxt.Width = HexDisplay.Width / 10
Showtxt.Height = HexDisplay.Height / 10
Attributes (False)

ColSet.Print " 1     2     3     4     5     6     7     8     9    10"


End Sub

Private Sub gob_Click()
bytemnu_Click
End Sub

Private Sub gos_Click()
searchmnu_Click
End Sub

Private Sub hexdisp_KeyPress(KeyAscii As Integer)
Character = Chr(KeyAscii)
KeyAscii = Asc(UCase(Character))

If Chr(KeyAscii) <> vbBack Then
    If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 70) Then
        DoEvents
    Else
        KeyAscii = 0
    End If
End If
End Sub

Private Sub hexdisp_KeyUp(KeyCode As Integer, Shift As Integer)
Dim AscStore As Integer
AscStore = HexToDec(hexdisp)
asciidisp.Text = AscStore
chardisp.Text = Chr(AscStore)
binarytxt.Text = GetBinary(hexdisp.Text)
End Sub

Private Sub HexDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) ' complete
On Error Resume Next
Dim Col, Row, no As Integer, HexSet As String, SAlpha As String, SetTemp As Integer

If Button = 2 And Fileopen = True Then
PopupMenu popup, , HexDisplay.Left + X + 30, HexDisplay.Top + Y + 380
Exit Sub
End If

Edit.Left = Int((X / HexDisplay.Width) * 10) * (HexDisplay.Width / 10)
Edit.Top = Int((Y / HexDisplay.Height) * 10) * (HexDisplay.Height / 10)

Showtxt.Left = Int((X / HexDisplay.Width) * 10) * (HexDisplay.Width / 10)
Showtxt.Top = Int((Y / HexDisplay.Height) * 10) * (HexDisplay.Height / 10)


Col = Int((X / HexDisplay.Width) * 10) + 1
Row = Int((Y / HexDisplay.Height) * 10) + 1

SetRow = Row
SetCol = Col

no = ((Row - 1) * 10) + Col


SetTemp = HexDisplayed(no)

If SetTemp = 0 Or SetTemp = 13 Or SetTemp = 10 Then
    SAlpha = " "
Else
    SAlpha = Chr(SetTemp)
End If

Showtxt.Caption = SAlpha

If Fileopen = True Then
    ByteNo.Caption = CurrentPos + (no - 1)
End If

HexSet = Hex(HexDisplayed(no))
If Len(HexSet) = 1 Then HexSet = "0" & HexSet

Edit.Visible = True
Showtxt.Visible = True

If Fileopen = False Then Edit.Visible = False: Showtxt.Visible = False
If HexSet <> "100" Then Edit.Text = HexSet Else Edit.Text = ""
End Sub

Function OpenFile()
On Error Resume Next
Dim Fno As Integer
Fno = FreeFile

Open FileName For Binary As #Fno
    SizeOfFile = LOF(Fno)
    ReDim ByteArray(1 To SizeOfFile) As Byte
    Get #Fno, , ByteArray
Close #Fno

CurrentPos = 1
StartByte = 0
Attributes (True)
Size.Caption = " " & SizeOfFile & " bytes"
Me.Caption = "Hex Editor Pro - " & FileName
Call SortHex
End Function

Function SortHex()
On Error Resume Next
Dim Counter As Integer, Counter2 As Integer, HexSet As String
Dim Line1 As String, Lines(1 To 10) As String, SAlpha As String, SetTemp As Integer
Static Pos As Integer

For Counter = 1 To 100
    If ((CurrentPos - 1) + Counter) > SizeOfFile Then
        HexDisplayed(Counter) = 256
    Else
        HexDisplayed(Counter) = ByteArray((CurrentPos - 1) + Counter)
    End If
Next Counter

For Counter = 1 To 10
    Pos = (Counter - 1) * 10
    For Counter2 = 1 To 10
        Pos = Pos + 1
        HexSet = Hex(HexDisplayed(Pos))
        If Len(HexSet) = 1 Then HexSet = "0" & HexSet
        If HexSet <> "100" Then Lines(Counter) = Lines(Counter) & HexSet & " "
    Next Counter2
Next Counter

HexDisplay.Cls
For Counter = 1 To 10
    HexDisplay.Print Lines(Counter)
Next Counter
DispTxt.Cls

For Counter = 1 To 10
Line1 = ""
    For Counter2 = 1 To 10
        SetTemp = HexDisplayed(((Counter - 1) * 10) + Counter2)
        If SetTemp < 256 Then
            If SetTemp = 0 Or SetTemp = 13 Or SetTemp = 10 Then
                SAlpha = " "
            Else
                SAlpha = Chr(SetTemp)
            End If
        Else
            SAlpha = ""
        End If
        Line1 = Line1 & "  " & SAlpha
    Next Counter2
    Line1 = Mid(Line1, 3, Len(Line1) - 2)
DispTxt.Print Line1
Next Counter

Position.Cls
For Counter = 1 To 10
    Position.Print (((Counter - 1) * 10) + (CurrentPos) - 1)
Next Counter

End Function

Private Sub insertb_Click()
CmdInsert_Click
End Sub

Private Sub insertbytes_Click()
CmdInsert_Click
End Sub

Private Sub openmnu_Click()
On Error Resume Next
Dim File As String
Edit.Visible = False
Showtxt.Visible = False
File = CommonDialog.ShowOpenDlg(Me.hwnd, "All files (*.*)" & Chr(0) & "*.*")

If File <> "Cancel" Then
    FileName = File
Else
    Exit Sub
End If

OpenFile
End Sub

Private Sub rembyte_Click()
cmdremove_Click
End Sub

Private Sub removeb_Click()
cmdremove_Click
End Sub

Private Sub savemnu_Click()
Dim Fno As Integer
If Fileopen = False Then Exit Sub
If MsgBox("Are you sure you want to save the changes?", vbYesNo) = vbYes Then
Fno = FreeFile

Open FileName For Binary As #Fno
Put #Fno, , ByteArray
Close #Fno
End If
End Sub

Private Sub searchmnu_Click()
If Fileopen = False Then Exit Sub
Form3.Show
End Sub


Private Sub textDisplay_DblClick()
MsgBox Len(textDisplay.Text)
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error Resume Next
If Button.Key = "open" Then openmnu_Click
If Button.Key = "save" Then savemnu_Click
If Button.Key = "close" Then closemnu_Click


If Fileopen = True Then
    If Button.Key = "Remove" Then cmdremove_Click
    If Button.Key = "Add" Then AddBytes_Click
    If Button.Key = "Insert" Then CmdInsert_Click
    If Button.Key = "search" Then searchmnu_Click
    If Button.Key = "goto" Then bytemnu_Click
    If Button.Key = "Edit" Then
        If Button.Image = 2 Then
            editmodemnu_Click
            Button.Image = 3
        Else
            editmodemnu_Click
            Button.Image = 2
        End If
    End If
End If
End Sub

Private Sub vTop_Click()
ByteNo.Caption = ""
CurrentPos = 1
Edit.Visible = False
Showtxt.Visible = False
SortHex
End Sub

Private Sub Up1_Click()
If CurrentPos - 10 < 1 Then vTop_Click: Exit Sub
ByteNo.Caption = ""
Edit.Visible = False
Showtxt.Visible = False
CurrentPos = CurrentPos - 10
SortHex
End Sub

Private Sub Up10_Click()
If CurrentPos - 100 < 1 Then vTop_Click: Exit Sub
ByteNo.Caption = ""
Edit.Visible = False
Showtxt.Visible = False
CurrentPos = CurrentPos - 100
SortHex
End Sub

Function HexSearch(HexVal As String, StartVal As Long) As Long
Dim ASCIIVal As Integer, Counter As Long
ASCIIVal = HexToDec(HexVal)
For Counter = StartVal To SizeOfFile
If ByteArray(Counter) = ASCIIVal Then HexSearch = Counter: Exit Function Else HexSearch = -1
Next Counter
End Function

Function SearchChars(SearchString As String, StartVal As Long) As Long
Dim Counter As Long, StrArr() As Integer, Counter2 As Integer, Check As Boolean

ReDim StrArr(1 To Len(SearchString)) As Integer
Check = True

For Counter = 1 To Len(SearchString)
StrArr(Counter) = Asc(Mid(SearchString, Counter, 1))
Next Counter

For Counter = StartVal To SizeOfFile
If ByteArray(Counter) = StrArr(1) Then

    If Len(SearchString) > 1 Then
        For Counter2 = 2 To Len(SearchString)
            If ByteArray(Counter + (Counter2 - 1)) <> StrArr(Counter2) Then
                Check = False
            End If
        Next Counter2
        If Check = True Then SearchChars = Counter: Exit Function
    Else
        SearchChars = Counter
        Exit Function
    End If
    
End If

Next Counter
SearchChars = -1
End Function

Function GetBinary(ByVal inHex As String) As String
    Dim mDec As Integer
    Dim s As String
    Dim i
    mDec = CInt("&h" & inHex)
    s = Trim(CStr(mDec Mod 2))
    i = mDec \ 2
    Do While i <> 0
        s = Trim(CStr(i Mod 2)) & s
        i = i \ 2
    Loop
    Do While Len(s) < 8
        s = "0" & s
    Loop
    GetBinary = s
End Function


Function Attributes(Value As Boolean)
Fileopen = Value
Down1.Enabled = Value
Down10.Enabled = Value
Bottom.Enabled = Value
Up1.Enabled = Value
Up10.Enabled = Value
VTop.Enabled = Value
CmdInsert.Enabled = Value
insertbytes.Enabled = Value
rembyte.Enabled = Value
addbyte.Enabled = Value
AddBytes.Enabled = Value
SearchChar = Value
cmdremove.Enabled = Value
If Value = False Then Edit.Visible = False: Showtxt.Visible = False
End Function


