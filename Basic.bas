Attribute VB_Name = "Basic"
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Global ByteArray() As Byte, FileName As String, SizeOfFile As Long
Global HexDisplayed(1 To 100) As Integer, StartByte As Long, CurrentPos As Long
Global Fileopen As Boolean, SetCol As Integer, SetRow As Integer
Global HexSearchVal As Long, CharSearchVal As Long, Selected As Boolean
Global TempArr() As Byte

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Public Function HexToDec(ByVal HexStr As String) As Double
Dim mult As Double
Dim DecNum As Double
Dim ch As String
Dim i As Integer
mult = 1
DecNum = 0
For i = Len(HexStr) To 1 Step -1
    ch = Mid(HexStr, i, 1)
    If (ch >= "0") And (ch <= "9") Then
        DecNum = DecNum + (Val(ch) * mult)
    Else
        If (ch >= "A") And (ch <= "F") Then
            DecNum = DecNum + ((Asc(ch) - Asc("A") + 10) * mult)
        Else
            If (ch >= "a") And (ch <= "f") Then
                DecNum = DecNum + ((Asc(ch) - Asc("a") + 10) * mult)
            Else
                HexToDec = 0
                Exit Function
            End If
        End If
    End If
    mult = mult * 16
Next i
HexToDec = DecNum
End Function

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hwnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub


Function insertbyte(BytePos As Long)
On Error Resume Next
Dim Counter As Long
ReDim TempArr(1 To SizeOfFile) As Byte
For Counter = 1 To SizeOfFile
    TempArr(Counter) = ByteArray(Counter)
Next Counter
SizeOfFile = SizeOfFile + 1
ReDim ByteArray(1 To SizeOfFile) As Byte
For Counter = 1 To (BytePos - 1)
    ByteArray(Counter) = TempArr(Counter)
Next Counter
ByteArray(BytePos) = 0
For Counter = (BytePos + 1) To SizeOfFile
    ByteArray(Counter) = TempArr(Counter - 1)
Next Counter
Form1.Size.Caption = " " & SizeOfFile & " bytes"
Form1.Edit.Visible = False
Form1.Showtxt.Visible = False
End Function


Function AddBytesToEnd(NoToAdd As Long)
On Error Resume Next
Dim Counter As Long, OldLength As Long
ReDim TempArr(1 To SizeOfFile) As Byte
For Counter = 1 To SizeOfFile
    TempArr(Counter) = ByteArray(Counter)
Next Counter
OldLength = SizeOfFile
SizeOfFile = SizeOfFile + NoToAdd
ReDim ByteArray(1 To SizeOfFile) As Byte
For Counter = 1 To OldLength
    ByteArray(Counter) = TempArr(Counter)
Next Counter
For Counter = (OldLength + 1) To SizeOfFile
    ByteArray(Counter) = 0
Next Counter
Form1.Size.Caption = " " & SizeOfFile & " bytes"
Form1.Edit.Visible = False
Form1.Showtxt.Visible = False
End Function

Function RemoveByte(ByteNo As Long)
On Error Resume Next
Dim Counter As Long, OldLength As Long
ReDim TempArr(1 To SizeOfFile) As Byte
For Counter = 1 To SizeOfFile
    TempArr(Counter) = ByteArray(Counter)
Next Counter
OldLength = SizeOfFile
SizeOfFile = SizeOfFile - 1
ReDim ByteArray(1 To SizeOfFile) As Byte
For Counter = 1 To ByteNo - 1
    ByteArray(Counter) = TempArr(Counter)
Next Counter
For Counter = ByteNo To SizeOfFile
    ByteArray(Counter) = TempArr(Counter + 1)
Next Counter
Form1.Size.Caption = " " & SizeOfFile & " bytes"
Form1.Edit.Visible = False
Form1.Showtxt.Visible = False
End Function
