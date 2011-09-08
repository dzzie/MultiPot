VERSION 5.00
Begin VB.Form frmHexEdit 
   BackColor       =   &H8000000B&
   Caption         =   "Hexedit"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   9
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   291.75
   ScaleMode       =   2  'Point
   ScaleWidth      =   474.75
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtChr 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.VScrollBar scroll 
      Height          =   5820
      Left            =   9240
      Max             =   10
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtEdit 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      MaxLength       =   2
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picDisp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FF0000&
      Height          =   5820
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   288
      ScaleMode       =   2  'Point
      ScaleWidth      =   471.75
      TabIndex        =   0
      Top             =   0
      Width           =   9495
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuOption 
         Caption         =   "Save Changes"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOption 
         Caption         =   "Save As"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOption 
         Caption         =   "JumpTo Offset"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmHexEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Copyright David Zimmer 2001 <dzzie@yahoo.com>
'Site:   http://sandsprite.com
'
'This basic hexeditor is free software and can be used in any type
'of project without license or royalty.

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const EM_GETSEL = &HB0

Dim fPath As String
Dim File() As Byte
Dim pageChanges() As String
Dim SelStart As Long
Dim SelLength As Long
Dim Dirty As Boolean


Private Sub Form_Load()
    SelStart = -1
    'loadfile App.path & "\flp.tmp"
    picDisp.FontSize = 20
    picDisp.Print "Dzzies Hexeditor v.1" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
    picDisp.FontSize = 12
    picDisp.Print "Drop File here to edit" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                  "Note: If you change the font size, face or" & vbCrLf & _
                  "picDisp Scale mode you will break everything..." & vbCrLf & vbCrLf & _
                  "so ughh...dont ! :P" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                  "pssst...Dont let this cheesy start up screen sway" & vbCrLf & _
                  "you..this is a decent framework."
                  
    picDisp.FontSize = 9 '<--really really needs to be 9 !
                         'all my math depends on all characters being 9 high and 6 wide
                         'and the hexdump format being exactly the way it is..dont say i
                         'didnt warn yaaaaa :P
    
    If Command <> Empty Then
        c = Replace(Command, """", Empty)
        If FileExists(c) Then loadfile c
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If Dirty Then If MsgBox("File been changed since last save, would you like to save it now?", vbYesNo + vbInformation) = vbYes Then SaveChanges fPath
End Sub

Private Sub mnuOption_Click(Index As Integer)
    Select Case Index
        'Case 0: SaveChanges fPath
        'Case 1: SaveChanges 'will prompt for path
        Case 2:
                a = InputBox("Enter Hex Offset to jump to, note it will be stepped to &H200 boundry")
                a = RoundUp(cHex(a), &H200)
                a = a / &H200
                If a > scroll.Max Then scroll.value = scroll.Max Else scroll.value = a
    End Select
End Sub

Private Sub picDisp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If fPath = Empty Then Exit Sub
    
    sx = SnapX(x)
    sy = snapY(Y)
    curoffset = GetOffsetFromEitherGrid(sx, sy)
    
    If Button = 1 Then               'left click =edit
         If Shift = 0 Then ResetAll  'meant for internal uses
         If sx = Empty Then Exit Sub 'click outside hexdata area
         If curoffset = -1 Then ResetAll True: Exit Sub 'at end of file
         'If X < 350 Then EditByte curoffset Else EditChar sx, sy
         SelectByte curoffset
    Else
        PopupMenu mnuPopup
    End If
    
End Sub

Private Sub picDisp_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Me.Caption = SnapX(x) & " " & snapY(y) & " " & x & " " & y
End Sub

Private Sub picDisp_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    f = data.Files(1)
    If FileExists(f) Then loadfile data.Files(1)
End Sub

Private Sub scroll_Change()
   On Error GoTo oops
    picDisp.Cls
    txtEdit.Visible = False
    txtChr.Visible = False
    sv = scroll.value 'scroll.value * &h200 = overflow if >= 64 !
    picDisp.Print HexDumpByteArray(File(), sv * &H200, &H1FF)
    ShowChanges
   Exit Sub
oops: 'yes i am a big fat cheater !
      If scroll.value = scroll.Max Then
        scroll.value = scroll.value - 1
        scroll.Max = scroll.value
      End If
End Sub

Private Sub txtchr_Change()
    i = SendMessageLong(txtChr.hwnd, EM_GETSEL, 0, 0&) \ &H10000
    If i = 1 Then
        Dirty = True
        tmp = GetGridFromOffset(SelStart) 'set when box shown
        x = txtChr.Left
        Y = snapY(txtChr.Top)
        h = Hex(Asc(txtChr))
        If Len(h) = 1 Then h = "0" & h
        RemberChange SelStart
        OverWrite x, Y, txtChr, vbYellow, vbBlack
        OverWrite tmp(0), tmp(1), h & " ", vbYellow, vbBlack
        ChangeByteFromGrid tmp(0), tmp(1), h
        tmp = GetCharGridFromOffset(SelStart + 1)
        picDisp_MouseDown 1, 1, CSng(tmp(0)), CSng(tmp(1))
    End If
End Sub

Private Sub txtEdit_Change()
    i = SendMessageLong(txtEdit.hwnd, EM_GETSEL, 0, 0&) \ &H10000
    If i = 2 Then
        Dirty = True
        x = txtEdit.Left
        Y = snapY(txtEdit.Top)
        RemberChange SelStart 'SelStart set when box shown
        tmp = GetCharGridFromOffset(SelStart)
        OverWrite tmp(0), tmp(1), GetDisplayChar(txtEdit), vbYellow, vbBlack
        OverWrite x, Y, txtEdit.Text & " ", vbYellow, vbBlack
        ChangeByteFromGrid x, Y, txtEdit
        tmp = GetGridFromOffset(SelStart + 1)
        picDisp_MouseDown 1, 1, CSng(tmp(0)), CSng(tmp(1))
    End If
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 33: 'page up
                 If scroll.value > 0 Then scroll.value = scroll.value - 1
        Case 34: 'pagedown
                 If scroll.value <> scroll.Max Then scroll.value = scroll.value + 1
        Case 38: 'uparrow
        Case 40: 'downarrow
    End Select
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    KeyAscii = FilterHexKey(KeyAscii)
    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or _
        (KeyAscii >= Asc("A") And KeyAscii <= Asc("F")) Then
        i = SendMessageLong(txtEdit.hwnd, EM_GETSEL, 0, 0&) \ &H10000
        If i > 1 Then i = 1
        txtEdit.SelStart = i
        txtEdit.SelLength = 1
    End If
End Sub

Private Sub txtchr_KeyPress(KeyAscii As Integer)
    i = SendMessageLong(txtEdit.hwnd, EM_GETSEL, 0, 0&) \ &H10000
    If i > 1 Then i = 1
    txtChr.SelStart = i
    txtChr.SelLength = 1
End Sub

Function FilterHexKey(mInkey) As Integer
    If mInkey < Asc("0") Or mInkey > Asc("9") Then
        If Not (mInkey >= Asc("A") And mInkey <= Asc("F")) Then
            If Not (mInkey >= Asc("a") And mInkey <= Asc("f")) Then
                 If mInkey <> 8 Then
                      mInkey = 0
                 End If
            End If
        End If
    End If
    If mInkey >= Asc("a") And mInkey <= Asc("f") Then
        mInkey = mInkey - 32
    End If
    FilterHexKey = mInkey
End Function

Private Sub ResetAll(Optional andSelStart = False)
    txtEdit.Visible = False: txtChr.Visible = False
    If andSelStart Then SelStart = -1 Else _
    If SelStart >= 0 Then SelectByte SelStart, False, True
End Sub

'Sub SaveChanges(Optional path = Empty)
'    Dirty = False
'    ReDim pageChanges(scroll.Max)
'    scroll_Change
'
'    If path = Empty Then
'        path = CmnDlg.ShowSave(App.path, AllFiles, "Save File As")
'        If path = Empty Then Exit Sub
'        fPath = path
'        Me.Caption = "Editing " & fPath
'    End If
'
'    f = FreeFile
'    Open path For Binary Access Write As f
'    Put f, , file()
'    Close f
'End Sub

Sub loadfile(path)
    'If Dirty Then If MsgBox("File been changed since last save, would you like to save it now?", vbYesNo + vbInformation) = vbYes Then SaveChanges fPath
    picDisp.Picture = LoadPicture()
    fPath = path
    scroll.Visible = True
    Me.Caption = "Editing " & fPath
    
    f = FreeFile
    Open path For Binary As f
        ReDim File(1 To LOF(f))
        Get f, , File()
    CloseFile CLng(f)
    
    pages = UBound(File) / &H200
    scroll.Max = IIf(InStr(pages, "."), RoundUp(pages, 1), pages - 1)
    ReDim pageChanges(scroll.Max)
    
    Me.Visible = True
    picDisp.Cls
    picDisp.Print HexDumpByteArray(File(), 0, &H1FF)
End Sub

'--------------------------------------------------------------------
'Editor Api functions
'--------------------------------------------------------------------

Function GetByteFromGrid(x, Y) As String
    'If x = Empty Then Exit Function 'used to indicate click out of bounds
    'rember file() = 1 based ! editor = 0 based !
    Dim ret As String
    offset = GetOffsetFromGrid(x, Y) + 1
    ret = Hex(File(offset))
    GetByteFromGrid = IIf(Len(ret) = 1, "0" & ret, ret)
End Function

Function GetOffsetFromEitherGrid(x, Y) As Long
    'wrapped again because i have 3 versions to correct ugh
    Dim offset As Long
    If x > 350 Then offset = CLng(GetOffsetFromCharGrid(x, Y)) _
    Else offset = GetOffsetFromGrid(x, Y)
    GetOffsetFromEitherGrid = IIf(offset < UBound(File), offset, -1)
End Function

Function GetOffsetFromGrid(x, Y) As Long
    Dim offset As Long
    a = (x - 54) / 18                '54 points (9chars) before hex data starts
    topoffset = GetTopOffset()       'what page are we viewing?
    B = topoffset + ((Y / 9) * 16)   '16 characters per line each 9 y points = one line
    offset = B + a                   'editor view is 0 based ! file() = 1 base
    GetOffsetFromGrid = offset
End Function

Function GetGridFromOffset(offset)
    topoffset = GetTopOffset()
    If offset > (topoffset + &H200) Then MsgBox "Ughh sel off page?!": Exit Function
    linesdown = (offset - topoffset) / 16
    x = ((offset Mod 16) * 18) + 54
    Y = linesdown * 9
    Dim ret()
    push ret(), SnapX(x)
    push ret(), snapY(Y)
    GetGridFromOffset = ret()
End Function

Function GetCharGridFromOffset(offset)
    topoffset = GetTopOffset()
    If offset > (topoffset + &H200) Then MsgBox "Ughh sel off page?!": Exit Function
    linesdown = (offset - topoffset) / 16
    x = ((offset Mod 16) * 6) + 360
    Y = linesdown * 9
    Dim ret()
    push ret(), x
    push ret(), snapY(Y)
    GetCharGridFromOffset = ret()
End Function

Function GetOffsetFromCharGrid(x, Y)
    topoffset = GetTopOffset()
    modulus = (x - 360) / 6
    linesdown = Y / 9
    base = topoffset + (linesdown * 16)
    GetOffsetFromCharGrid = base + modulus
End Function

Sub ChangeByteFromGrid(x, Y, hexStrNewVal)
    File(GetOffsetFromGrid(x, Y) + 1) = CByte("&H" & hexStrNewVal)
End Sub

Sub ChangeByteFromOffset(offset, hexstrValue)
    File(offset + 1) = CByte("&h" & hexstrValue)
End Sub

Sub OverWrite(x, Y, data, Optional bc = -1, Optional fc = -1)
    With frmHexEdit.picDisp
        If x = Empty Then Exit Sub
        orig = .ForeColor
         c = Array("M", "Z", "T") 'these 3 will overwrite all areas of block
         For i = 0 To 2
            .CurrentX = x
            .CurrentY = Y
            .ForeColor = IIf(bc = -1, .BackColor, bc)
            frmHexEdit.picDisp.Print String(Len(data), c(i))
         Next
        .CurrentX = x
        .CurrentY = Y
        .ForeColor = IIf(fc = -1, orig, fc)
        frmHexEdit.picDisp.Print data
        .ForeColor = orig
    End With
End Sub

Function snapY(it)
    snapY = RoundUp(it, 9)
End Function

Function GetTopOffset()
    sv = frmHexEdit.scroll.value     'bastard overflows on high offsets cause of this ! even when just multiplying and not savign to it !
    GetTopOffset = sv * &H200        'scrolls at &h200 pages
End Function

Function SnapX(it)
    Dim x As Integer 'characters are fixed at 6 points wide..hexbytes=2 +space
    x = CInt(it)     ' > 360 = char edit mode
    If x < 54 Or (x > 340 And x < 360) Then Exit Function 'x=empty = marker for some actions
    If x < 340 Then SnapX = RoundUp(x, 18) Else SnapX = RoundUp(x, 6)
End Function

Sub RemberChange(offset)
    'array has as many elements as there are pages
    'each element = comma delimited list of offsets changed
    pageChanges(frmHexEdit.scroll.value) = pageChanges(frmHexEdit.scroll.value) & offset & ","
End Sub

Sub ShowChanges()
    If pageChanges(frmHexEdit.scroll.value) = Empty Then Exit Sub
    t = Split(pageChanges(frmHexEdit.scroll.value), ",")
    For i = 0 To UBound(t) - 1
        SelectByte CLng(t(i)) 'boy is that clng necessary! through in 5min bug hunt cause of datatype dont even mumble option explicit bug hunting is half the ughh fun :P
    Next
End Sub

Sub SelectByte(offset, Optional selected = True, Optional char2 = True)
    If selected Then SelStart = offset 'Else SelStart = -1
    byteval = HexString(File(), offset + 1)
    If selected Then bc = vbYellow: fc = vbBlack Else bc = -1: fc = -1
    tmp = GetGridFromOffset(offset)
    OverWrite tmp(0), tmp(1), byteval & " ", bc, fc
    If char2 Then
        tmp = GetCharGridFromOffset(offset)
        OverWrite tmp(0), tmp(1), GetDisplayChar(byteval), bc, fc
    End If
End Sub

Sub EditByte(offset)
    If (offset + 1) > UBound(File) Then Exit Sub
    SelectByte offset
    tmp = GetGridFromOffset(offset)
    txtEdit.Move tmp(0), tmp(1) + picDisp.Top + 1, 15, 9
    txtEdit = HexString(File(), offset + 1)
    txtEdit.Visible = True
    txtEdit.SetFocus
End Sub

Sub EditChar(x, Y)
        offset = GetOffsetFromCharGrid(x, Y)
        If (offset + 1) > UBound(File) Then Exit Sub
        SelectByte offset 'this sets selstart
        txtChr.Move x, Y + picDisp.Top + 1, 6, 9
        txtChr = GetDisplayChar(HexString(File(), offset + 1))
        txtChr.Visible = True
        txtChr.SetFocus
End Sub

'-------------------------------------------------------------------
'hex editor formatting functions
'-------------------------------------------------------------------
Function HexDumpByteArray(ary() As Byte, offset, Length) As String
    Dim strArray() As String, x As Variant
    'ary = base 1 byte array offset and length assume base 0 numbers!
    Length = Length + 1 ' editor display is base 0 need base1 for array
    If offset = UBound(ary) Then MsgBox "Ughh one page to far man"
    If offset + Length > UBound(ary) Then Length = UBound(ary) - offset
    
    ReDim strArray(1 To Length)
    For i = (offset + 1) To (offset + Length)
      x = x + 1
      strArray(x) = Hex(ary(i))
      If Len(strArray(x)) = 1 Then strArray(x) = "0" & strArray(x)
    Next
    HexDumpByteArray = HexDump(strArray, offset)
End Function

Public Function HexDump(ary, ByVal offset) As String
    Dim s() As String, chars As String, tmp As String
    
    If offset > 0 And offset Mod 16 <> 0 Then MsgBox "Hexdump isnt being used right! Offset not on boundry"

    'i am lazy and simplicity rules, make sure offset read
    'starts at standard mod 16 boundry or all offsets will
    'be wrong ! it is okay to read a length that ends off
    'boundry though..that was easy to fix...
    
    chars = "   "
    For i = 1 To UBound(ary)
        tmp = tmp & ary(i) & " "
        x = CInt("&h" & ary(i))
        chars = chars & IIf((x > 32 And x < 127) Or x > 191 Or x = &H90, Chr(x), ".")
        If i > 1 And i Mod 16 = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            push s, h & "   " & tmp & chars
            offset = offset + 16:  tmp = Empty: chars = "   "
        End If
    Next
    'if read length was not mod 16=0 then
    'we have part of line to account for
    If tmp <> Empty Then
        h = Hex(offset)
        While Len(h) < 6: h = "0" & h: Wend
        h = h & "   " & tmp
        While Len(h) <= 56: h = h & " ": Wend
        push s, h & chars
    End If
    
    HexDump = Join(s, vbCrLf)
End Function

Function GetDisplayChar(hIt)
    x = CLng("&h" & hIt)
    GetDisplayChar = IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".")
End Function

Function HexString(it() As Byte, offset) As String
    Dim ret As String
    ret = Hex(it(offset))
    If Len(ret) = 1 Then ret = "0" & ret
    HexString = ret
End Function


'-------------------------------------------------------------------
' Library functions
'-------------------------------------------------------------------
Function cHex(v) As Long
    On Error Resume Next
    cHex = CLng("&h" & v)
End Function

Function FileExists(path) As Boolean
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

'ok so this one is misnamed and my math is stinkey poo poo :(
Function RoundUp(s, step)
    Dim r As Long
    r = CLng(s)
    If r Mod step <> 0 Then
        If r < step Then
            r = 0
        Else
            r = r - step
            While r Mod step <> 0: r = r + 1: Wend
        End If
    End If
    RoundUp = r
End Function
