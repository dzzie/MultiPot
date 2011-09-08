VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Emulation Honeypot  -  http://labs.iDefense.com"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   Icon            =   "MAINUI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Infected Hosts"
      Height          =   2115
      Left            =   0
      TabIndex        =   24
      Top             =   5880
      Width           =   2175
      Begin VB.ListBox lstHosts 
         Height          =   1815
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Other Settings "
      Height          =   2595
      Left            =   4560
      TabIndex        =   15
      Top             =   3240
      Width           =   2295
      Begin MSComctlLib.Slider sLogLevel 
         Height          =   255
         Left            =   840
         TabIndex        =   29
         Top             =   900
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   3
         SelStart        =   1
         Value           =   1
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   480
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   10
         Min             =   30
         Max             =   240
         SelStart        =   30
         TickFrequency   =   30
         Value           =   30
      End
      Begin VB.CheckBox chkAntiHammer 
         Caption         =   "Enable Anti-Hammer"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Tag             =   "notme"
         Top             =   240
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.Label Label3 
         Caption         =   "Log Level"
         Height          =   255
         Left            =   60
         TabIndex        =   30
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Minutes"
         Height          =   255
         Left            =   400
         TabIndex        =   26
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblHotLink 
         Caption         =   "Search Capture Records"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   120
         MouseIcon       =   "MAINUI.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   2040
         Width           =   1875
      End
      Begin VB.Label lblHotLink 
         Caption         =   "Build Capture Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   120
         MouseIcon       =   "MAINUI.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblHotLink 
         Caption         =   "Test Shellcode Handlers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   120
         MouseIcon       =   "MAINUI.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label lblHotLink 
         Caption         =   "View Logfile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         MouseIcon       =   "MAINUI.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   1560
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Server Modules "
      Height          =   2595
      Left            =   0
      TabIndex        =   6
      Top             =   3240
      Width           =   2175
      Begin VB.CheckBox chkVeritas 
         Caption         =   "Veritas -  6101 && 10000"
         Height          =   255
         Left            =   60
         TabIndex        =   28
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CheckBox chkMyDoom 
         Caption         =   "My Doom - port 3127"
         Height          =   315
         Left            =   60
         TabIndex        =   13
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkOptix 
         Caption         =   "Optix - port 2060 && 500 "
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   600
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkBagle 
         Caption         =   "Bagle port 2745 && 12345"
         Height          =   315
         Left            =   60
         TabIndex        =   11
         Top             =   900
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkSub7 
         Caption         =   "Sub 7 - port 27347"
         Height          =   375
         Left            =   60
         TabIndex        =   10
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1755
      End
      Begin VB.CheckBox chkKuang 
         Caption         =   "Kuang2 -port 17300"
         Height          =   315
         Left            =   60
         TabIndex        =   9
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkLsass 
         Caption         =   "RPC445"
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   2220
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.TextBox txtLsassPort 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Text            =   "445"
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label lblRCP445 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1080
         TabIndex        =   31
         Top             =   2175
         Width           =   195
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Banned IPs"
      Height          =   2595
      Left            =   2280
      TabIndex        =   4
      Top             =   3240
      Width           =   2235
      Begin VB.CheckBox chkIgnoreList 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Tag             =   "notme"
         Top             =   240
         Width           =   915
      End
      Begin VB.ListBox List2 
         Height          =   2010
         Left            =   180
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblHotLink 
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   1500
         MouseIcon       =   "MAINUI.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblHotLink 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   1140
         MouseIcon       =   "MAINUI.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   2220
      TabIndex        =   3
      Top             =   5940
      Width           =   9315
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10320
      TabIndex        =   1
      Top             =   5460
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8940
      TabIndex        =   0
      ToolTipText     =   "Important -> Spam"
      Top             =   5460
      Width           =   1215
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3255
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "File Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Remote IP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Honeypot"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "MD5"
         Text            =   "MD5"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1725
      Left            =   6900
      Picture         =   "MAINUI.frx":1546
      Stretch         =   -1  'True
      Top             =   3300
      Width           =   4635
   End
   Begin VB.Label Label1 
      Caption         =   "^(-_-)^"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7800
      TabIndex        =   14
      Top             =   4140
      Width           =   1695
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuChangeName 
         Caption         =   "Change Name"
      End
      Begin VB.Menu mnuDeleteSize 
         Caption         =   "Delete Size"
      End
      Begin VB.Menu mnuSpacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHexEdit 
         Caption         =   "Hexedit"
      End
      Begin VB.Menu mnuWhois 
         Caption         =   "Whois"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy All"
      End
      Begin VB.Menu mnuCopyFileName 
         Caption         =   "Copy FileName"
      End
      Begin VB.Menu mnuCopyfiletoDesktop 
         Caption         =   "Copy file to Desktop"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearList 
         Caption         =   "Clear List"
      End
   End
   Begin VB.Menu mnuPopup2 
      Caption         =   "mnuPopup2"
      Visible         =   0   'False
      Begin VB.Menu mnuLowerList 
         Caption         =   "Clear"
         Index           =   0
      End
      Begin VB.Menu mnuLowerList 
         Caption         =   "Copy All"
         Index           =   1
      End
      Begin VB.Menu mnuLowerList 
         Caption         =   "Copy Line"
         Index           =   2
      End
   End
   Begin VB.Menu mnuPopup3 
      Caption         =   "mnuPOpup3"
      Visible         =   0   'False
      Begin VB.Menu mnuUniqueHosts 
         Caption         =   "Copy"
         Index           =   0
      End
      Begin VB.Menu mnuUniqueHosts 
         Caption         =   "Copy All"
         Index           =   1
      End
      Begin VB.Menu mnuUniqueHosts 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuUniqueHosts 
         Caption         =   "Ban"
         Index           =   3
      End
      Begin VB.Menu mnuUniqueHosts 
         Caption         =   "Clear"
         Index           =   4
      End
      Begin VB.Menu mnuUniqueHosts 
         Caption         =   "Whois"
         Index           =   5
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Columns: time, size, filename, ip, name, honeypot
'tblData: autoid  starttime   ip  fname   fsize   md5 honeypot


'Author: david@idefense.com
'
'License: Copyright (C) 2005 David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA

'Stuff to do/cleanup

'Right now i am also duplicating the shellcode handlers TryToProcess() function
'individually inside each handlers class file. I am keeping them seperate atm
'because i need to wait a while and see if they need to specialized much or if
'it makes sense to conglomerate them. Originally they were in a much better design
'framework, but I had to butcher it up because i needed the handlers checks integrated
'tighter into the code because of minimal time windows to react.
'
'The veritas 10000 module has had any hits yet so not sure how it will deal.

'ChangeLog:
'--------------------------------------------------------------------
' 7.17.05
'           added 3mb max download limit to ftp and tftp classes
'           fixed bug in CGenericURL.TryToHandle - ftp downloads would not display in list bad file name
'
' 8.13.05
'           changed lsass to be rpc445,
'           added support for determining GUID from DCE Bind packets
'           rpc445 now sub catagorizes received exploits based on GUID, (ASN shows up as generic RPC445)
'           added shellcode handler for a PNP cmd shell payload.
'
' 8.17.05
'           added 2nd pnp cmd shellcode handler
'           fixed display bug with antihammer times
'
'todo: conglomerat all cmd shellcode handlers into one that accepts multiple
'      signatures, and only has seperate small functions to extract port/ip data from
'      payload, all repetive code in those 3 cmd handlers atm

Private WithEvents optix As clsOptix
Attribute optix.VB_VarHelpID = -1
Private WithEvents MyDoom As clsMyDoom
Attribute MyDoom.VB_VarHelpID = -1
Private WithEvents bagle As clsBagle
Attribute bagle.VB_VarHelpID = -1
Private WithEvents sub7 As clsSub7
Attribute sub7.VB_VarHelpID = -1
Private WithEvents kuang As clsKuang2
Attribute kuang.VB_VarHelpID = -1
Private WithEvents rpc445 As CRpc445
Attribute rpc445.VB_VarHelpID = -1
Private WithEvents veritas As CVeritas
Attribute veritas.VB_VarHelpID = -1
Private WithEvents veritas2 As CVeritas_II
Attribute veritas2.VB_VarHelpID = -1
'Private WithEvents dcom As clsDCOM

Dim selli As ListItem

Dim UniqueHosts As New Collection

Sub AddUniqueHost(ip As String)
    On Error GoTo hell
    If KeyExistsInCollection(UniqueHosts, ip) Then Exit Sub
    UniqueHosts.Add ip, ip
    lstHosts.AddItem ip
hell:
End Sub

Private Sub chkAntiHammer_Click()
    Dim setting As Boolean
    setting = CBool(chkAntiHammer.value)
    optix.hammer.Enabled = setting
    MyDoom.hammer.Enabled = setting
    bagle.hammer.Enabled = setting
    sub7.hammer.Enabled = setting
    kuang.hammer.Enabled = setting
    rpc445.hammer.Enabled = setting
    veritas.hammer.Enabled = setting
    veritas2.hammer.Enabled = setting
   'dcom.hammer.Enabled = setting
End Sub
 



Private Sub lblRCP445_Click()

    Const msg = "RPC 445 emulates a crude RPC server typically found on port 445\n\n" & _
                "This module can capture several exploits including:\n" & _
                "LSASS, ASN.1, and the new PNP vuln."
    
    MsgBox Replace(msg, "\n", vbCrLf)
    
End Sub

 

Private Sub Slider1_Change()
    Dim setting As Long
    On Error Resume Next
    setting = Slider1.value
    optix.hammer.MinutesBlocked = setting
    MyDoom.hammer.MinutesBlocked = setting
    bagle.hammer.MinutesBlocked = setting
    sub7.hammer.MinutesBlocked = setting
    kuang.hammer.MinutesBlocked = setting
    rpc445.hammer.MinutesBlocked = setting
    veritas.hammer.MinutesBlocked = setting
    veritas2.hammer.MinutesBlocked = setting
    'dcom.hammer.MinutesBlocked = setting
End Sub


Private Sub lblHotLink_Click(Index As Integer)
    
    Dim ip As String
    On Error Resume Next
    
    Select Case Index
        Case 0: Shell "notepad """ & logFile & """", vbNormalFocus
        Case 1: frmScTest.Show
        Case 2: frmStats.DoReport
        Case 3: frmSearch.Show
        
        Case 4:
                ip = InputBox("Enter IP to block")
                If Len(ip) = 0 Then Exit Sub
                List2.AddItem ip
                
        Case 5:
                If List2.SelCount < 1 Then Exit Sub
                List2.RemoveItem GetSelIndex(List2)
                
    End Select
    
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup2
End Sub

Private Sub lstHosts_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup3
End Sub

Private Sub mnuCopyfiletoDesktop_Click()
    On Error Resume Next
    If selli Is Nothing Then Exit Sub
    fso.Copy selli.SubItems(2), UserDeskTopFolder
End Sub

Private Sub mnuHexEdit_Click()
    
    If selli Is Nothing Then Exit Sub
    
    On Error GoTo hell
    Dim D As frmHexEdit
    Set D = New frmHexEdit
    
    D.loadfile selli.SubItems(2)
    
    Exit Sub
hell: Unload D
    
End Sub

Private Sub mnuDeleteSize_Click()
    
    If selli Is Nothing Then Exit Sub
    
    Dim sql As String
    Dim size As Long
    Dim li As ListItem
    
    On Error GoTo hell
    
    size = selli.SubItems(1)
    
    If MsgBox("Are you sure you want to remove the size entry for filesize: " & size, vbYesNo) = vbYes Then
        ado.Execute "Delete from knownsizes where fsize=" & size
           
nextone:
        For Each li In lv.ListItems
            If li.SubItems(1) = size Then
                lv.ListItems.remove li.Index
                GoTo nextone
            End If
        Next
           
        'MsgBox "Entry Removed..", vbInformation
        Set selli = Nothing
        
    End If
    
    Exit Sub
hell: MsgBox Err.Description
End Sub


Private Sub mnuChangeName_Click()
    If selli Is Nothing Then Exit Sub
    ChangeNameForSize selli, lv, 1, 4
    Set selli = Nothing
End Sub

Private Sub mnuLowerList_Click(Index As Integer)
    
    On Error Resume Next
    Dim tmp() As String
    Dim i As Integer
                  
    Select Case Index
        Case 0:  List1.Clear
        Case 1:
        
                For i = 0 To List1.ListCount
                    push tmp, List1.list(i)
                Next
                Clipboard.Clear
                Clipboard.SetText Join(tmp, vbCrLf)
        Case 2:
                Clipboard.Clear
                Clipboard.SetText List1.list(List1.ListIndex)
    End Select
    
End Sub

 

Private Sub mnuUniqueHosts_Click(Index As Integer)
    'copy, copyall, - , ban, clear
    On Error GoTo hell
    Dim i As Long, tmp() As String
    
    With lstHosts
        Select Case Index
            Case 0:
                    Clipboard.Clear
                    Clipboard.SetText .list(.ListIndex)
            Case 1:
                    For i = 0 To .ListCount
                        push tmp, .list(i)
                    Next
                    Clipboard.Clear
                    Clipboard.SetText Join(tmp, vbCrLf)
            Case 2:
            Case 3: List2.AddItem .list(.ListIndex)
            Case 4: .Clear
            Case 5: RunWhois .list(.ListIndex)
        End Select
    End With
            
            
            
    Exit Sub
hell:    MsgBox Err.Description
End Sub

Private Sub mnuWhois_Click()
    If selli Is Nothing Then Exit Sub
    RunWhois CStr(selli.SubItems(3))
End Sub




Sub AddtoList(fPath, ip, from As dumpDirs, Optional subSploit As SubSploits)

    On Error Resume Next
    
    Dim li As ListItem
    Dim rs As Recordset
    Dim md5 As String
    Dim fsize As Long
    Dim tmp As String
    Dim sql As String
    Dim wasNew As Boolean
    Dim unkName As String
    
    If Len(fPath) = 0 Then Exit Sub
    
    
    AddUniqueHost CStr(ip) 'covers full uploads only, lsass shellcode
                           'handler may have caused a failure so we have
                           'to do this check somewhere else too...humm
    
    fsize = FileLen(fPath)
    md5 = hash.HashFile(CStr(fPath))
          
    If Len(md5) = 0 Then md5 = hash.error_message
          
    Set li = lv.ListItems.Add
    li.Text = Now
    li.SubItems(1) = fsize
    li.SubItems(2) = fPath
    li.SubItems(3) = ip
    
    sql = "Select * from knownsizes where fsize=" & fsize
    
    Set rs = ado(sql)
    
    If rs.BOF And rs.EOF Then
              
        wasNew = True
        
        If from = eRPC445 And subSploit <> ssUnk Then
            unkName = "Unk_" & SSNameFromEnum(subSploit) & "_" & fsize
        Else
            unkName = "Unk_" & HPNameFromEnum(from) & "_" & fsize
        End If
        
        li.SubItems(4) = "-> " & unkName
        
        ado.CloseConnection
        ado.Insert "knownsizes", "malcode,fsize,md5", unkName, fsize, md5
        
    Else
        li.SubItems(4) = rs!malcode
        
        If InStr(1, fPath, "ftp:") > 0 Then 'http:// or ftp:// (bagle)
            DoEvents
        Else
            If fso.FileExists(CStr(fPath)) Then 'already have a sample of this
                tmp = archiveDir & fso.FileNameFromPath(CStr(fPath))
               
                If fso.FileExists(tmp) Then 'rare
                    fso.DeleteFile CStr(fPath)
                    fPath = "deleted due to name conflict - " & tmp & " org: " & fPath
                Else
                    fPath = fso.Move(CStr(fPath), archiveDir)
                End If
                
                li.SubItems(2) = fPath
            End If
        End If
        
    End If
    
    If from = eRPC445 And subSploit <> ssUnk Then
        li.SubItems(5) = SSNameFromEnum(subSploit)
    Else
        li.SubItems(5) = HPNameFromEnum(from)
    End If
    
    li.SubItems(6) = md5
    
    ado.CloseConnection
    
    ado.Insert "tbldata", "starttime,ip,fname,fsize,md5,honeypot", _
                           Now, ip, fPath, fsize, md5, li.SubItems(5)
    
    ado.CloseConnection
    
End Sub



 


Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuClearList_Click()
    Set selli = Nothing
    lv.ListItems.Clear
End Sub

Private Sub mnuCopyAll_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText GetAllElements(lv)
    MsgBox "Copy complete", vbInformation
End Sub

Private Sub mnuCopyFileName_Click()
    If selli Is Nothing Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText selli.SubItems(2)
End Sub

Sub SetChecks(Enabled As Boolean)
    Dim c As Control
    On Error Resume Next
    For Each c In Me.Controls
        If TypeName(c) = "CheckBox" Then
            If c.Tag <> "notme" Then c.Enabled = Enabled
        End If
    Next
End Sub


Private Sub cmdStart_Click()
    
    On Error Resume Next
    
    dbg llreal, "Started..."
    SetChecks False
    cmdStart.Enabled = False
    chkAntiHammer_Click
    Slider1_Change
    
    If chkOptix.value = 1 Then optix.StartUp
    If chkMyDoom.value = 1 Then MyDoom.StartUp
    If chkBagle.value = 1 Then bagle.StartUp: BagleFtp.StartUp
    If chkSub7.value = 1 Then sub7.StartUp
    If chkKuang.value = 1 Then kuang.StartUp
    If chkVeritas.value = 1 Then veritas.StartUp: veritas2.StartUp
    
    If chkLsass.value = 1 Then
        If IsNumeric(txtLsassPort) Then
            If txtLsassPort = 135 Then
                MsgBox "Lsass listener is not designed to work on port 135 via RPC please see documentation for more details", vbInformation
                chkLsass.value = 0
                Exit Sub
            End If
            rpc445.server.port = CLng(txtLsassPort)
            txtLsassPort.Enabled = False
            rpc445.StartUp
        Else
            MsgBox "Could not start lsass listener invalid port specified", vbInformation
            chkLsass.value = 0
        End If
    End If
    
'    If chkDCOM.value = 1 Then
'        If IsNumeric(txtDcom) Then
'            If txtDcom = 445 Then
'                MsgBox "Dcom listener is not designed to work on port 445 via RPC please see documentation for more details", vbInformation
'                chkDCOM.value = 0
'                Exit Sub
'            End If
'            'dcom.server.port = CLng(txtDcom)
'            txtDcom.Enabled = False
'            dcom.StartUp
'        Else
'            MsgBox "Could not start Dcom listener invalid port specified", vbInformation
'            chkDCOM.value = 0
'        End If
'    End If
    
    
End Sub

Private Sub cmdStop_Click()
    dbg llreal, "Stopped..."
    SetChecks True
    cmdStart.Enabled = True
    txtLsassPort.Enabled = True
    'txtDcom.Enabled = True
    optix.ShutDown
    MyDoom.ShutDown
    bagle.ShutDown
    BagleFtp.ShutDown
    sub7.ShutDown
    kuang.ShutDown
    rpc445.ShutDown
    veritas.ShutDown
    veritas2.ShutDown
    'dcom.ShutDown
End Sub

 

Private Sub Form_Load()
    InitDumpDirs
    LoadConfig
    
    lv.ColumnHeaders(lv.ColumnHeaders.Count).Width = lv.Width - lv.ColumnHeaders(lv.ColumnHeaders.Count).Left - 100
    
    logFile = App.path & "\log.txt"
    
    Const dbpath As String = "C:\honeypot\honeypot.mdb"
    
    ado.BuildConnectionString Access, dbpath
    
    With config
        chkAntiHammer.value = .AntiHammer
        chkBagle = .bagle
        chkKuang = .kuang
        chkLsass = .lsass
        chkMyDoom = .doom
        chkOptix = .optix
        chkIgnoreList = .banips
        chkSub7 = .sub7
        chkVeritas = .veritas
        'chkDCOM = .dcom
        txtLsassPort = .lsass_port
        'txtDcom = .dcom_port
        If .hammer_time < 30 Or .hammer_time > 240 Then .hammer_time = 60
        If .logLevel < 1 Or .logLevel > 3 Then .logLevel = 3
        Slider1.value = .hammer_time
        sLogLevel.value = .logLevel
    End With
    
    If Not fso.FileExists(dbpath) Then
        MsgBox "Could not find database: " & dbpath
        cmdStart.Enabled = False
        cmdStop.Enabled = False
    End If
    
    Set optix = New clsOptix
    Set MyDoom = New clsMyDoom
    Set bagle = New clsBagle
    Set sub7 = New clsSub7
    Set BagleFtp = New clsBagleFtpRecv
    Set kuang = New clsKuang2
    Set rpc445 = New CRpc445
    Set veritas = New CVeritas
    Set veritas2 = New CVeritas_II
    'Set dcom = New clsDCOM
    
    Dim blockedIps, i As Long
    
    blockedIps = App.path & "\blocked.txt"
    
    If fso.FileExists(CStr(blockedIps)) Then
        blockedIps = Split(fso.ReadFile(blockedIps), vbCrLf)
        For i = 0 To UBound(blockedIps)
            blockedIps(i) = Trim(blockedIps(i))
            If Len(blockedIps(i)) > 0 Then List2.AddItem blockedIps(i)
        Next
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
   
    Dim blockedIps, i As Long
    Dim dat() As String
    
    If Not cmdStart.Enabled Then cmdStop_Click
    
    blockedIps = App.path & "\blocked.txt"
    
    For i = 0 To List2.ListCount
         push dat, List2.list(i)
    Next
    
    fso.WriteFile CStr(blockedIps), Join(dat, vbCrLf)
    
    With config
        .AntiHammer = chkAntiHammer.value
        .bagle = chkBagle
        .kuang = chkKuang
        .lsass = chkLsass
        .doom = chkMyDoom
        .optix = chkOptix
        .banips = chkIgnoreList
        .sub7 = chkSub7
        .veritas = chkVeritas
        .lsass_port = CInt(txtLsassPort)
        '.dcom_port = CInt(txtDcom)
        .hammer_time = Slider1.value
        .logLevel = sLogLevel.value
        '.dcom = chkDCOM
    End With
    
    SaveConfig
    
    Dim f As Form
    For Each f In Forms
        Unload f
    Next
    
    End
    
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selli = Item
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.Width = 11720
    Me.Height = 8400
End Sub

Private Sub sLogLevel_Change()
    LogThreshHold = sLogLevel.value
End Sub


' Connection requests
'-------------------------------------------------------------------------
Private Sub MyDoom_Connection(ip As String)
     dbg llinfo, "MyDoom Connection: " & ip
End Sub
Private Sub Optix_Connection(ip As String)
    dbg llinfo, "Optix Connection: " & ip
End Sub
Private Sub Sub7_Connection(ip As String)
    dbg llinfo, "Sub7 Connection: " & ip
End Sub
Private Sub Bagle_Connection(ip As String)
    dbg llinfo, "Bagle 1 Connection: " & ip
End Sub
Private Sub Kuang_Connection(ip As String)
    dbg llinfo, "Kuang Connection: " & ip
End Sub

Private Sub MyDoom_ConnectionRequest(ip As String, Block As Boolean)
    If chkIgnoreList.value = 0 Then Exit Sub
    If IPinList(ip, List2) Then
        Block = True
        dbg llspam, "Blocked MyDoom ip " & ip & "  ( " & Now & " )"
    End If
End Sub

Private Sub Optix_ConnectionRequest(ip As String, Block As Boolean)
    If chkIgnoreList.value = 0 Then Exit Sub
    If IPinList(ip, List2) Then
        Block = True
        dbg llspam, "Blocked Optix ip " & ip & "  ( " & Now & " )"
    End If
End Sub

Private Sub Sub7_ConnectionRequest(ip As String, Block As Boolean)
    If chkIgnoreList.value = 0 Then Exit Sub
    If IPinList(ip, List2) Then
        Block = True
        dbg llspam, "Blocked Sub7 ip " & ip & "  ( " & Now & " )"
    End If
End Sub

Private Sub Kuang_ConnectionRequest(ip As String, Block As Boolean)
    If chkIgnoreList.value = 0 Then Exit Sub
    If IPinList(ip, List2) Then
        Block = True
        dbg llspam, "Blocked Kuang2 ip " & ip & "  ( " & Now & " )"
    End If
End Sub

Private Sub Bagle_ConnectionRequest(ip As String, Block As Boolean)
    If chkIgnoreList.value = 0 Then Exit Sub
    If IPinList(ip, List2) Then
        Block = True
        dbg llspam, "Blocked Bagle ip " & ip & "  ( " & Now & " )"
    End If
End Sub

'Private Sub dcom_Connection(ip As String, Block As Boolean)
'    dbg llspam, "Dcom Connection: " & ip
'    If chkIgnoreList.value = 0 Then Exit Sub
'    If IPinList(ip, List2) Then
'        Block = True
'        dbg llspam, "Blocked Dcom ip " & ip & "  ( " & Now & " )"
'    End If
'End Sub
Private Sub Rpc445_Connection(ip As String, Block As Boolean)
    If chkIgnoreList.value = 0 Then Exit Sub
    If IPinList(ip, List2) Then
        Block = True
        dbg llspam, "Blocked Rpc445 ip " & ip & "  ( " & Now & " )"
    End If
End Sub
Private Sub veritas_Connection(ip As String, Block As Boolean)
    If chkIgnoreList.value = 0 Then Exit Sub
    If IPinList(ip, List2) Then
        Block = True
        dbg llspam, "Blocked Veritas ip " & ip & "  ( " & Now & " )"
    End If
End Sub
Private Sub veritas2_Connection(ip As String, Block As Boolean)
    If chkIgnoreList.value = 0 Then Exit Sub
    If IPinList(ip, List2) Then
        Block = True
        dbg llspam, "Blocked Veritas2 ip " & ip & "  ( " & Now & " )"
    End If
End Sub


Private Sub MyDoom_BadUpload(ip As String, fPath As String)
     dbg llspam, "MyDoom Failed Upload: " & ip & " fpath:" & fPath
End Sub

' timeouts
'------------------------------------------------------

Private Sub Optix_TimedOut(ip As String)
    dbg llspam, "Optix Timeout: " & ip
End Sub
Private Sub Sub7_TimedOut(ip As String)
    dbg llspam, "Sub7 Timeout: " & ip
End Sub
Private Sub Bagle_TimedOut(ip As String)
    dbg llspam, "Bagle 1 Timeout: " & ip
End Sub
Private Sub Kuang_TimedOut(ip As String)
    dbg llspam, "Kuang Timeout: " & ip
End Sub
Private Sub MyDoom_TimedOut(ip As String)
    dbg llspam, "MyDoom Timeout: " & ip
End Sub
Private Sub Rpc445_TimedOut(ip As String)
    dbg llspam, "Rpc445 Timeout: " & ip
End Sub
Private Sub veritas_TimedOut(ip As String)
    dbg llspam, "Veritas 6101 Timeout: " & ip
End Sub
Private Sub veritas2_TimedOut(ip As String)
    dbg llspam, "Veritas 10000 Timeout: " & ip
End Sub
'Private Sub dcom_TimedOut(ip As String)
'    dbg llspam, "Dcom Timeout: " & ip
'End Sub

'-----------------------------------------------------------------------

Private Sub Sub7_UploadComplete(ip As String, fPath As String)
    dbg llspam, "Sub7 Upload Complete: " & ip & " fpath:" & fPath
    AddtoList fPath, ip, esub7
End Sub

Private Sub Optix_UploadComplete(ip As String, fPath As String)
    dbg llspam, "Optix Upload Complete: " & ip & " File: " & fPath
    AddtoList fPath, ip, eOptix
End Sub

Private Sub MyDoom_UploadComplete(ip As String, fPath As String)
     dbg llspam, "Upload Complete: " & ip & " fpath: " & fPath
     AddtoList fPath, ip, eMyDoom
End Sub

Private Sub Bagle_UploadComplete(ip As String, fPath As String)
     dbg llspam, "Bagle 1 Upload Complete: " & ip & " DL-Location: " & fPath
     AddtoList fPath, ip, eBagle
End Sub

Private Sub Kuang_UploadComplete(ip As String, fPath As String)
    AddtoList fPath, ip, eKuang
    dbg llspam, "Upload Complete: " & ip & " " & fPath
End Sub

Private Sub Rpc445_UploadComplete(ip As String, fPath As String, GUID As String, stage As Integer)
    If Not fso.FileExists(fPath) Then Exit Sub
    
    On Error Resume Next
    AddtoList fPath, ip, eRPC445, SubSploitFromIID(GUID)
    Err.Clear
    
    dbg llspam, "Upload Complete: " & ip & " Stage: " & stage & " " & fPath
End Sub

Private Sub veritas_UploadComplete(ip As String, fPath As String)
    If Not fso.FileExists(fPath) Then Exit Sub
    AddtoList fPath, ip, everitas
    dbg llspam, "Veritas 6101 Upload Complete: " & ip & " " & fPath
End Sub

Private Sub veritas2_UploadComplete(ip As String, fPath As String)
    If Not fso.FileExists(fPath) Then Exit Sub
    AddtoList fPath, ip, everitas
    dbg llspam, "Veritas 10000 Upload Complete: " & ip & " " & fPath
End Sub
'Private Sub dcom_UploadComplete(ip As String, fPath As String)
'    If Not fso.FileExists(fPath) Then Exit Sub
'    AddtoList fPath, ip, EDCOM
'    dbg llspam, "Dcom Upload Complete: " & ip & " " & fPath
'End Sub



'___________________________________________________info
Private Sub Rpc445_info(msg As String)
    dbg llinfo, "Rpc445 info: " & msg
End Sub
Private Sub veritas_info(msg As String)
    dbg llinfo, "Veritas 6101 info: " & msg
End Sub
Private Sub veritas2_info(msg As String)
    dbg llinfo, "Veritas 10000 info: " & msg
End Sub
'Private Sub dcom_Info(msg As String)
'    dbg llinfo, "Dcom info: " & msg
'End Sub

'_____________________________________AntiHammer Messages

Private Sub Rpc445_AntiHammer(ByVal ip As String, blockedUntil As String, remove As Boolean)
    dbg llspam, "Rpc445 Antihammer: " & ip & " blocked until: " & Format(blockedUntil, "h:n:ss")
End Sub
Private Sub kuang_AntiHammer(ByVal ip As String, blockedUntil As String, remove As Boolean)
    dbg llspam, "Kuang Antihammer: " & ip & " blocked until: " & Format(blockedUntil, "h:n:ss")
End Sub
Private Sub bagle_AntiHammer(ByVal ip As String, blockedUntil As String, remove As Boolean)
    dbg llspam, "Bagle Antihammer: " & ip & " blocked until: " & Format(blockedUntil, "h:n:ss")
End Sub
Private Sub mydoom_AntiHammer(ByVal ip As String, blockedUntil As String, remove As Boolean)
    dbg llspam, "MyDoom Antihammer: " & ip & " blocked until: " & Format(blockedUntil, "h:n:ss")
End Sub
Private Sub optix_AntiHammer(ByVal ip As String, blockedUntil As String, remove As Boolean)
    dbg llspam, "Optix Antihammer: " & ip & " blocked until: " & Format(blockedUntil, "h:n:ss")
End Sub
Private Sub sub7_AntiHammer(ByVal ip As String, blockedUntil As String, remove As Boolean)
    dbg llspam, "Sub7 Antihammer: " & ip & " blocked until: " & Format(blockedUntil, "h:n:ss")
End Sub
Private Sub veritas_AntiHammer(ByVal ip As String, blockedUntil As String, remove As Boolean)
    dbg llspam, "Veritas 6101 Antihammer: " & ip & " blocked until: " & Format(blockedUntil, "h:n:ss")
End Sub
Private Sub veritas2_AntiHammer(ByVal ip As String, blockedUntil As String, remove As Boolean)
    dbg llspam, "Veritas 10000 Antihammer: " & ip & " blocked until: " & Format(blockedUntil, "h:n:ss")
End Sub
'Private Sub dcom_AntiHammer(ByVal ip As String, blockedUntil As String, remove As Boolean)
'    dbg llspam, "DCOM Antihammer: " & ip & " blocked unti: " & blockedUntil
'End Sub


'_________________________________________Recgonized exploits
Private Sub Rpc445_RecgonizedExploit(ByVal ip As String, ByVal HandlerName As String)
    'just a place to do the unique host thing that will process everytime..
    AddUniqueHost ip
End Sub
Private Sub veritas_RecgonizedExploit(ByVal ip As String, ByVal HandlerName As String)
    'just a place to do the unique host thing that will process everytime..
    AddUniqueHost ip
End Sub
Private Sub veritas2_RecgonizedExploit(ByVal ip As String, ByVal HandlerName As String)
    'just a place to do the unique host thing that will process everytime..
    AddUniqueHost ip
End Sub
