VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9825
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   9825
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "IP"
      Height          =   255
      Index           =   3
      Left            =   3300
      TabIndex        =   8
      Top             =   4620
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy"
      Height          =   375
      Left            =   8520
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   4560
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4260
      TabIndex        =   4
      Top             =   4560
      Width           =   2595
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ByteSize"
      Height          =   255
      Index           =   2
      Left            =   2340
      TabIndex        =   3
      Top             =   4620
      Width           =   1035
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Name"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   4620
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Date"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   4620
      Width           =   855
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4515
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   7964
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Connected At"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Pot"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Host"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Search By:"
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4620
      Width           =   795
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuWhois 
         Caption         =   "Whois"
      End
      Begin VB.Menu mnuHexedit 
         Caption         =   "Hexedit"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeName 
         Caption         =   "Change Name"
      End
      Begin VB.Menu mnuCopyFile 
         Caption         =   "Copy File to Desktop"
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
Dim selli As ListItem

Private Sub Command1_Click()
    If lv.ListItems.Count > 0 Then lv.ListItems.Clear
   
    Text1 = Trim(Text1)
    If Len(Text1) = 0 Then
        MsgBox "Enter Search criteria", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo hell
    Dim rs As Recordset, rs2 As Recordset
    Dim li As ListItem
    Dim fsize As Long, i As Long
    
    If Option1(0).value Then 'date
        
        Set rs = ado("Select * from tblData where starttime like '%" & Replace(Text1, "'", """") & "%'")
        fillLV rs
        
    ElseIf Option1(1).value Then 'name
        
        Set rs = ado("Select * from knownsizes where malcode like '%" & Replace(Text1, "'", """") & "%'")
        If rs.EOF Then
            MsgBox "Name not found", vbInformation
            Exit Sub
        End If
        While Not rs.EOF
             fsize = rs!fsize
             Set rs2 = ado("Select * from tblData where fsize=" & fsize)
             fillLV rs2
             rs.MoveNext
        Wend
        
    ElseIf Option1(2).value Then
        Set rs = ado("Select * from tblData where fsize=" & Replace(Text1, "'", """"))
        fillLV rs
    Else
        Set rs = ado("Select * from tbldata where ip like '" & Replace(Text1, "'", """") & "%'")
        fillLV rs
    End If
    
    Set rs = ado("Select * from knownsizes")
    
    While Not rs.EOF
        fsize = ado.RsField("fsize", rs)
        For i = 1 To lv.ListItems.Count
            If lv.ListItems(i).SubItems(2) = fsize Then
                lv.ListItems(i).SubItems(5) = ado.RsField("malcode", rs)
            End If
            DoEvents
        Next
        rs.MoveNext
    Wend
    
    Me.Caption = "Search: [" & lv.ListItems.Count & "] Results"
    
    Exit Sub
hell:     MsgBox Err.Description
End Sub

Sub fillLV(rs As Recordset)
    Dim li As ListItem
    
    On Error Resume Next
    While Not rs.EOF
        Set li = lv.ListItems.Add
        li.Text = ado.RsField("starttime", rs)
        li.SubItems(1) = ado.RsField("honeypot", rs)
        li.SubItems(2) = ado.RsField("fsize", rs)
        li.SubItems(3) = ado.RsField("fName", rs)
        li.SubItems(4) = ado.RsField("ip", rs)
        DoEvents
        rs.MoveNext
    Wend
    
End Sub



Private Sub Command2_Click()
On Error Resume Next
    Dim i, v As String
    v = Replace("Start\tHoneypot\tBytes\tFile\tMalcode", "\t", vbTab) & vbCrLf
    For i = 1 To lv.ListItems.Count
        With lv.ListItems(i)
            v = v & .Text _
                & vbTab & .SubItems(1) & vbTab _
                        & .SubItems(2) & vbTab _
                        & .SubItems(3) & vbTab _
                        & .SubItems(4) & vbTab _
                        & .SubItems(5) & vbCrLf
        End With
    Next
    Clipboard.Clear
    Clipboard.SetText v
End Sub



Private Sub Form_Load()
    lv.ColumnHeaders(4).Width = lv.Width - lv.ColumnHeaders(5).Left - lv.ColumnHeaders(5).Width - 100
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selli = Nothing
    Set selli = Item
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub



Private Sub mnuChangeName_Click()
    If selli Is Nothing Then Exit Sub
    ChangeNameForSize selli, lv, 2, 5
    Set selli = Nothing
End Sub

Private Sub mnuCopyFile_Click()
  On Error GoTo hell
    If selli Is Nothing Then Exit Sub
    
    Dim fName As String
    fName = selli.SubItems(3)
    If Len(Trim(fName)) = 0 Then Exit Sub
   
    If Not fso.FileExists(fName) Then
        MsgBox "File not found: " & vbCrLf & vbCrLf & fName
    Else
        fso.Copy fName, UserDeskTopFolder()
        MsgBox "Copy Complete", vbInformation
    End If
    
    Exit Sub
hell: MsgBox Err.Description
End Sub

Private Sub mnuHexEdit_Click()
     
   On Error GoTo hell
    If selli Is Nothing Then Exit Sub
    
    Dim fName As String
    fName = selli.SubItems(3)
    
    Dim d As New frmHexEdit
    d.loadfile fName
    
Exit Sub
hell:
On Error Resume Next
    MsgBox Err.Description
    
 
End Sub

Private Sub mnuWhois_Click()
    On Error GoTo hell
    If selli Is Nothing Then Exit Sub
    
    Dim ip As String
    ip = Trim(selli.SubItems(4))
    
    If Len(ip) = 0 Then Exit Sub
    
    ip = "cmd /k whois " & ip
    
    Shell ip, vbNormalFocus
    
    Exit Sub
hell: MsgBox Err.Description
End Sub
