VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Interception Stats"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   5220
      TabIndex        =   2
      Top             =   3540
      Width           =   1215
   End
   Begin VB.CommandButton cmdCOpy 
      Caption         =   "Copy"
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   3540
      Width           =   1215
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   6165
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Interceptions"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmStats"
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
Sub DoReport()
    On Error GoTo hell
    
    Dim rs As Recordset
    Dim li As ListItem
    Dim i As Integer
    Dim tally As Long, total As Long
    
    Me.Visible = True
    
    With Form1
        Set rs = ado("Select * from knownsizes")
        While Not rs.EOF
            Set li = lv.ListItems.Add
            li.Text = ado.RsField("malcode", rs)
            li.SubItems(1) = ado.RsField("fsize", rs)
            li.SubItems(2) = ado("Select count(autoid) as c from tbldata where fsize=" & li.SubItems(1))!c
            tally = tally + li.SubItems(2)
            rs.MoveNext
            DoEvents
        Wend
   
        Set li = lv.ListItems.Add
        total = ado("Select count(autoid) as c from tbldata")!c
        li.Text = "[UnAccounted]"
        li.SubItems(1) = "[Varies]"
        li.SubItems(2) = total - tally
        
        Set li = lv.ListItems.Add
        li.Text = "---------"
        li.SubItems(1) = "-------"
        li.SubItems(2) = "-------"
        
        Set li = lv.ListItems.Add
        li.Text = "TOTAL:"
        li.SubItems(1) = ado("Select count(autoid) as c from tblData")!c & " hits"
        li.SubItems(2) = "from " & ado("Select distinct count(ip) as c from tblData")!c & "distinct ips"
         
   End With
      
      Me.Visible = False
      Me.Show 1
      
   Exit Sub
hell: MsgBox Err.Description
        Unload Me
End Sub

Private Sub cmdCOpy_Click()
    On Error Resume Next
    Dim i As Integer
    Dim v As String
    
    v = Replace("Malcode \tByteLen \tIntercepts", "\t", vbTab) & vbCrLf
    For i = 1 To lv.ListItems.Count
        With lv.ListItems(i)
            v = v & .Text & vbTab & _
                    .SubItems(1) & vbTab & _
                    .SubItems(2) & vbCrLf
        End With
    Next
    Clipboard.Clear
    Clipboard.SetText v
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

