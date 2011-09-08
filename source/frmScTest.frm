VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmScTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shellcode Signature Tester"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHexedit 
      Caption         =   "Hexedit"
      Height          =   315
      Left            =   5640
      TabIndex        =   7
      Top             =   60
      Width           =   915
   End
   Begin MSComctlLib.ListView lv 
      Height          =   1395
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2461
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Loaded Handlers"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run Selected On File"
      Height          =   315
      Left            =   3120
      TabIndex        =   5
      Top             =   1740
      Width           =   3555
   End
   Begin VB.CommandButton chkAutoSigCheck 
      Caption         =   "Auto Test Signature Match"
      Height          =   315
      Left            =   3060
      TabIndex        =   4
      Top             =   1020
      Width           =   3555
   End
   Begin VB.CommandButton cmdCheckSelected 
      Caption         =   "Check Selected for Signature Match"
      Height          =   315
      Left            =   3060
      TabIndex        =   3
      Top             =   660
      Width           =   3555
   End
   Begin VB.TextBox txtScFile 
      Height          =   315
      Left            =   2100
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   60
      Width           =   3435
   End
   Begin VB.Label Label2 
      Caption         =   "Shellcode File (Drag &&Drop )"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "Shellcode Handlers"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmScTest"
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

Private Sub Command1_Click()
    
    If lv.SelectedItem Is Nothing Then
        MsgBox "Select Handler to test"
        Exit Sub
    End If
    
    If Len(txtScFile) = 0 Or Not fso.FileExists(txtScFile) Then
        MsgBox "Sc File not found"
        Exit Sub
    End If
    
    Dim buffer As String
    Dim ret As Boolean
    
    buffer = fso.ReadFile(txtScFile)
    
    Select Case lv.SelectedItem.Index
        Case 1:
                ret = recv_cmd.HandleShellcode("127.0.0.1", txtScFile, buffer)
                MsgBox "Recv Cmd " & IIf(ret, "Succeeded", "Failed") & vbCrLf _
                        & " Msg: " & recv_cmd.err_msg
        Case 2:
                ret = recv_file.HandleShellcode("127.0.0.1", txtScFile, buffer)
                MsgBox "Recv File " & IIf(ret, "Succeeded", "Failed") & vbCrLf & _
                       " Msg: " & recv_file.err_msg
        Case 3:
                ret = generic_url.HandleShellcode("127.0.0.1", txtScFile, buffer)
                MsgBox "Generic URL " & IIf(ret, "Succeeded", "Failed") & vbCrLf & _
                       " Msg: " & generic_url.err_msg & vbCrLf & _
                       " URL: " & generic_url.URL
        Case 4:
                ret = sc_tftp.HandleShellcode("127.0.0.1", txtScFile, buffer)
                MsgBox "sc_tftp " & IIf(ret, "Succeeded", "Failed") & vbCrLf _
                        & " Msg: " & sc_tftp.err_msg & vbCrLf & _
                        " URL: " & sc_tftp.URL
                        
        Case 5: ret = vs_cmd.HandleShellcode("127.0.0.1", txtScFile, buffer)
                MsgBox "VS Cmd " & IIf(ret, "Succeeded", "Failed") & vbCrLf _
                        & " Msg: " & vs_cmd.err_msg
         
        Case 6: ret = pnp_cmd.HandleShellcode("127.0.0.1", txtScFile, buffer)
                MsgBox "PNP_Cmd " & IIf(ret, "Succeeded", "Failed") & vbCrLf _
                        & " Msg: " & pnp_cmd.err_msg
 
 

    End Select
    
    
    If Err.Number > 0 Then
        MsgBox "Error: " & Err.Description
    End If
    
End Sub

Private Sub chkAutoSigCheck_Click()
    
    If Len(txtScFile) = 0 Or Not fso.FileExists(txtScFile) Then
        MsgBox "Sc File not found"
        Exit Sub
    End If
    
    Dim buffer As String
    Dim ret As Boolean
    
    buffer = fso.ReadFile(txtScFile)
    
    
    If recv_cmd.CheckSignature(buffer) Then
        MsgBox "Recv_cmd signature matchs"
        Exit Sub
    End If
    
    If recv_file.CheckSignature(buffer) Then
        MsgBox "recv_file signature matchs"
        Exit Sub
    End If
    
    If generic_url.CheckSignature(buffer) Then
        MsgBox "generic_url Signature matchs"
        Exit Sub
    End If
    
    If sc_tftp.CheckSignature(buffer) Then
        MsgBox "sc_tftp Signature matchs"
        Exit Sub
    End If
    
    If vs_cmd.CheckSignature(buffer) Then
        MsgBox "vs_cmd Signature matchs"
        Exit Sub
    End If
    
    If pnp_cmd.CheckSignature(buffer) Then
        MsgBox "pnp_cmd Signature matchs"
        Exit Sub
    End If
    
    MsgBox "No handler could detect a signature in this file"
    
End Sub

Private Sub cmdCheckSelected_Click()

    If lv.SelectedItem Is Nothing Then
        MsgBox "Select Handler to test"
        Exit Sub
    End If
    
    If Len(txtScFile) = 0 Or Not fso.FileExists(txtScFile) Then
        MsgBox "Sc File not found"
        Exit Sub
    End If
    
    Dim buffer As String
    
    buffer = fso.ReadFile(txtScFile)
    
    Select Case lv.SelectedItem.Index
        Case 1: MsgBox "Recv_cmd.CheckSignature returns: " & recv_cmd.CheckSignature(buffer)
        Case 2: MsgBox "Recv_File.CheckSignature returns: " & recv_file.CheckSignature(buffer)
        Case 3: MsgBox "generic_url.CheckSignature returns: " & generic_url.CheckSignature(buffer)
        Case 4: MsgBox "sc_tftp.CheckSignature returns: " & sc_tftp.CheckSignature(buffer)
        Case 5: MsgBox "vs_cmd.CheckSignature returns: " & vs_cmd.CheckSignature(buffer)
        Case 6: MsgBox "pnp_cmd.CheckSignature returns: " & pnp_cmd.CheckSignature(buffer)
    End Select
    

End Sub

Private Sub cmdHexedit_Click()
    
    Dim D As frmHexEdit
    
    If fso.FileExists(txtScFile) Then
        Set D = New frmHexEdit
        D.loadfile txtScFile
    Else
        MsgBox "File not found"
    End If
    
End Sub



Private Sub Form_Load()

    
    lv.ListItems.Add , , "recv_cmd"
    lv.ListItems.Add , , "recv_file"
    lv.ListItems.Add , , "generic_url"
    lv.ListItems.Add , , "sc_tftp"
    lv.ListItems.Add , , "vs_cmd"
    lv.ListItems.Add , , "pnp_cmd"
    
    lv.ColumnHeaders(1).Width = lv.Width - 300

End Sub

Private Sub txtScFile_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    txtScFile = data.Files(1)
End Sub
