VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSub7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
Event ConnectionRequest(ip As String, Block As Boolean)
Event AntiHammer(ByVal ip As String, blockedUntil As String, remove As Boolean)
Event Connection(ip As String)
Event UploadComplete(ip As String, fPath As String)
Event TimedOut(ip As String)

Private WithEvents server As clsServer
Attribute server.VB_VarHelpID = -1
Public hammer As New CAntiHammer

Private uploads() As clsUpload

Sub StartUp()
    On Error GoTo hell
    server.StartServer
    Exit Sub
hell: dbg llreal, "Sub7 Startup Error: " & Err.Description
End Sub

Sub ShutDown()
    server.StopServer
End Sub

Private Sub Class_Initialize()
    
    Set server = New clsServer
 
    
    ReDim uploads(0)
    Set uploads(0) = New clsUpload
    uploads(0).DumpDir = esub7
    
    With server
        .WelcomeConnect_Data = "PWD"
        .port = 27347
        .TimeOut = 60000
    End With
 
End Sub

Private Sub Server_Closed(Index As Integer)
    'Debug.Print "Bagle Closed :" & Index
    
    With uploads(Index)
            
            .CloseFile
            
            RaiseEvent UploadComplete(.RemoteHostIp, .filename)
            hammer.AddIp .RemoteHostIp
    
    End With
    
    Set uploads(Index) = New clsUpload
    uploads(Index).DumpDir = esub7
    
End Sub

Private Sub Server_ConnectionRequest(remoteHost As String, Block As Boolean)

    Dim blockUntil As String
    Dim remove As Boolean
    
    If hammer.shouldBlock(remoteHost, blockUntil) Then
        RaiseEvent AntiHammer(remoteHost, blockUntil, remove)
        If remove Then
            hammer.RemoveIp remoteHost
        Else
           Block = True
           Exit Sub
        End If
    End If
    
    'Debug.Print "ConnectReq: " & remoteHost
    RaiseEvent ConnectionRequest(remoteHost, Block)
    
End Sub

Private Sub Server_Error(Index As Integer, Number As Integer, desc As String)
   'Debug.Print "Error: " & Index & " Desc: " & desc
End Sub

Private Sub Server_DataReceived(Index As Integer, data() As Byte, wsk As Winsock)
    Dim cmd As String
    cmd = StrConv(data, vbUnicode)
    
    On Error Resume Next
    
    'Sleep 100
    DoEvents
    
    If uploads(Index).state = aPreLogin Then
               
        If InStr(cmd, "PWD715") > 0 Or InStr(cmd, "PWD14438136782715101980") > 0 Then
            
            Debug.Print "Stage1"
            
            RaiseEvent Connection(wsk.RemoteHostIp)
        
            uploads(Index).state = bLoggedIn
            uploads(Index).RemoteHostIp = wsk.RemoteHostIp

            wsk.SendData "connected."
            DoEvents
            
        Else
            wsk.Close
            DoEvents
        End If
         
        
    ElseIf uploads(Index).state = bLoggedIn Then
            
        If InStr(cmd, "UPS") > 0 Then
            Debug.Print "Stage2"
            uploads(Index).state = cPrepUpload
            wsk.SendData "TID"
            DoEvents
        Else
            wsk.Close
            DoEvents
        End If
        
    
    ElseIf uploads(Index).state = cPrepUpload Then
        
        If InStr(cmd, "SFT") > 0 Then
            Debug.Print "Stage3"
            uploads(Index).state = dUploading
            uploads(Index).FileSize = CLng(Mid(cmd, 6, 5))
            Debug.Print "FileSize=" & uploads(Index).FileSize & " len(cmd) " & Len(cmd)
            uploads(Index).StartFileDump
            If Len(cmd) > 10 Then
                uploads(Index).WriteChunk data()
            End If
        Else
            wsk.Close
            DoEvents
        End If
        
    
    ElseIf uploads(Index).state = dUploading Then
        
        uploads(Index).WriteChunk data()
        DoEvents
        
        If uploads(Index).UploadComplete Then
            Debug.Print "Upload Done!"
            wsk.SendData ""
            DoEvents
            server.CloseIndex Index
            DoEvents
            RaiseEvent UploadComplete(uploads(Index).RemoteHostIp, uploads(Index).filename)
            Set uploads(Index) = New clsUpload
            uploads(Index).DumpDir = esub7
        End If
        
    ElseIf uploads(Index).state = eUploadComplete Then
        
        
        
    End If
        
    DoEvents
    DoEvents
    DoEvents
    
End Sub

Private Sub Server_NewIndexLoaded(Index As Integer)
    ReDim Preserve uploads(Index)
    Set uploads(Index) = New clsUpload
    uploads(Index).DumpDir = esub7
End Sub

Private Sub Server_Timeout(Index As Integer, closeIt As Boolean)
    'Debug.Print "Timeout: " & Index
    
    uploads(Index).state = gTimedOut
    uploads(Index).CloseFile
    
    RaiseEvent TimedOut(uploads(Index).RemoteHostIp)
    
    closeIt = True

End Sub






