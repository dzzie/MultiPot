VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMyDoom"
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
Event BadUpload(ip As String, fPath As String)

Private WithEvents server As clsServer
Attribute server.VB_VarHelpID = -1
Public hammer As New CAntiHammer

Private uploads() As clsUpload

Sub StartUp()
    On Error GoTo hell
    server.StartServer
       Exit Sub
hell: dbg llreal, "MyDoom Startup Error: " & Err.Description
End Sub

Sub ShutDown()
    server.StopServer
End Sub

Private Sub Class_Initialize()
    
    Set server = New clsServer
 
    
    ReDim uploads(0)
    Set uploads(0) = New clsUpload
    uploads(0).DumpDir = eMyDoom
    
    With server
        .port = 3127
        .TimeOut = 60000
    End With
 
End Sub

Private Sub Server_Closed(Index As Integer)
   'Debug.Print "Closed :" & Index
    
    With uploads(Index)
    
        If .state = dUploading Then
            .state = eUploadComplete
            .CloseFile
            RaiseEvent UploadComplete(.RemoteHostIp, .filename)
            hammer.AddIp .RemoteHostIp
        Else
            .CloseFile
            .state = fTerminate
            If Len(.RemoteHostIp) > 0 Then RaiseEvent BadUpload(.RemoteHostIp, .filename)
        End If
    
    End With
    
    Set uploads(Index) = New clsUpload
    uploads(Index).DumpDir = eMyDoom
    
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
    
    If data(0) = &H85 And data(1) = &H13 Then 'doom login
        
        RaiseEvent Connection(wsk.RemoteHostIp)
        
        uploads(Index).state = dUploading
        uploads(Index).RemoteHostIp = wsk.RemoteHostIp
        uploads(Index).StartFileDump
        uploads(Index).WriteChunk TrimDoomLogin(data())
    
    ElseIf uploads(Index).state = dUploading Then
            
        uploads(Index).WriteChunk data()
    
    Else
        
        wsk.Close
    
    End If
    
    
    
End Sub

Private Sub Server_NewIndexLoaded(Index As Integer)
    ReDim Preserve uploads(Index)
    Set uploads(Index) = New clsUpload
    uploads(Index).DumpDir = eMyDoom
End Sub

Private Sub Server_Timeout(Index As Integer, closeIt As Boolean)
    'Debug.Print "Timeout: " & Index
    
    uploads(Index).state = gTimedOut
     closeIt = True
     
    If Len(uploads(Index).RemoteHostIp) = 0 Then Exit Sub
    RaiseEvent TimedOut(uploads(Index).RemoteHostIp)
    
End Sub


Function TrimDoomLogin(data() As Byte) As Byte()
    Dim B() As Byte
    ReDim B(UBound(data) - 5)
    
    CopyMemory B(0), data(5), (UBound(B) + 1)
    
    TrimDoomLogin = B()
    
End Function

