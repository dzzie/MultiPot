VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOptix"
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
Private WithEvents DlPort As clsServer
Attribute DlPort.VB_VarHelpID = -1
Public hammer As New CAntiHammer

Private uploads() As clsUpload


Sub StartUp()
    On Error GoTo hell
    server.StartServer
    DlPort.StartServer
    
    Exit Sub
hell: dbg llreal, "Optix Startup Error: " & Err.Description
End Sub

Sub ShutDown()
    server.StopServer
    DlPort.StopServer
End Sub

Private Sub DlPort_DataReceived(Index As Integer, data() As Byte, wsk As MSWinsockLib.Winsock)
    Dim i As Integer, ip As String
    
    On Error Resume Next
    
    ip = wsk.RemoteHostIp
    i = FindUploderIndexByIP(ip)
    If i = 0 Then
        wsk.SendData "+PASSWORD PLEAZ"
        DoEvents
        wsk.Close
        Exit Sub
    End If
    
    Dim cmd As String
    cmd = StrConv(data, vbUnicode)
    If uploads(i).state = cPrepUpload Then
            uploads(i).state = dUploading
            uploads(i).dlIndex = Index
            
            'sprintf("C:\\a.exe\r\n%d\r\n", dwSize)
             If CountOccurances(cmd, vbCrLf) > 0 Then
                  uploads(i).SetFileSize Split(cmd, vbCrLf)(1)
             End If
             
             uploads(i).StartFileDump
             
             wsk.SendData "+OK REDY" & vbCrLf
    Else
            uploads(i).WriteChunk data()
            
            If uploads(i).UploadComplete Then
                uploads(i).state = eUploadComplete
                uploads(i).CloseFile
                RaiseEvent UploadComplete(ip, uploads(i).filename)
                hammer.AddIp uploads(i).RemoteHostIp
                Set uploads(i) = New clsUpload
                uploads(i).DumpDir = eOptix
            End If
    
    End If
    
End Sub

Private Sub DlPort_Timeout(Index As Integer, closeIt As Boolean)
    Dim i As Integer
    Dim ip As String
    Dim upIdx As Integer
    
    On Error Resume Next
    ip = FindUploderIPByDLIndex(Index)
    upIdx = FindUploderIndexByIP(ip)
    
    Debug.Print "DlPort Timeout: " & Index & "Ip: " & ip
    
    closeIt = True
    
    If upIdx = 0 Then Exit Sub
    
    uploads(upIdx).TimeOut = True
    uploads(upIdx).CloseFile
    
    RaiseEvent TimedOut(ip)
    
    Set uploads(upIdx) = New clsUpload
    uploads(Index).DumpDir = eOptix
      
End Sub

Private Sub Class_Initialize()
    
    Set server = New clsServer
    Set DlPort = New clsServer
    
    ReDim uploads(0)
    Set uploads(0) = New clsUpload
    uploads(0).DumpDir = eOptix
    
    With server
        .port = 2060
        .TimeOut = 60000
    End With

    With DlPort
        .port = 500
        .TimeOut = 60000
    End With
    
End Sub

Private Sub Server_Closed(Index As Integer)
    'Debug.Print "Closed :" & Index
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
    
    RaiseEvent Connection(wsk.RemoteHostIp)
    
    Select Case Left(cmd, 3)
        Case "022"  '"022¬OPtest¬v1.1\r\n"); //OPtest = Default Password, 1.1 Server
            uploads(Index).state = bLoggedIn
            uploads(Index).DumpDir = eOptix
            uploads(Index).RemoteHostIp = wsk.RemoteHostIp
            wsk.SendData "001¬Aweee go fuckyourself why dont ya" & vbCrLf
        Case "019"  '"019¬\r\n" //Prepare Optix For Upload
            uploads(Index).state = cPrepUpload
            wsk.SendData "020¬" & vbCrLf
    End Select
    
End Sub

Private Sub Server_NewIndexLoaded(Index As Integer)
    ReDim Preserve uploads(Index)
    Set uploads(Index) = New clsUpload
    uploads(Index).DumpDir = eOptix
End Sub

Private Sub Server_Timeout(Index As Integer, closeIt As Boolean)
    'Debug.Print "Timeout: " & Index
    closeIt = True
End Sub

Private Function FindUploderIndexByIP(ip As String) As Integer
    
    On Error Resume Next
    Dim i As Integer
    
    For i = 0 To UBound(uploads)
        If uploads(i).RemoteHostIp = ip Then
            FindUploderIndexByIP = i
            Exit Function
        End If
    Next
    
End Function

Private Function FindUploderIPByDLIndex(dlIdx As Integer) As String
    
    On Error Resume Next
    Dim i As Integer
    
    For i = 0 To UBound(uploads)
        If uploads(i).dlIndex = dlIdx Then
            FindUploderIPByDLIndex = uploads(i).RemoteHostIp
            Exit Function
        End If
    Next
    
End Function
