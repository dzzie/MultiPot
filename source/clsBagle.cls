VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBagle"
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
'Event BadUpload(ip As String, fpath As String)

Private WithEvents server As clsServer
Attribute server.VB_VarHelpID = -1
Public hammer As New CAntiHammer

Private uploads() As clsUpload

Sub StartUp()
    On Error GoTo hell
    server.StartServer
   Exit Sub
hell: dbg llreal, "Bagle Startup Error: " & Err.Description
End Sub

Sub ShutDown()
    server.StopServer
End Sub

Private Sub Class_Initialize()
    
    Set server = New clsServer
    
    ReDim uploads(0)
    Set uploads(0) = New clsUpload
    uploads(0).DumpDir = eBagle
    
    With server
        .port = 2745
        .TimeOut = 60000
    End With
 
End Sub

Private Sub Server_Closed(Index As Integer)
    Debug.Print "Bagle Closed :" & Index
    
    Dim fPath As String
    Dim d As clsBagleDownloader
    
    With uploads(Index)
            
            Set d = New clsBagleDownloader
                
            fPath = .filename
                
            If InStr(fPath, "0.0.0.0") > 0 Then fPath = Replace(fPath, "0.0.0.0", .RemoteHostIp)
            
            If d.LoadFTPUrl(fPath) Then
                If d.DoBagleDownload() Then fPath = d.myFileName Else fPath = "DlFailed: " & fPath
                Unload d
                Set d = Nothing
                hammer.AddIp .RemoteHostIp
            Else
                Unload d
                Set d = Nothing
                fPath = "Load UrlFailed: " & fPath
            End If
            
            RaiseEvent UploadComplete(.RemoteHostIp, fPath)
            
    
    End With
    
    Set uploads(Index) = New clsUpload
    uploads(Index).DumpDir = eBagle
    
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
    
    If data(0) = &H43 And data(1) = &HFF Then 'bagle login
       
        RaiseEvent Connection(wsk.RemoteHostIp)
        
        uploads(Index).state = dUploading
        uploads(Index).RemoteHostIp = wsk.RemoteHostIp
        
        Dim resp() As Byte
        resp = StrConv(String(8, "A"), vbFromUnicode)
        
        wsk.SendData resp()
        DoEvents
        
    ElseIf uploads(Index).state = dUploading Then
            
        uploads(Index).filename = cmd
        DoEvents
        wsk.Close
    
    Else
    
        DoEvents
        wsk.Close
    
    End If
    
    
    
End Sub

Private Sub Server_NewIndexLoaded(Index As Integer)
    ReDim Preserve uploads(Index)
    Set uploads(Index) = New clsUpload
    uploads(Index).DumpDir = eBagle
End Sub

Private Sub Server_Timeout(Index As Integer, closeIt As Boolean)
    'Debug.Print "Timeout: " & Index
    
    uploads(Index).state = gTimedOut
    
    RaiseEvent TimedOut(uploads(Index).RemoteHostIp)
    
    closeIt = True

End Sub




