VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBagleFtpRecv"
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
Event Connect(ip As String)
Event DataRecv(index As Integer, size As Long)
Event TimeOut(index As Integer)
Event Closed(index As Integer)

Private WithEvents FtpDataConnection As clsServer
Attribute FtpDataConnection.VB_VarHelpID = -1

Private uploads() As clsUpload

Sub StartUp()
    FtpDataConnection.StartServer
End Sub
Sub ShutDown()
    FtpDataConnection.StopServer
End Sub


Property Get remoteHost(index As Integer) As String
    On Error Resume Next
    remoteHost = uploads(index).RemoteHostIp
End Property

Property Get Filename(index As Integer) As String
    On Error Resume Next
    Filename = uploads(index).Filename
End Property

Sub CloseFile(index As Integer)
    On Error Resume Next
    uploads(index).CloseFile
End Sub

Private Sub Class_Initialize()
    
    Set FtpDataConnection = New clsServer
    ReDim uploads(0)
    
    Set uploads(0) = New clsUpload
    uploads(0).DumpDir = eBagle
    
    With FtpDataConnection
        .port = 12345
        .TimeOut = 60000
    End With
    
End Sub

Private Sub FtpDataConnection_Closed(index As Integer)
    RaiseEvent Closed(index)
    Set uploads(index) = New clsUpload
    uploads(index).DumpDir = eBagle
End Sub

Private Sub FtpDataConnection_ConnectionRequest(remoteHost As String, Block As Boolean)
    Debug.Print "FtpDataConnection ConnectReq: " & remoteHost
    RaiseEvent Connect(remoteHost)
End Sub

Private Sub FtpDataConnection_DataReceived(index As Integer, data() As Byte, wsk As MSWinsockLib.Winsock)
   On Error GoTo hell
   
   With uploads(index)
        
        If Len(.RemoteHostIp) = 0 Then .RemoteHostIp = wsk.RemoteHostIp
        
        If .state = aPreLogin Then
              .state = bLoggedIn
              .StartFileDump
        End If
        
        .WriteChunk data()
        RaiseEvent DataRecv(index, UBound(data))
    
    End With

Exit Sub
hell:
    Debug.Print "FtpConnectiondatarecv err: " & Err.Description
End Sub

Private Sub FtpDataConnection_Error(index As Integer, number As Integer, desc As String)
    Debug.Print "FtpDataConnection Err: " & desc
    RaiseEvent TimeOut(index)
End Sub

Private Sub FtpDataConnection_NewIndexLoaded(index As Integer)
    ReDim Preserve uploads(index)
    Set uploads(index) = New clsUpload
    uploads(index).DumpDir = eBagle
End Sub

Private Sub FtpDataConnection_TimeOut(index As Integer, closeIt As Boolean)
    RaiseEvent TimeOut(index)
    closeIt = True
End Sub
