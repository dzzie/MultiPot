VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form clsServer 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   60000
      Left            =   540
      Top             =   60
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   0
      Left            =   60
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "clsServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this class handles the implementation details of creating a server
'socket that can serve manage multiple clients at the time and can
'sets up timeout intervals etc.
'
'this class in turn presents upper layers with a simplified connection
'based notification of socket events and the data they provide.


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

Public maxSockets As Integer
Public TimeOut As Long
Public port As Long

Public WelcomeConnect_Data As String 'CRLF delimited list of commands to send on conenct

Event ConnectionRequest(remoteHost As String, Block As Boolean)
Event Closed(Index As Integer)
Event DataReceived(Index As Integer, data() As Byte, wsk As Winsock)
Event Error(Index As Integer, Number As Integer, desc As String)
Event TimeOut(Index As Integer, closeIt As Boolean)
Event NewIndexLoaded(Index As Integer)

Property Get SocketCount() As Integer
    On Error Resume Next
    SocketCount = ws.UBound
End Property

Property Get SocketStats() As String
    On Error Resume Next
    Dim i As Integer
    Dim ret() As String
    Dim ip As String
    
    push ret, "Index       State       IP"
    
    For i = 0 To ws.UBound
        ip = ws(i).RemoteHostIp
        push ret(), i & vbTab & ws(i).state & vbTab & ip
    Next
    
    SocketStats = Join(ret, vbCrLf)
    
End Property

Sub StartServer(Optional mport As Long)

   If mport <> 0 Then port = mport
   If maxSockets = 0 Then maxSockets = 30
   If TimeOut = 0 Then TimeOut = 60000
   
   ws(0).LocalPort = port
   ws(0).Listen
   
End Sub

 

Sub StopServer()
    Dim i As Integer
    
    On Error Resume Next
    
    ws(0).Close
    
    For i = ws.UBound To 1 Step -1
        ws(i).Close
        tmrTimeout(i).Enabled = False
        Unload tmrTimeout(i)
        Unload ws(i)
    Next
    
End Sub

Public Sub CloseIndex(Index As Integer)
    ws_Close Index
End Sub

    

Private Sub tmrTimeout_Timer(Index As Integer)
    Dim closeIt As Boolean
    On Error Resume Next
    
    RaiseEvent TimeOut(Index, closeIt)
    
    tmrTimeout(Index).Enabled = False
    
    If closeIt Then
        ws(Index).Close
        ws_Close Index
    End If
    
End Sub

Private Sub ws_Close(Index As Integer)
    On Error Resume Next
    ws(Index).Close
    tmrTimeout(Index).Enabled = False
    RaiseEvent Closed(Index)
End Sub

Private Sub ws_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim x As Integer
    Dim i As Integer
    Dim shouldBlock As Boolean
    
    On Error Resume Next
    
    RaiseEvent ConnectionRequest(CStr(ws(0).RemoteHostIp), shouldBlock)
    
    If shouldBlock Then Exit Sub

    x = -1
    For i = 1 To ws.UBound
        If ws(i).state <> sckConnected And _
           ws(i).state <> sckConnecting And _
           ws(i).state <> sckConnectionPending Then
           '------
           x = i
           Exit For
        End If
    Next

    If x < 1 Then
        If ws.UBound > maxSockets Then
            Log "Maxsockets Sockets Reached (" & maxSockets & ") Denying connection>" & ws(Index).RemoteHostIp
            Exit Sub
        Else
            x = ws.UBound + 1
            Load ws(x)
            Load tmrTimeout(x)
            RaiseEvent NewIndexLoaded(x)
        End If
    End If
    
    ws(x).Close
    ws(x).Accept requestID
    tmrTimeout(x).Interval = TimeOut
    tmrTimeout(x).Enabled = True
    
    If Len(WelcomeConnect_Data) > 0 Then
        Dim tmp() As String
        tmp = Split(WelcomeConnect_Data, vbCrLf)
        For i = 0 To UBound(tmp)
            tmp(i) = Replace(tmp(i), "\n", vbCrLf)
            ws(x).SendData tmp(i)
            DoEvents
        Next
    End If
        
End Sub


Private Sub ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim hnd As Long
    Dim B() As Byte
    Dim fsize As Long
    
    On Error Resume Next
    
    ReDim B(bytesTotal)
    
    ws(Index).GetData B(), vbByte, bytesTotal
    
    RaiseEvent DataReceived(Index, B(), ws(Index))
    

    tmrTimeout(Index).Enabled = False '\__ Reset timeout interval
    tmrTimeout(Index).Enabled = True  '/
    
End Sub

Private Sub ws_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    
    Debug.Print "Err sck: " & Index & " Desc: " & Description
    
    RaiseEvent Error(Index, Number, Description)
   
    ws(Index).Close
    ws_Close Index
    
End Sub





Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Integer
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
