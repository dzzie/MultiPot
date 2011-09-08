VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVeritas"
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

Event Connection(ip As String, Block As Boolean)
Event UploadComplete(ip As String, fPath As String)
Event TimedOut(ip As String)
Event AntiHammer(ByVal ip As String, blockedUntil As String, remove As Boolean)
Event Info(msg As String)
Event RecgonizedExploit(ByVal ip As String, ByVal HandlerName As String)

Public WithEvents server As clsServer
Attribute server.VB_VarHelpID = -1
Public hammer As New CAntiHammer
Private uploads() As clsUpload
Private responses(2) As String


Sub StartUp()
    On Error GoTo hell
    server.StartServer
   Exit Sub
hell: dbg llreal, "Veritas Startup Error: " & Err.Description
End Sub

Sub ShutDown()
    server.StopServer
End Sub

Private Sub Class_Initialize()
    
    Set server = New clsServer
    
    ReDim uploads(0)
    Set uploads(0) = New clsUpload
    uploads(0).DumpDir = everitas
    
    With server
        .port = 6101
        .TimeOut = 500 'has to be small so we can process fast before
    End With           'they close their own attempt to connect to us
    
 
     
 
End Sub

Private Sub class_terminate()
    On Error Resume Next
    Unload server
    Set server = Nothing
End Sub

Private Sub Server_Timeout(Index As Integer, closeIt As Boolean)
    'Debug.Print "Timeout: " & index
    
    With uploads(Index)
            .CloseFile
            closeIt = True
    End With

End Sub

Private Sub Server_Closed(Index As Integer)
    
    On Error Resume Next
    
    With uploads(Index)
            .CloseFile
            
            If .state > 0 Then
                
                If .Handled Then
                    RaiseEvent Info("Shellcode Handled..Connection Closed " & .RemoteHostIp & " Stage: " & .state & " Filename: " & .filename)
                ElseIf looksLikeJunk(Index) Then
                    dbg llspam, "Moving to Junk Folder: " & .filename & " Size: " & .RealFileSize & " Stage: " & .state & " Host: " & .RemoteHostIp
                    MoveFileToDumpDir .filename, everitas, "Junk"
                ElseIf Not .Handled Then 'mabey new shellcode mabey junk
                    If Not TryToHandle(Index) Then 'check again just in case (above condition possible)
                        dbg llreal, "[*] No shellcode handler for " & .filename
                        hammer.AddIp .RemoteHostIp
                        RaiseEvent UploadComplete(.RemoteHostIp, .filename)
                    Else
                        RaiseEvent Info("Shellcode Handled..Connection Closed " & .RemoteHostIp & " Stage: " & .state & " Filename: " & .filename)
                    End If
                End If
                
            End If
            
    End With
    
    Set uploads(Index) = New clsUpload
    uploads(Index).DumpDir = everitas
    
End Sub

Function looksLikeJunk(Index As Integer) As Boolean
    
    On Error Resume Next
    Dim tmp As String
    Dim ok As Boolean
    
    'will need to expand this out more as we go to reduce false positives...
    'right now can not detect truncated shellcodes, or those with non traditionsl
    'nop slides...
    
    On Error Resume Next
    
    With uploads(Index)
        If Len(.filename) = 0 Then GoTo iThinkSo
        If .RealFileSize < 1000 Then GoTo iThinkSo
                
        tmp = .GetBufStr
        
        If InStr(tmp, String(Chr(&H90), 10)) > 0 Then ok = True
        If InStr(tmp, Chr(0) & Chr(90) & Chr(0) & Chr(90)) > 0 Then ok = True
        If FileLen(.filename) > 1200 Then ok = True 'mabey non-traditional nop slide ?
        
        If Not ok Then GoTo iThinkSo
        
        looksLikeJunk = False
    End With

Exit Function
iThinkSo: looksLikeJunk = True
End Function

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
    
    RaiseEvent Connection(remoteHost, Block)
    Debug.Print "ConnectReq: " & remoteHost
    
End Sub

Private Sub Server_Error(Index As Integer, Number As Integer, desc As String)
    Debug.Print "Error: " & Index & " Desc: " & desc
    server.CloseIndex Index
End Sub

Private Sub Server_DataReceived(Index As Integer, data() As Byte, wsk As Winsock)
    
    On Error Resume Next
    

    Dim buffer As String
    Dim ret As Boolean
    
    DoEvents
    
 With uploads(Index)
        
        .state = .state + 1
        
        If .state = 1 Then
            uploads(Index).RemoteHostIp = wsk.RemoteHostIp
            .StartFileDump
        End If

        .WriteChunk data
        .AppendBuffer data
        
        Select Case .state
            Case 0 ', 1, 2, 3, 4, 5, 6, 7:
            Case Else: If TryToHandle(Index, wsk) Then Exit Sub
        End Select
         
    End With
End Sub

Sub CloseSocket(s As Winsock)
    On Error Resume Next
    s.Close
End Sub

Private Sub Server_NewIndexLoaded(Index As Integer)
    ReDim Preserve uploads(Index)
    Set uploads(Index) = New clsUpload
    uploads(Index).DumpDir = everitas
End Sub

Private Function TryToHandle(Index As Integer, Optional wsk As Winsock) As Boolean
    
On Error Resume Next

    Dim buffer As String
    Dim ret As Boolean
   
    Dim gurl As CGenericURL
    Dim vscmd As CVeritasCmd
    
    With uploads(Index)
        
        If .Handled Then Exit Function

        buffer = StrConv(.GetBuffer, vbUnicode)
        
        If vs_cmd.CheckSignature(buffer) Then
            .CloseFile
            .Handled = True
            hammer.AddIp .RemoteHostIp
            If Not wsk Is Nothing Then CloseSocket wsk

            RaiseEvent RecgonizedExploit(.RemoteHostIp, "VS_CMD")
            
            Set vscmd = New CVeritasCmd
            ret = vscmd.HandleShellcode(.RemoteHostIp, .filename, buffer)
            
            If ret Then
                RaiseEvent Info("VS Command Succeeded:" & vscmd.err_msg)
                RaiseEvent UploadComplete(.RemoteHostIp, vscmd.Success_FileName)
                'hammer.AddIp .RemoteHostIp
            Else
                 RaiseEvent Info("VS Cmd Failed " & vscmd.err_msg & " " & .filename)
            End If
            
            GoTo done
            
        End If

        Set gurl = New CGenericURL
       
        If gurl.CheckSignature(buffer) Then
              .CloseFile
              .Handled = True
              hammer.AddIp .RemoteHostIp
              If Not wsk Is Nothing Then CloseSocket wsk
               
              RaiseEvent RecgonizedExploit(.RemoteHostIp, "Generic URL")

              ret = gurl.HandleShellcode(.RemoteHostIp, .filename, buffer, everitas)
                
              If ret Then
                    RaiseEvent Info("[*] Generic URL Downloaded file: " & gurl.SampleFile & " from " & gurl.URL)
                    RaiseEvent UploadComplete(.RemoteHostIp, gurl.SampleFile)
                    'hammer.AddIp .RemoteHostIp
              Else
                    RaiseEvent Info("Generic Url Failed: " & gurl.URL)
                    RaiseEvent Info("Generic Url ErrMsg: " & gurl.err_msg & " File:" & .filename)
              End If

         End If
        
    End With
    
done:
      

    If Err.Number > 0 Then
        RaiseEvent Info("Had Error in TrytoHandle: " & Err.Description & " Line: " & Erl)
    End If
    
    TryToHandle = uploads(Index).Handled
    
    
End Function







