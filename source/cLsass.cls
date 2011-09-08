VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRpc445"
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
Event UploadComplete(ip As String, fPath As String, GUID As String, stage As Integer)
Event TimedOut(ip As String)
Event AntiHammer(ByVal ip As String, blockedUntil As String, remove As Boolean)
Event Info(msg As String)
Event RecgonizedExploit(ByVal ip As String, ByVal HandlerName As String)

Public WithEvents server As clsServer
Attribute server.VB_VarHelpID = -1
Public hammer As New CAntiHammer
Private uploads() As clsUpload
Private responses(7) As String

Const r1 = "00 00 55 FF 53 4D 42 72 00 00 00 00 98 53 C8 00 " & _
           "00 00 00 00 00 00 00 00 00 00 00 00 00 FF FE 00 " & _
           "00 00 00 11 05 00 03 0A 00 01 00 04 11 00 00 00 " & _
           "00 01 00 00 00 00 00 FD E3 00 80 96 98 CB DE 56 " & _
           "4F C5 01 A4 01 00 10 00 9D 61 37 F5 8B 36 EA 4E " & _
           "93 2D 9B ED BF 3A 06 2B"

Const r2 = "00 00 01 4F FF 53 4D 42 73 16 00 00 C0 98 07 C8 " & _
           "00 00 00 00 00 00 00 00 00 00 00 00 00 00 FF FE " & _
           "00 08 10 00 04 FF 00 4F 01 00 00 DA 00 24 01 4E " & _
           "54 4C 4D 53 53 50 00 02 00 00 00 1E 00 1E 00 30 " & _
           "00 00 00 15 82 8A E0 10 F4 2C C2 49 C0 F3 BC 00 " & _
           "00 00 00 00 00 00 00 8C 00 8C 00 4E 00 00 00 42 " & _
           "00 4C 00 41 00 48 00 2D 00 38 00 38 00 33 00 39 " & _
           "00 39 00 42 00 37 00 36 00 36 00 44 00 02 00 1E " & _
           "00 42 00 4C 00 41 00 48 00 2D 00 38 00 38 00 33 " & _
           "00 39 00 39 00 42 00 37 00 36 00 36 00 44 00 01 " & _
           "00 1E 00 42 00 4C 00 41 00 48 00 2D 00 38 00 38 " & _
           "00 33 00 39 00 39 00 42 00 37 00 36 00 36 00 44 " & _
           "00 04 00 1E 00 62 00 6C 00 61 00 68 00 2D 00 38 " & _
           "00 38 00 33 00 39 00 39 00 62 00 37 00 36 00 36 " & _
           "00 64 00 03 00 1E 00 62 00 6C 00 61 00 68 00 2D " & _
           "00 38 00 38 00 33 00 39 00 39 00 62 00 37 00 36 " & _
           "00 36 00 64 00 00 00 00 00 00 57 00 69 00 6E 00 " & _
           "64 00 6F 00 77 00 73 00 20 00 35 00 2E 00 30 00 " & _
           "00 00 57 00 69 00 6E 00 64 00 6F 00 77 00 73 00 " & _
           "20 00 32 00 30 00 30 00 30 00 20 00 4C 00 41 00 " & _
           "4E 00 20 00 4D 00 61 00 6E 00 61 00 67 00 65 00 " & _
           "72 00 00"

Const r3 = "00 00 00 75 FF 53 4D 42 73 00 00 00 00 98 07 C8 " & _
           "00 00 00 00 00 00 00 00 00 00 00 00 00 00 FF FE " & _
           "00 08 20 00 04 FF 00 75 00 00 00 00 00 4A 00 4E " & _
           "57 00 69 00 6E 00 64 00 6F 00 77 00 73 00 20 00 " & _
           "35 00 2E 00 30 00 00 00 57 00 69 00 6E 00 64 00 " & _
           "6F 00 77 00 73 00 20 00 32 00 30 00 30 00 30 00 " & _
           "20 00 4C 00 41 00 4E 00 20 00 4D 00 61 00 6E 00 " & _
           "61 00 67 00 65 00 72 00 00"
            
Const r4 = "00 00 00 38 FF 53 4D 42 75 00 00 00 00 98 07 C8 " & _
           "00 00 00 00 00 00 00 00 00 00 00 00 00 08 FF FE " & _
           "00 08 30 00 07 FF 00 38 00 01 00 FF 01 00 00 FF " & _
           "01 00 00 07 00 49 50 43 00 00 00 00"
            
Const r5 = "00 00 00 87 FF 53 4D 42 A2 00 00 00 00 98 07 C8 " & _
           "00 00 00 00 00 00 00 00 00 00 00 00 00 08 DC 04 " & _
           "00 08 40 00 2A FF 00 87 00 00 00 40 01 00 00 00 " & _
           "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 " & _
           "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 " & _
           "80 00 00 00 00 10 00 00 00 00 00 00 00 00 00 00 " & _
           "00 00 00 00 02 00 FF 05 00 00 00 06 00 06 00 40 " & _
           "00 00 00 10 00 10 00 47 00 00 00 15 8A 88 E0 48 " & _
           "00 9B 01 12 00 9B 01 12 00 7A F2"

Const r6 = "00 00 00 7C FF 53 4D 42 25 00 00 00 00 98 07 C8 " & _
           "00 00 00 00 00 00 00 00 00 00 00 00 00 08 DC 04 " & _
           "00 08 50 00 0A 00 00 44 00 00 00 00 00 38 00 00 " & _
           "00 44 00 38 00 00 00 00 00 45 00 00 05 00 0C 03 " & _
           "10 00 00 00 44 00 00 00 01 00 00 00 B8 10 B8 10 " & _
           "E5 74 00 00 0C 00 5C 50 49 50 45 5C 6C 73 61 73 " & _
           "73 00 00 00 01 00 00 00 00 00 00 00 04 5D 88 8A " & _
           "EB 1C C9 11 9F E8 08 00 2B 10 48 60 02 00 00 00"

Const r7 = "00 00 00 2F FF 53 4D 42 2F 00 00 00 00 98 07 C8 " & _
           "00 00 00 00 00 00 00 00 00 00 00 00 00 08 FF FE " & _
           "00 08 60 00 06 FF 00 2F 00 B8 10 FF FF 00 00 00 " & _
           "00 00 00"

Sub StartUp()
    On Error GoTo hell
    server.StartServer
   Exit Sub
hell: dbg llreal, "Rpc445 Startup Error: " & Err.Description
End Sub

Sub ShutDown()
    server.StopServer
End Sub

Private Sub Class_Initialize()
    
    Set server = New clsServer
    
    ReDim uploads(0)
    Set uploads(0) = New clsUpload
    uploads(0).DumpDir = eRPC445
    
    With server
        .port = 2070
        .TimeOut = 900 'has to be small so we can process fast before
    End With           'they close their own attempt to connect to us
    
 
    responses(1) = toHex(r1)
    responses(2) = toHex(r2)
    responses(3) = toHex(r3)
    responses(4) = toHex(r4)
    responses(5) = toHex(r5)
    responses(6) = toHex(r6)
    responses(7) = toHex(r7)
    
 
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
    'Debug.Print "Lsass Closed :" & index
    
    On Error Resume Next
    
    With uploads(Index)
            .CloseFile
            
            If .state > 0 Then
                
                If .Handled Then
                    
                    RaiseEvent Info("Shellcode Handled..Connection Closed " & .RemoteHostIp & " Stage: " & .state & " Filename: " & .filename)
                    
                ElseIf looksLikeJunk(Index) Then
                
                    dbg llspam, "Moving to Junk Folder: " & .filename & " Size: " & .RealFileSize & " Stage: " & .state & " Host: " & .RemoteHostIp
                    MoveFileToDumpDir .filename, eRPC445, "Junk"
                    
                ElseIf Not .Handled Then 'mabey new shellcode mabey junk
                    
                    If Not TryToHandle(Index) Then 'check again just in case (above condition possible)
                        dbg llreal, "[*] No shellcode handler for " & .filename
                        hammer.AddIp .RemoteHostIp
                        RaiseEvent UploadComplete(.RemoteHostIp, .filename, .GUID, .state)
                    Else
                        RaiseEvent Info("Shellcode Handled..Connection Closed " & .RemoteHostIp & " Stage: " & .state & " Filename: " & .filename)
                    End If
                    
                End If
            
            End If
            
    End With
    
    Set uploads(Index) = New clsUpload
    uploads(Index).DumpDir = eRPC445
    
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
    Dim resp As Long
    Dim dce As dce_packet
    Dim startOffset As Long
    
    DoEvents
    
 With uploads(Index)
        
        .state = .state + 1
        resp = IIf(.state > 7, 7, .state)
        
        If .state = 1 Then
            uploads(Index).RemoteHostIp = wsk.RemoteHostIp
            .StartFileDump
        End If

        .WriteChunk data
        .AppendBuffer data
        
        Select Case .state
            Case 0, 1, 2, 3, 4, 5, 7:
            Case 6:
                    
                    If UBound(data) > Len(dce) Then
                        startOffset = UBound(data) - Len(dce) + 1
                        CopyMemory dce, data(startOffset), Len(dce)
                        .GUID = StringFromGUID(dce.interfaceGUID)
                        dbg llreal, "Host: " & .RemoteHostIp & " Type: 0x" & Hex(dce.packetType) & " UUID: " & .GUID & " (" & SSNameFromEnum(SubSploitFromIID(.GUID)) & ")"
                    End If
                    
            Case Else: If TryToHandle(Index, wsk) Then Exit Sub
        End Select
        
        wsk.SendData responses(resp)
         
         
    End With
End Sub

Sub CloseSocket(s As Winsock)
    On Error Resume Next
    s.Close
End Sub

Private Sub Server_NewIndexLoaded(Index As Integer)
    ReDim Preserve uploads(Index)
    Set uploads(Index) = New clsUpload
    uploads(Index).DumpDir = eRPC445
End Sub

Private Function TryToHandle(Index As Integer, Optional wsk As Winsock) As Boolean
    
On Error Resume Next

    Dim buffer As String
    Dim ret As Boolean
    
    Dim cmd As CLsassCmd
    Dim rfile As CLsassRecvFile
    Dim sctftp As CSc_tftp
    Dim gurl As CGenericURL
    Dim pnp As CPnpCmd
    
    Err.Clear
    
    With uploads(Index)
        
        If .Handled Then Exit Function
        
1        buffer = StrConv(.GetBuffer, vbUnicode)
        
        'this one the bot does not wait at all we have to be snappy
        'but quitting here could truncate other shellcodes so only this oen
        'checked here...
2        If recv_cmd.CheckSignature(buffer) Then
                .CloseFile
                .Handled = True
                hammer.AddIp .RemoteHostIp
               If Not wsk Is Nothing Then CloseSocket wsk
    
                RaiseEvent RecgonizedExploit(.RemoteHostIp, "Recv_CMD")
                
                Set cmd = New CLsassCmd
3                ret = cmd.HandleShellcode(.RemoteHostIp, .filename, buffer)
                                
                If ret Then
                    RaiseEvent Info("Command:" & cmd.err_msg)
                    RaiseEvent UploadComplete(.RemoteHostIp, cmd.Success_FileName, .GUID, .state)
                    'hammer.AddIp .RemoteHostIp
                Else
                    RaiseEvent Info("Recv Cmd Failed " & cmd.err_msg & " " & .filename)
                End If
                
                GoTo done
  
        End If
        
4        If recv_file.CheckSignature(buffer) Then
                .CloseFile
               .Handled = True
               hammer.AddIp .RemoteHostIp
               If Not wsk Is Nothing Then CloseSocket wsk
               
               RaiseEvent RecgonizedExploit(.RemoteHostIp, "Recv_File")
               
                Set rfile = New CLsassRecvFile
5                ret = rfile.HandleShellcode(.RemoteHostIp, .filename, buffer)
                             
                If ret Then
                    RaiseEvent Info("[*] Received file: " & rfile.SampleFile & " from " & rfile.RemoteHostIp)
                    RaiseEvent UploadComplete(.RemoteHostIp, rfile.SampleFile, .GUID, .state)
                    'hammer.AddIp .RemoteHostIp
                Else
                    RaiseEvent Info("Recv File Failed: " & rfile.err_msg & " " & .filename)
                End If
                
                GoTo done
                
        End If
        
        
6        If sc_tftp.CheckSignature(buffer) Then
              .CloseFile
              .Handled = True
              hammer.AddIp .RemoteHostIp
             If Not wsk Is Nothing Then CloseSocket wsk
            
             RaiseEvent RecgonizedExploit(.RemoteHostIp, "sc_tftp")
             
7             Set sctftp = New CSc_tftp
8            ret = sctftp.HandleShellcode(.RemoteHostIp, .filename, buffer)
                            
             If ret Then
                  RaiseEvent Info("[*] sc_tftp download: " & sctftp.URL)
                  RaiseEvent UploadComplete(.RemoteHostIp, sctftp.SampleFile, .GUID, .state)
                  'hammer.AddIp .RemoteHostIp
             Else
                 RaiseEvent Info("sc_tftp Failed: " & sctftp.err_msg & " " & .filename)
             End If
             
             GoTo done
             
        End If

        If pnp_cmd.CheckSignature(buffer) Then
              .CloseFile
              .Handled = True
              hammer.AddIp .RemoteHostIp
             If Not wsk Is Nothing Then CloseSocket wsk
            
             RaiseEvent RecgonizedExploit(.RemoteHostIp, "pnp_cmd")
             
            Set pnp = New CPnpCmd
            ret = pnp.HandleShellcode(.RemoteHostIp, .filename, buffer)
                            
             If ret Then
                  RaiseEvent Info("[*] pnpCmd download: " & pnp.Success_FileName)
                  RaiseEvent UploadComplete(.RemoteHostIp, pnp.Success_FileName, .GUID, .state)
                  'hammer.AddIp .RemoteHostIp
             Else
                 RaiseEvent Info("pnpCmd Failed: " & pnp.err_msg & " " & .filename)
             End If
             
             GoTo done
             
        End If


9        Set gurl = New CGenericURL
        
10        If gurl.CheckSignature(buffer) Then
                .CloseFile
                .Handled = True
                hammer.AddIp .RemoteHostIp
                If Not wsk Is Nothing Then CloseSocket wsk
                
                RaiseEvent RecgonizedExploit(.RemoteHostIp, "Generic URL")
                
12              ret = gurl.HandleShellcode(.RemoteHostIp, .filename, buffer)
                
                If ret Then
                    RaiseEvent Info("[*] Generic URL Downloaded file: " & gurl.SampleFile & " from " & gurl.URL)
                    RaiseEvent UploadComplete(.RemoteHostIp, gurl.SampleFile, .GUID, .state)
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







'Response 1 -  Negotioate protocol response server
'------------------------------------------------------------------------------
'00000000   00 00 55 FF 53 4D 42 72  00 00 00 00 98 53 C8 00   ..UˇSMBr....òS».
'00000010   00 00 00 00 00 00 00 00  00 00 00 00 00 FF FE 00   .............ˇ˛.
'00000020   00 00 00 11 05 00 03 0A  00 01 00 04 11 00 00 00   ................
'00000030   00 01 00 00 00 00 00 FD  E3 00 80 96 98 CB DE 56   .......˝„.ÄñòÀﬁV
'00000040   4F C5 01 A4 01 00 10 00  9D 61 37 F5 8B 36 EA 4E   O≈.§....ùa7ıã6ÍN
'00000050   93 2D 9B ED BF 3A 06 2B                            ì-õÌø:.+
'
'Response 2 - session setup response - ntlm challenge server
'------------------------------------------------------------------------------
'00000000   00 00 01 4F FF 53 4D 42  73 16 00 00 C0 98 07 C8   ...OˇSMBs...¿ò.»
'00000010   00 00 00 00 00 00 00 00  00 00 00 00 00 00 FF FE   ..............ˇ˛
'00000020   00 08 10 00 04 FF 00 4F  01 00 00 DA 00 24 01 4E   .....ˇ.O...⁄.$.N
'00000030   54 4C 4D 53 53 50 00 02  00 00 00 1E 00 1E 00 30   TLMSSP.........0
'00000040   00 00 00 15 82 8A E0 10  F4 2C C2 49 C0 F3 BC 00   ....Çä‡.Ù,¬I¿Ûº.
'00000050   00 00 00 00 00 00 00 8C  00 8C 00 4E 00 00 00 42   .......å.å.N...B
'00000060   00 4C 00 41 00 48 00 2D  00 38 00 38 00 33 00 39   .L.A.H.-.8.8.3.9
'00000070   00 39 00 42 00 37 00 36  00 36 00 44 00 02 00 1E   .9.B.7.6.6.D....
'00000080   00 42 00 4C 00 41 00 48  00 2D 00 38 00 38 00 33   .B.L.A.H.-.8.8.3
'00000090   00 39 00 39 00 42 00 37  00 36 00 36 00 44 00 01   .9.9.B.7.6.6.D..
'000000A0   00 1E 00 42 00 4C 00 41  00 48 00 2D 00 38 00 38   ...B.L.A.H.-.8.8
'000000B0   00 33 00 39 00 39 00 42  00 37 00 36 00 36 00 44   .3.9.9.B.7.6.6.D
'000000C0   00 04 00 1E 00 62 00 6C  00 61 00 68 00 2D 00 38   .....b.l.a.h.-.8
'000000D0   00 38 00 33 00 39 00 39  00 62 00 37 00 36 00 36   .8.3.9.9.b.7.6.6
'000000E0   00 64 00 03 00 1E 00 62  00 6C 00 61 00 68 00 2D   .d.....b.l.a.h.-
'000000F0   00 38 00 38 00 33 00 39  00 39 00 62 00 37 00 36   .8.8.3.9.9.b.7.6
'00000100   00 36 00 64 00 00 00 00  00 00 57 00 69 00 6E 00   .6.d......W.i.n.
'00000110   64 00 6F 00 77 00 73 00  20 00 35 00 2E 00 30 00   d.o.w.s. .5...0.
'00000120   00 00 57 00 69 00 6E 00  64 00 6F 00 77 00 73 00   ..W.i.n.d.o.w.s.
'00000130   20 00 32 00 30 00 30 00  30 00 20 00 4C 00 41 00    .2.0.0.0. .L.A.
'00000140   4E 00 20 00 4D 00 61 00  6E 00 61 00 67 00 65 00   N. .M.a.n.a.g.e.
'00000150   72 00 00                                           r..
'
'
'response 3 - Server
'------------------------------------------------------------------------------
'00000000   00 00 00 75 FF 53 4D 42  73 00 00 00 00 98 07 C8   ...uˇSMBs....ò.»
'00000010   00 00 00 00 00 00 00 00  00 00 00 00 00 00 FF FE   ..............ˇ˛
'00000020   00 08 20 00 04 FF 00 75  00 00 00 00 00 4A 00 4E   .. ..ˇ.u.....J.N
'00000030   57 00 69 00 6E 00 64 00  6F 00 77 00 73 00 20 00   W.i.n.d.o.w.s. .
'00000040   35 00 2E 00 30 00 00 00  57 00 69 00 6E 00 64 00   5...0...W.i.n.d.
'00000050   6F 00 77 00 73 00 20 00  32 00 30 00 30 00 30 00   o.w.s. .2.0.0.0.
'00000060   20 00 4C 00 41 00 4E 00  20 00 4D 00 61 00 6E 00    .L.A.N. .M.a.n.
'00000070   61 00 67 00 65 00 72 00  00                        a.g.e.r..
'
'response 4 - Server
'------------------------------------------------------------------------------
'00000000   00 00 00 38 FF 53 4D 42  75 00 00 00 00 98 07 C8   ...8ˇSMBu....ò.»
'00000010   00 00 00 00 00 00 00 00  00 00 00 00 00 08 FF FE   ..............ˇ˛
'00000020   00 08 30 00 07 FF 00 38  00 01 00 FF 01 00 00 FF   ..0..ˇ.8...ˇ...ˇ
'00000030   01 00 00 07 00 49 50 43  00 00 00 00               .....IPC....
'
'response 5 - Create And response - Server
'-------------------------------------------------------------------------------
'00000000   00 00 00 87 FF 53 4D 42  A2 00 00 00 00 98 07 C8   ...áˇSMB¢....ò.»
'00000010   00 00 00 00 00 00 00 00  00 00 00 00 00 08 DC 04   ..............‹.
'00000020   00 08 40 00 2A FF 00 87  00 00 00 40 01 00 00 00   ..@.*ˇ.á...@....
'00000030   00 00 00 00 00 00 00 00  00 00 00 00 00 00 00 00   ................
'00000040   00 00 00 00 00 00 00 00  00 00 00 00 00 00 00 00   ................
'00000050   80 00 00 00 00 10 00 00  00 00 00 00 00 00 00 00   Ä...............
'00000060   00 00 00 00 02 00 FF 05  00 00 00 06 00 06 00 40   ......ˇ........@
'00000070   00 00 00 10 00 10 00 47  00 00 00 15 8A 88 E0 48   .......G....äà‡H
'00000080   00 9B 01 12 00 9B 01 12  00 7A F2                  .õ...õ...zÚ
'
'response 6 - bind ack - server
'-------------------------------------------------------------------------------
'00000000   00 00 00 7C FF 53 4D 42  25 00 00 00 00 98 07 C8   ...|ˇSMB%....ò.»
'00000010   00 00 00 00 00 00 00 00  00 00 00 00 00 08 DC 04   ..............‹.
'00000020   00 08 50 00 0A 00 00 44  00 00 00 00 00 38 00 00   ..P....D.....8..
'00000030   00 44 00 38 00 00 00 00  00 45 00 00 05 00 0C 03   .D.8.....E......
'00000040   10 00 00 00 44 00 00 00  01 00 00 00 B8 10 B8 10   ....D.......∏.∏.
'00000050   E5 74 00 00 0C 00 5C 50  49 50 45 5C 6C 73 61 73   Ât....\PIPE\lsas
'00000060   73 00 00 00 01 00 00 00  00 00 00 00 04 5D 88 8A   s............]àä
'00000070   EB 1C C9 11 9F E8 08 00  2B 10 48 60 02 00 00 00   Î.….üË..+.H`....
'
'
'shellcode comes next as unicode string



