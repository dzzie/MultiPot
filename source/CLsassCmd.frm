VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form CLsassCmd 
   Caption         =   "CLsassCmd"
   ClientHeight    =   450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2910
   LinkTopic       =   "Form2"
   ScaleHeight     =   450
   ScaleWidth      =   2910
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   0
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "CLsassCmd"
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

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private start_signature As String
Private banner As String
Private state As Long
Private data  As String

Public err_msg As String
Public Busy As Boolean
Public Success_FileName As String

'should we use this to block multiple cmds from various bots to same cmdlien?
'only work for multibots, that use central server for downloads...humm
'Private cmdCache As Collection


Function BlockWhileBusy()
    
    While Busy
        DoEvents
        Sleep 30
    Wend
    
End Function


Sub Form_Load()

    Timer1.Interval = 5000
    
    banner = "Microsoft Windows 2000 [Version 5.00.2195]" & vbCrLf & _
             "(C) Copyright 1985-2000 Microsoft Corp." & vbCrLf & _
             vbCrLf & _
             "C:\WinNT\System32>"
      
    '751c123c value can change..its not an opcode portion, so we wil shift our
    'signature 14 bytes forward to cover the specific encoder,
    start_signature = "eb 6 eb 6 3c 12 1c 75 90 90 90 90 90 90 90 90 eb 10 5a 4a"
    'start_signature = "90 90 eb 10 5a 4a 33 c9 66 b9 7d 01 80 34 0a"
    
    Dim tmp() As String, i As Long
    tmp = Split(start_signature, " ")
    
    For i = 0 To UBound(tmp)
        tmp(i) = Chr(CInt("&h" & tmp(i)))
    Next
    
    start_signature = Join(tmp, "")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
End Sub

Function CheckSignature(ByVal MemBuffer As String) As Boolean
    
    Dim tmp As String, start As Long
    
    If Len(start_signature) = 0 Then Form_Load
    
    tmp = Replace(MemBuffer, Chr(0), "")
    start = InStrRev(tmp, start_signature)
   
    If start > 1 Then CheckSignature = True
    
End Function

Function HandleShellcode(remoteHost As String, filepath As String, MemBuffer As String) As Boolean
     
   Dim ftp As CFtpGet
   Dim tftp As CTFTPClient
   Dim myFile As String
   Dim ip As String, File As String
   Dim ret As Boolean
   Dim mode As dataMode
   
   On Error Resume Next
   
   Busy = True
   Success_FileName = Empty
   'mode = mAuto
                  
   If Len(start_signature) = 0 Then Form_Load
   
   ret = processfile(MemBuffer)
   MoveFileToDumpDir filepath, eRPC445, "\recv_cmd"
   'dbg "Parsing LsassCmd Received: " & data
   
   If Not ret Then
        dbg llinfo, "Failed to process file in LsassCmd handler Error:" & err_msg
        Exit Function
   End If
   
    If AllOfTheseInstr(data, "echo,>>,ftp") Then
           Set ftp = New CFtpGet
           If ftp.LoadEchoString(data, remoteHost) Then
               ftp.saveAs = GetFreeFileInDumpDir(eRPC445, "./../")
               If Len(ftp.user) = 1 Then mode = mport
               If ftp.GetFile(mode) Then
                   dbg llreal, "[*] FTP Downloaded File: " & ftp.saveAs
                   Success_FileName = ftp.saveAs
                   HandleShellcode = True
               Else
                   dbg llinfo, "Ftp download failed for ErrorLog: " & Join(ftp.ErrorLog, vbCrLf)
                   'if instr(ftp.ErrorMessage ,"Incomplete Download") > 0 and
                   Exit Function
               End If
           Else
               dbg llinfo, "FTP.LoadEchoString Failed :" & data
               Exit Function
           End If
           Unload ftp
           Set ftp = Nothing
     ElseIf AllOfTheseInstr(data, "tftp,get") Then
           Set tftp = New CTFTPClient
           If tftp.ParseTftpCmd(data, remoteHost, ip, File) Then
               myFile = GetFreeFileInDumpDir(eRPC445, "./../")
               If tftp.GetFile(File, myFile, ip) Then
                   dbg llreal, "[*] TFTP Download Succeeded: " & myFile
                   HandleShellcode = True
                   Success_FileName = myFile
               Else
                   dbg llinfo, "TFTP Download Failed Error: " & tftp.ErrorMsg & " Filename: " & tftp.RemoteFileName
               End If
           Else
               dbg llinfo, "TFTP could not parse cmd string: " & data
           End If
           Unload tftp
           Set tftp = Nothing
      Else
           dbg llreal, "[*] Unknown Command Received: " & data
      End If
 
   Busy = False
   CloseSocket
   
End Function


Function processfile(ByVal memBuf As String) As Boolean
    Dim tmp As String
    Dim start As Long, endat As Long
    Dim buf() As Byte, xor_key As Byte
    
    tmp = Replace(memBuf, Chr(0), Empty)
    
    start = InStrRev(tmp, start_signature)
    'start = start - 14 'to realign offsets below with old signature start
    
    If start < 1 Then
        err_msg = "Signature not found!"
        Exit Function
    End If
    
    endat = InStr(start + 20, tmp, String(6, Chr(&H90)))
    
    If endat < 1 Then
        err_msg = "End sig not found shellcode truncated?"
        Exit Function
    End If
    
    tmp = Mid(tmp, start, endat - start)
    buf = StrConv(tmp, vbFromUnicode)
    
    xor_key = buf(&H1D)
    
    If xor_key <> &H99 Then
        err_msg = "Altered xorkey?"
        Exit Function
    End If
    
    Dim i As Integer
    'For i = &H27 To UBound(buf)
    For i = &HBD To &HC3             'only xor decode what we need
        buf(i) = buf(i) Xor xor_key
    Next
    
    If UBound(buf) < &H1A3 Then
        err_msg = "Invalid shellcode length?"
        Exit Function
    End If
    
    CopyMemory i, buf(&HBE), 2
    If i <> 2 Then
        err_msg = "Invalid offset?"
        Exit Function
    End If
    
    Dim port As Long
    
    port = swapLong(buf(&HC0), buf(&HC1))
    
    On Error GoTo hell
    state = 0
    err_msg = Empty
    data = Empty
    CloseSocket
1   ws.LocalPort = port
2   ws.Listen
    Timer1.Enabled = True
    
    While Timer1.Enabled
        DoEvents
        Sleep 30
    Wend
    
    If Len(data) = Empty Then
        err_msg = "port: " & port & " Failed! " & err_msg
    Else
        err_msg = "port: " & port & data
        processfile = True
    End If
    
    
    
 Exit Function
hell:  err_msg = "Error line: " & Erl & " Desc: " & Err.Description
End Function

Private Function CloseSocket()
    On Error Resume Next
    ws.Close
End Function

Private Function swapLong(b1 As Byte, b2 As Byte) As Long
    Dim B(3) As Byte

    B(0) = b2
    B(1) = b1
    B(2) = 0
    B(3) = 0
    
    CopyMemory swapLong, B(0), 4
     
End Function



Private Sub Timer1_Timer()
    Timer1.Enabled = False
    CloseSocket
    err_msg = err_msg & " timeout"
End Sub

Private Sub ws_Close()
    Timer1.Enabled = False
End Sub

Private Sub ws_Connect()
    DoEvents
End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)
     CloseSocket
     ws.Accept requestID
     ws.SendData banner
End Sub
 
Private Sub ws_DataArrival(ByVal bytesTotal As Long)

    Dim tmp As String

    Timer1.Enabled = False
    Timer1.Enabled = True
    ws.GetData tmp
    data = data & tmp & vbCrLf

End Sub




'0 : eb  6                         jmp     short loc_8     \
'2 : eb  6                         jmp     short loc_A      |
'4 : 3c 12 1c 75                   dd 751C123Ch             |
'8 : 90                            nop                      |
'9 : 90                            nop                      |
'A : 90                            nop                      |--old sig
'B : 90                            nop                      |
'C : 90                            nop                      |
'D : 90                            nop                      |
'E : 90                            nop                      |
'F : 90                            nop                      |
'10 : eb 10                        jmp     short loc_22     |
'12 : 5a                           pop     edx              |
'13 : 4a                           dec     edx             /
'
'eb  6 eb  6 3c 12 1c 75 90 90 90 90 90 90 90 90 eb 10 5a 4a
'
'14 : 33 c9                         xor     ecx, ecx
'16 : 66 b9 7d  1                   mov     cx, 17Dh
'1A : 80 34  a 99                   xor     byte ptr [edx+ecx], 99h
'1E : e2 fa                         loop    loc_1A
'20 : eb  5                         jmp     short already_decoded
'22 : e8 eb ff ff ff                call    loc_12
'
'
'BB : 8b d8                         mov     ebx, eax
'BD : c7  7  2  0 2f  d             mov     dword ptr [edi], 0D2F0002h
'C3 : 33 c0                         xor     eax, eax
 
'0 : 90                            nop
'1 : 90                            nop
'2 : 90                            nop
'3 : 90                            nop
'4 : 90                            nop
'5 : 90                            nop
'6 : 90                            nop
'7 : eb  6                         jmp     short loc_F
'9 : eb  6                         jmp     short loc_11
'B : 3c 12 15 75                   dd 7515123Ch
'F : 90                            nop
'10 : 90                            nop
'11 : 90                            nop
'12 : 90                            nop
'13 : 90                            nop
'14 : 90                            nop
'15 : 90                            nop                               \
'16 : 90                            nop                                |
'17 : eb 10                         jmp     short loc_29               |
'19 : 5a                            pop     edx                        |
'1A : 4a                            dec     edx                        |
'1B : 33 c9                         xor     ecx, ecx                   | new sig
'1D : 66 b9 7d  1                   mov     cx, 17Dh                   |
'21 : 80 34  a 99                   xor     byte ptr [edx+ecx], 99h    |
'25 : e2 fa                         loop    loc_21                     |
'27 : eb  5                         jmp     short DECODED              |
'29 : e8 eb ff ff ff                call    loc_19                    /





