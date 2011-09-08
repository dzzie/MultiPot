VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form CPnpCmd 
   Caption         =   "CPnpCmd"
   ClientHeight    =   465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1875
   LinkTopic       =   "CPnpCmd"
   ScaleHeight     =   465
   ScaleWidth      =   1875
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   420
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
Attribute VB_Name = "CPnpCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'this form handles a shellcode that was recieved over PNP port
'that exports a shell back to the infector to receive a command

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
Private start_signature2 As String

'Private banner As String
Private state As Long
Private data  As String

Public err_msg As String
Public Busy As Boolean
Public Success_FileName As String

Dim emu As New CCmdEmulator

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
    
    'banner = "Microsoft Windows 2000 [Version 5.00.2195]" & vbCrLf & _
    '         "(C) Copyright 1985-2000 Microsoft Corp." & vbCrLf & _
    '         vbCrLf & _
    '         "C:\WinNT\System32>"
      
    'd9 ee                  fldz
    'd9 74 24 f4            fnstenv byte ptr [esp-0Ch]
    '5b                     pop     ebx
    '81 73 13 19 f5  4 37   xor     dword ptr [ebx+13h], 3704F519h
    '83 eb fc               sub     ebx, 0FFFFFFFCh
    'e2 f4                  loop    loc_D

    start_signature = toHex("d9 ee d9 74 24 f4 5b 81 73 13")
    
    
    '5a                            pop     edx
    '4a                            dec     edx
    '33 c9                         xor     ecx, ecx
    '66 b9 7d  1                   mov     cx, 17Dh
    '80 34  a 99                   xor     byte ptr [edx+ecx], 99h
    'e2 fa                         loop    loc_11
    'eb  5                         jmp     short loc_1E
    'e8 eb ff ff ff                call    sub_9

     start_signature2 = toHex("5a 4a 33 c9 66 b9 7d 1 80 34 a 99 e2 fa eb 5 e8 eb ff ff ff")
    
    
    'Dim buf As String
    'buf = ReadFile(App.path & "\PNP.sc")
    'Me.HandleShellcode "127.0.0.1", "", buf
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
End Sub

Function CheckSignature(ByVal MemBuffer As String) As Boolean
    
    If Len(start_signature) = 0 Then Form_Load
    
    Dim tmp As String, start As Long
    tmp = Replace(MemBuffer, Chr(0), "")
    start = InStrRev(tmp, start_signature)
   
    If start > 1 Then
        If InStr(start, tmp, String(6, Chr(&H90))) < 1 Then Exit Function
        CheckSignature = True
    End If
    
    start = InStrRev(tmp, start_signature2)
   
    If start > 1 Then
        If InStr(start, tmp, String(6, Chr(&H90))) < 1 Then Exit Function
        CheckSignature = True
    End If
    
    
End Function

Function HandleShellcode(remoteHost As String, filepath As String, MemBuffer As String) As Boolean
     
   Dim ftp As CFtpGet
   Dim tftp As CTFTPClient
   Dim myFile As String
   Dim ip As String, File As String
   Dim ret As Boolean
   Dim mode As dataMode
   
   On Error Resume Next
   
   emu.Reset
   Busy = True
   Success_FileName = Empty
   mode = mAuto
                  
   If Len(start_signature) = 0 Then Form_Load 'stupid weird bug!
                  
   ret = processfile(MemBuffer)
   'MoveFileToDumpDir filepath, ePNP, "\recv_cmd"
   dbg llinfo, "Parsing PNPCmd Received: " & data
   
   If Not ret Then
        dbg llinfo, "Failed to process file in PNPCmd handler Error:" & err_msg
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
    Dim buf() As Byte
    Dim method As Integer
    
    tmp = Replace(memBuf, Chr(0), Empty)
    start = InStrRev(tmp, start_signature)
   
    If start < 1 Then
        start = InStrRev(tmp, start_signature2)
        If start < 1 Then
            err_msg = "Neither Signature not found!"
            Exit Function
        Else
            method = 1
        End If
    End If
    
    endat = InStr(start, tmp, String(6, Chr(&H90)))
    
    If endat < 1 Then
        err_msg = "End sig not found shellcode truncated?"
        Exit Function
    End If
    
    tmp = Mid(tmp, start, endat - start)
    buf = StrConv(tmp, vbFromUnicode)
    
    If UBound(buf) < &HBD Then 'minimum length acceptable
        err_msg = "Invalid shellcode length?"
        Exit Function
    End If
    
    Dim port As Long
    Dim host As String
    Dim push As Byte
    'mabey we should check for the push..
    '66 68 22 b8                   push    small 0B822h    ; port 8888 hex(22b8)

    If method = 0 Then
        push = buf(&HB3) Xor buf(10)
        
        If push <> &H66 Then
            dbg llreal, "PnpCmd: Push port offset may not be correct manually examine this payload!"
        End If
        
        port = swapLong(buf(&HB5) Xor buf(12), buf(&HB6) Xor buf(13))
    Else
        '
        'c7  7  2  0 11 6c   mov dword ptr [edi], 6C110002h; 0x116c = 4460 (port)
        'xor key is believed static for this one (its in signature too)
        port = swapLong(buf(&HAE) Xor &H99, buf(&HAF) Xor &H99)
    
    End If

    On Error GoTo hell
    state = 0
    err_msg = Empty
    data = Empty
    CloseSocket
    ws.LocalPort = port
    ws.Listen
    Timer1.Enabled = True
    
    While Timer1.Enabled
        DoEvents
        Sleep 30
    Wend
    
    If Len(data) = Empty Then
        err_msg = "Method: " & method & " Started shell on " & port & " no data received: " & err_msg
    Else
        err_msg = "Method: " & method & " port: " & port & data
        processfile = True
    End If
    
    
    
 Exit Function
hell:  err_msg = "Error line: " & Erl & " Desc: " & Err.Description
End Function

Private Function CloseSocket()
    On Error Resume Next
    ws.Close
End Function

Private Function swapLong(B1 As Byte, b2 As Byte) As Long
    Dim B(3) As Byte

    B(0) = b2
    B(1) = B1
    B(2) = 0
    B(3) = 0
    
    CopyMemory swapLong, B(0), 4
     
End Function

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    CloseSocket
    err_msg = err_msg & " timeout port:" & ws.LocalPort
End Sub

Private Sub ws_Close()
    Timer1.Enabled = False
End Sub

Private Sub ws_Connect()

    On Error Resume Next
    ws.SendData emu.GetBanner
    
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)

    Dim tmp As String
    On Error Resume Next
    
    Timer1.Enabled = False
    Timer1.Enabled = True
    ws.GetData tmp
    data = data & tmp & vbCrLf
    
    ws.SendData emu.GetResponse(tmp)

End Sub



Function ReadFile(filename)
  Dim f As Long, temp As String
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

'start_signature
'9 : eb 6e                         jmp     short loc_79
'B : 33 c0                         xor     eax, eax
'D : 64 8b 40 30                   mov     eax, fs:[eax+30h]
'11 : 85 c0                         test    eax, eax
'13 : 78  d                         js      short loc_22
 




 
Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    err_msg = err_msg & "Ws Error: " & Description
End Sub
