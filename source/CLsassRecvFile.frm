VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form CLsassRecvFile 
   Caption         =   "CLsassRecvFile_2"
   ClientHeight    =   420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   LinkTopic       =   "Form2"
   ScaleHeight     =   420
   ScaleWidth      =   1560
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
Attribute VB_Name = "CLsassRecvFile"
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
Dim start_signature As String
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public err_msg As String
Public Busy As Boolean
Public SampleFile As String
Public RemoteHostIp As String

Private state As Long
Private fHand As Long
Private readyToReturn As Boolean
Private HadError As Boolean

Function BlockWhileBusy()
    
    While Busy
        DoEvents
        Sleep 30
    Wend
    
End Function

Sub Form_Load()
     
    Timer1.Interval = 8000
    
    'original slightly to specific, dd 7515123c value can change
    'start_signature = "90 eb 6 eb 6 3c 12 15 75 90 90 90 90 90 90 90 90 eb 2 eb 6b e8 f9 ff ff ff 53 55"
    
    start_signature = "90 90 90 90 90 90 90 90 eb 2 eb 6b e8 f9 ff ff ff 53 55 56 57 8b 6c 24 18"
    
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
     
   On Error Resume Next
   
   Busy = True
   
   If Len(start_signature) = 0 Then Form_Load
   MoveFileToDumpDir filepath, eRPC445, "recv_file"
   HandleShellcode = processfile(MemBuffer)
   
   Busy = False
   
End Function


Private Function processfile(ByVal memBuf As String) As Boolean
    Dim tmp As String
    Dim start As Long, endat As Long
    Dim buf() As Byte

    tmp = Replace(memBuf, Chr(0), Empty)
    
    start = InStrRev(tmp, start_signature)
    start = start - 9 'back to where original signature woudl start
    
    If start < 1 Then
        err_msg = "Signature not found!"
        Exit Function
    End If
    
    If (start + 20) > Len(tmp) Then
        err_msg = "Shellcode Truncated"
        Exit Function
    End If
    
    endat = InStr(start + 20, tmp, String(6, Chr(&H90)))
    
    If endat < 1 Then
        err_msg = "End sig not found shellcode truncated?"
        Exit Function
    End If
    
    tmp = Mid(tmp, start, endat - start)
    buf = StrConv(tmp, vbFromUnicode)
    
    If UBound(buf) < &H10D Then
        err_msg = "Shellcode Truncated?"
        Exit Function
    End If
    
    If buf(&H10B) <> &HB8 Then
        err_msg = "mov offset incorrect?"
        Exit Function
    End If
    
    Dim port As Long
    port = swapLong(buf(&H10E), buf(&H10F))
    
    err_msg = Empty
    readyToReturn = False
    HadError = False
    fHand = FreeFile
    state = 0
    RemoteHostIp = ""
    SampleFile = GetFreeFileInDumpDir(eRPC445, "./../")
    
    On Error Resume Next
    ws.Close
    ws.LocalPort = port
    ws.Listen
    Timer1.Enabled = True
     
    While Not readyToReturn
        DoEvents
        Sleep 30
    Wend
    
    CloseFile fHand
    
    If HadError Then
        err_msg = "port: " & port & " Failed! " & err_msg
    Else
        err_msg = "port: " & port & " saved as " & SampleFile
        processfile = True
    End If
    
    
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
    readyToReturn = True
    HadError = True
    err_msg = err_msg & " timeout"
End Sub

Private Sub ws_Close()
    Timer1.Enabled = False
    readyToReturn = True
End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)
     ws.Close
     ws.Accept requestID
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    
    On Error Resume Next
    Dim data() As Byte
    
    Timer1.Enabled = False
    Timer1.Enabled = True
        
    ws.GetData data
    
    If state = 0 Then
        Open SampleFile For Binary As fHand
        RemoteHostIp = ws.RemoteHostIp
        state = 1
        ws.SendData 4
    End If
    
    Put fHand, , data
    
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    readyToReturn = True
    HadError = True
End Sub




'A : 90                            nop
'B : 90                            nop
'C : 90                            nop
'D : 90                            nop
'E : 90                            nop
'F : 90                            nop
'10 : 90                            nop
'11 : 90                            nop
'12 : eb  2                         jmp     short loc_16
'14 : eb 6b                         jmp     short sub_81
'16 : e8 f9 ff ff ff                call    loc_14
'1B : 53                            push    ebx
'1C : 55                            push    ebp
'1D : 56                            push    esi
'1E : 57                            push    edi
'1F : 8b 6c 24 18                   mov     ebp, [esp+arg_4]

'90 90 90 90 90 90 90 90 eb 2 eb 6b e8 f9 ff ff ff 53 55 56 57 8b 6c 24 18



'0 : eb  6                         jmp     short loc_8
'2 : eb  6                         jmp     short loc_A
'4 : 3c 12 1c 75                   dd 751C123Ch
'8 : 90                            nop
'9 : 90                            nop
'A : 90                            nop
'B : 90                            nop
'C : 90                            nop
'D : 90                            nop
'E : 90                            nop
'F : 90                            nop
'10 : eb 10                        jmp     short loc_22
'12 : 5a                           pop     edx
'13 : 4a                           dec     edx
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


