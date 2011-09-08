VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form CTFTPClient 
   BorderStyle     =   0  'None
   Caption         =   "TFTPClient"
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   465
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer tmrReAck 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   840
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   18000
      Left            =   0
      Top             =   0
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   420
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   69
   End
End
Attribute VB_Name = "CTFTPClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'minimal implementation of a tftp client to download bots


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

Public ErrorMsg As String
Public HadError As Boolean
Public TimedOut As Boolean
Public LocalFilename As String
Public remoteHost As String
Public RemotePort As String
Public RemoteFileName As String
Public EnforceMaxFileSize As Boolean

Private fHand As Long
Private readyToReturn As Boolean
Private currentBlock As Long
Private Const MAX_FILE_SIZE As Long = 3000000 '~3mb

Property Get RealFileSize() As Long
    On Error Resume Next
    RealFileSize = FileLen(LocalFilename)
End Property


Private Sub SendFileRequest(ByVal filename As String)
    
    Dim B() As Byte
    Dim buf As String
    
    buf = Chr(0) & Chr(1) & filename & Chr(0) & "octet" & Chr(0)
    B = StrConv(buf, vbFromUnicode)
        
    On Error Resume Next
    ws.SendData B()
    
    Debug.Print "Filereq for " & filename & " sent"
    
End Sub

Private Sub AckBlock(Block As Integer)
        
    Dim B() As Byte
    Dim buf As String
     
    buf = Chr(0) & Chr(4) & Space(22)
    B() = StrConv(buf, vbFromUnicode)
        
    Dim tmp() As Byte
    tmp = IntToSwap(Block)
    CopyMemory B(2), tmp(0), 2
    
    On Error Resume Next
    ws.SendData B()
        
    ResetTimer
    
    If Block Mod 10 = 0 Then
        Debug.Print "Tftp Ack sent for block: " & Block & " filesize:" & Hex(LOF(fHand))
    End If
    
End Sub

Private Sub SendErrorPacket(Optional msg As String = "Transfer Error")
        
    Dim B() As Byte
    Dim buf As String
     
    '         error opcode   user defined
    buf = Chr(0) & Chr(5) & Chr(0) & Chr(0) & msg & Chr(0)
    B() = StrConv(buf, vbFromUnicode)
    
    On Error Resume Next
    ws.SendData B()
    
    
End Sub

Function ParseTftpCmd(ByVal cmdString, remoteip, server, File) As Boolean
    
    On Error GoTo hell
    Dim tmp() As String
    
    If Len(cmdString) = 0 Then Exit Function
     
    If InStr(cmdString, "0.0.0.0") > 0 Then cmdString = Replace(cmdString, "0.0.0.0", remoteip)
        
    tmp = Split(cmdString, " ")
    server = tmp(2)
    File = tmp(4)
        
    If InStr(File, "&") > 0 Then File = Mid(File, 1, InStr(File, "&") - 1)
        
    'this only handles ip addresses and not accout for domain names..
    'If CountOccurances(server, ".") <> 3 Then Exit Function
    
    If CountOccurances(server, ".") < 1 Then Exit Function
    
    ParseTftpCmd = True
    
hell:
End Function

Function GetFile(filename As String, localFile As String, server As String, Optional port As Integer = 69) As Boolean

On Error GoTo hell

    RemoteFileName = filename
    LocalFilename = localFile
    RemotePort = port
    remoteHost = server
    
    ErrorMsg = Empty
    HadError = False
    TimedOut = False
    EnforceMaxFileSize = True
     
    readyToReturn = False
    currentBlock = 1
    
    ws.remoteHost = server
    ws.RemotePort = RemotePort
    
    fHand = FreeFile
    Open LocalFilename For Binary As fHand
    
    Timer1.Enabled = True
    SendFileRequest RemoteFileName
    
    Do While Not readyToReturn
        DoEvents
        Sleep 60
        If TimedOut Then Exit Do
        If HadError Then Exit Do
    Loop
    
    On Error Resume Next
    Timer1.Enabled = False
    tmrReAck.Enabled = False
    If TimedOut Then SendErrorPacket
    CloseFile fHand
    
    If Not TimedOut And Not HadError Then
        GetFile = True
    Else
        CloseFile fHand
        Kill LocalFilename
    End If
    
    ws.Close
    
    Exit Function
hell: ErrorMsg = ErrorMsg & " & tftp.getfile error: " & Err.Description
End Function

Private Sub Timer1_Timer()
    TimedOut = True
    Timer1.Enabled = False
    Debug.Print "Timeout! " & Format(Now, "h:n:s")
End Sub

Private Sub tmrReAck_Timer()
    AckBlock CInt(currentBlock - 1)
    'Debug.Print "tmrReAck " & (currentBlock - 1)
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    
    On Error GoTo hell
    
    Dim B() As Byte
    Dim data() As Byte
    Dim opcode As Integer
    Dim Block As Integer
    Dim dataLength As Long
    Dim sz As Long
    
    ws.GetData B, bytesTotal
    
    opcode = SwapToInt(B(0), B(1))
    Block = SwapToInt(B(2), B(3))
    
    tmrReAck.Enabled = False
    tmrReAck.Enabled = True
    
    Select Case opcode
        Case 3 'data packet
                If Block = currentBlock Then
                
                    currentBlock = currentBlock + 1
                    dataLength = UBound(B) - 4
                    
                    If dataLength >= 0 Then
                        
                        If fHand < 1 Then
                            ErrorMsg = "Invalid file handle"
                            Exit Sub
                        End If
                        
                        ReDim data(dataLength)
                        CopyMemory data(0), B(4), UBound(data) + 1
                        
                        If EnforceMaxFileSize Then
                            sz = RealFileSize()
                             If sz > MAX_FILE_SIZE Then
                                dbg llreal, "TFTP REACHED MAX FILE SIZE CLOSING FILE EARLY " & LocalFilename
                                CloseFile fHand
                                HadError = True
                                CloseSocket ws
                                Exit Sub
                            End If
                        End If
                        
                        Put fHand, , data()
                        AckBlock Block
                        
                    End If
                    
                    If dataLength < 511 Then
                        readyToReturn = True
                        CloseFile fHand
                        fHand = 0
                    End If
                    
                Else
                    AckBlock Block
                    'Debug.Print "Had to Reack block " & Block
                End If
        Case 5 'error
                dataLength = bytesTotal - 5
                ReDim data(dataLength)
                CopyMemory data(0), B(4), (dataLength + 1)
                ErrorMsg = StrConv(data, vbUnicode)
                AckBlock Block
                HadError = True
                
    End Select
    
                
 Exit Sub
hell:
    ErrorMsg = "CTftpClient.DataArrival Error: " & Err.Description & " Localfile: " & LocalFilename
    HadError = True
End Sub

Private Sub ResetTimer()
    Timer1.Enabled = False '\_reset interval
    DoEvents: Sleep 10     ' |
    Timer1.Enabled = True  '/
    'Debug.Print "Reset " & Format(Now, "h:n:s")
End Sub
 
Private Function SwapToInt(b1 As Byte, b2 As Byte) As Integer
    Dim B(1 To 2) As Byte
    B(1) = b2
    B(2) = b1
    CopyMemory SwapToInt, B(1), 2
End Function

Private Function IntToSwap(i As Integer) As Byte()
    Dim B(1 To 2) As Byte
    Dim ret(0 To 1) As Byte
    CopyMemory B(1), i, 2
    ret(0) = B(2)
    ret(1) = B(1)
    IntToSwap = ret
End Function
    
