VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form clsBagleDownloader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bagle Exploit FTP Bot Downloader"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLog 
      Height          =   2595
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   300
      Width           =   5775
   End
   Begin VB.Timer tmrBullshit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5340
      Top             =   0
   End
   Begin MSWinsockLib.Winsock wsControl 
      Left            =   4380
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4860
      Top             =   0
   End
   Begin VB.Label Label3 
      Caption         =   "Connect Back port : 12345 (make sure reachable from net)"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "clsBagleDownloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
Public fsize As Long
Dim ftp As ftpConnectionStats
Dim myIP As String

Dim readyToReturn As Boolean
Dim haveSize As Boolean
Dim TimedOut As Boolean
Dim onlyGetSize As Boolean
Dim myFtpIndex As Integer
Public myFileName As String

Private WithEvents sckFTP As clsBagleFtpRecv
Attribute sckFTP.VB_VarHelpID = -1

Private Type ftpConnectionStats
    folder As String
    server As String
    port As Integer
    user As String
    pass As String
End Type

Function LoadFTPUrl(URL As String) As Boolean
    On Error GoTo hell
    ftp = ParseFTPString(URL)
    
    'dbg "Port: " & ftp.port & vbCrLf & _
            "Server: " & ftp.server & vbCrLf & _
            "User: " & ftp.user & vbCrLf & _
            "Pass: " & ftp.pass & vbCrLf & _
            "Folder:" & ftp.folder
    
    myIP = Replace(PublicIP, ".", ",")
    
    'If Len(PublicIP) = 0 Then MsgBox "Set Public IP addres !": Err.Raise 1
    
    LoadFTPUrl = True
    Exit Function
hell:
    LoadFTPUrl = False
End Function
 

 

Function GetInfectionSize() As Long
On Error GoTo hell

    haveSize = False
    TimedOut = False
    onlyGetSize = True
    fsize = 0
    
    Set sckFTP = BagleFtp
    
    With wsControl
        .Close
        .RemotePort = ftp.port
        .remoteHost = ftp.server
        .Connect
        tmrBullshit.Enabled = True
        tmrTimeout.Enabled = True
    End With

    While Not haveSize
        DoEvents
        If TimedOut Then Exit Function
        Sleep 2
    Wend
    
hell:
    tmrTimeout.Enabled = False
    
    GetInfectionSize = fsize
    
    If Err.Number > 0 Then MsgBox "BagleDld GetSize: " & Err.Description
    
End Function
Function DoBagleDownload() As Boolean
    
    On Error GoTo hell
    
    readyToReturn = False
    TimedOut = False
    onlyGetSize = False
    
    Set sckFTP = BagleFtp
    
    With wsControl
        .Close
        .RemotePort = ftp.port
        .remoteHost = ftp.server
        .Connect
        tmrBullshit.Enabled = True
        tmrTimeout.Enabled = True
    End With

    While Not readyToReturn
        DoEvents
        If TimedOut Then Exit Function
        Sleep 2
    Wend
    
    tmrTimeout.Enabled = False
    
    DoBagleDownload = True

    
Exit Function
hell:
    MsgBox "BagleDl Download: " & Err.Description
    tmrTimeout.Enabled = False
End Function

Private Sub sckFTP_Connect(ip As String)
    'dbg "FtpRecvConnect"
End Sub

Private Sub sckFTP_DataRecv(Index As Integer, size As Long)

    If myFtpIndex = 0 Then
        If sckFTP.remoteHost(Index) = ftp.server Then
            myFtpIndex = Index
            myFileName = sckFTP.filename(Index)
        End If
    End If
    
    If Index = myFtpIndex Then
        'dbg "Revc : & size"
    End If
    
End Sub

Private Sub sckFTP_TimeOut(Index As Integer)
    'dbg "sckFtp timeout"
    If Index = myFtpIndex Then TimedOut = True
End Sub

Private Sub tmrBullshit_Timer()
    On Error Resume Next
    
    If wsControl.BytesReceived > 0 Then
        'dbg "Stall recover"
        wsControl_DataArrival wsControl.BytesReceived
    End If
    
End Sub

Private Sub tmrTimeout_Timer()
    On Error Resume Next
    wsControl.Close
    tmrTimeout.Enabled = False
    TimedOut = True
End Sub

Private Sub wsControl_DataArrival(ByVal bytesTotal As Long)

    Dim tmp As String
    wsControl.GetData tmp, vbString
    'dbg "Control data:>  " & tmp
    
    tmrBullshit.Enabled = False '\_ . reset
    tmrBullshit.Enabled = True  '/
    
    tmrTimeout.Enabled = False '\_ . reset
    tmrTimeout.Enabled = True  '/
    
    With wsControl
        Select Case Left(tmp, 3)
            
            Case "220":  '220 Bot Server (Win32)
                        .SendData "PASS " & ftp.pass & vbCrLf
                        'dbg "Sending Pass"
                        
            Case "230": '230 Login successful. Have fun.
                       .SendData "SIZE" & vbCrLf
                       'dbg "asking size"
            
            Case "213": '213 294912
                      On Error Resume Next
                      fsize = CLng(Split(tmp, " ")(1))
                      haveSize = True
                      'tmp = "PORT " & myIP & ",48,57" & vbCrLf '"PORT 10,10,10,7,48,57"
                      tmp = "PORT " & Replace(wsControl.LocalIP, ".", ",") & ",48,57" & vbCrLf '"PORT 10,10,10,7,48,57"
                      
                      If onlyGetSize Then
                            On Error Resume Next
                            wsControl.Close
                            Exit Sub
                      End If
                      
                      .SendData tmp
                      'dbg "size is: " & fsize & " setting port"
                      'dbg tmp
                      
            
            Case "200" '200 PORT command successful
                       .SendData "RETR" & vbCrLf
                        'dbg "sending retr command"
                        
            Case "150" '150 Opening BINARY mode data connection.
            
                        tmrTimeout.Enabled = False 'defer to ftp steam timeout
            
            Case "226" '226 Transfer Complete
                        
                    'dbg "Transfer Complete Closing ftpfile for index: " & myFtpIndex
                    sckFTP.CloseFile myFtpIndex
                    readyToReturn = True
                    
        End Select
    
    End With
    
        
    DoEvents


End Sub
 

Private Sub wsControl_SendComplete()
    DoEvents
End Sub

'Sub dbg(msg)
'    txtLog.SelStart = Len(txtLog)
'    txtLog.SelText = msg & vbCrLf
'    Debug.Print msg
'End Sub





Private Function ParseFTPString(x) As ftpConnectionStats
    Dim at, authinfo, serverinfo, tmp, slash, sc, k
    Dim C As ftpConnectionStats
    x = Replace(x, "ftp://", "")
    at = InStrRev(x, "@")
    If at > 0 Then 'is user:pass@server:port type login
       authinfo = Mid(x, 1, at - 1)
       serverinfo = Mid(x, at + 1)
       tmp = Split(authinfo, ":")
       C.user = tmp(0)
       C.pass = tmp(1)
       If InStr(serverinfo, ":") > 0 Then
           tmp = Split(serverinfo, ":")
           C.server = tmp(0)
           
           
           If InStr(tmp(1), "/") > 0 Then
                k = Split(tmp(1), "/")
                C.port = k(0)
                C.folder = k(UBound(k))
           Else
                C.port = tmp(1)
           End If
       Else
           C.server = serverinfo
           C.port = 21
       End If
    Else 'is anonymous login type url
       C.user = "anonymous"
       C.pass = "someone@somewhere.com"
       slash = InStr(x, "/")
       sc = InStr(x, ":")
       If sc > 0 Then
          C.server = Mid(x, 1, sc - 1)
          If slash > 0 Then
                C.port = Mid(x, sc + 1, (slash - 1) - sc)
                C.folder = Mid(x, slash, Len(x))
          Else
                C.port = Mid(x, sc + 1, Len(x))
          End If
       Else
          C.port = 21
          If slash > 0 Then
                C.server = Mid(x, 1, slash - 1)
                C.folder = Mid(x, slash, Len(x))
          Else
                C.server = x
          End If
       End If
    End If
    
    If C.folder = "/" Then C.folder = Empty
    If Right(C.folder, 1) = "/" Then C.folder = Mid(C.folder, 1, Len(C.folder) - 1)
    
    'MsgBox "Port: " & c.port & vbCrLf & _
    '        "Server: " & c.server & vbCrLf & _
    '        "User: " & c.user & vbCrLf & _
    '        "Pass: " & c.pass & vbCrLf & _
    '        "Folder:" & c.folder
    
    ParseFTPString = C
End Function

Function FileExists(path) As Boolean
If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function



