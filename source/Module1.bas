Attribute VB_Name = "Module1"
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

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "Shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function StringFromGUID2 Lib "OLE32.DLL" (lpGuid As UUID, ByVal lpszOut As String, ByVal cchMax As Long) As Long

Global hash As New CWinHash
Global fso As New clsFileSystem
Global ado As New clsAdoKit
Global BagleFtp As clsBagleFtpRecv

'global shellcode handler instances..
Global recv_cmd As New CLsassCmd
Global recv_file As New CLsassRecvFile
Global generic_url As New CGenericURL
Global sc_tftp As New CSc_tftp
Global vs_cmd As New CVeritasCmd
Global pnp_cmd As New CPnpCmd

Global logFile As String
Global FileDumpDir(8) As String
Global Const archiveDir As String = "C:\honeypot\archive\"

Enum oState
    aPreLogin
    bLoggedIn
    cPrepUpload
    dUploading
    eUploadComplete
    fTerminate
    gTimedOut
End Enum

Enum dumpDirs
    eMyDoom = 0
    eOptix = 1
    esub7 = 2
    eBagle = 3
    eKuang = 4
    EDCOM = 5
    eRPC445 = 6
    everitas = 7
End Enum

Enum SubSploits
    ssUnk = 0
    ssDcom = 1
    sslsass = 2
    ssASN = 3
    ssPNP = 4
End Enum

Type udt_config
    lsass As Byte
    kuang As Byte
    bagle As Byte
    doom As Byte
    sub7 As Byte
    optix As Byte
    AntiHammer As Byte
    banips As Byte
    lsass_port As Integer
    hammer_time As Integer
    veritas As Byte
    logLevel As Byte
    dcom_port As Integer
    dcom As Byte
End Type

Enum LogLevels
    llunk = 0
    llreal = 1
    llinfo = 2
    llspam = 3
End Enum

Private Type UUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

Type dce_packet
  version As Byte
  vMinor As Byte
  packetType As Byte
  packetFlags As Byte
  dataRepresentation As Long
  fragLength As Integer
  authlength As Integer
  callID As Long
  xmitFrag As Integer
  recvFrag As Integer
  assocGroup As Long
  ctxItems As Byte
  unknown(4) As Byte
  transitems As Integer
  interfaceGUID As UUID
  interfaceVersion As Integer
  iMinor As Integer
  transferSyntax As UUID
  syntaxVersion As Long
End Type

Global LogThreshHold As LogLevels
Global config As udt_config

Dim CurrentDay As Date

Sub LoadConfig()
    On Error GoTo hell
    
    Dim cfg As String, f As Long
    
    cfg = App.path & "\settings.cfg"
    f = FreeFile
    
    If Not fso.FileExists(cfg) Then
        With config
            .lsass = 1
            .kuang = 1
            .bagle = 1
            .doom = 1
            .sub7 = 1
            .optix = 1
            .AntiHammer = 1
            .banips = 0
            .lsass_port = 445
            .dcom_port = 135
            .hammer_time = 240
            .veritas = 1
            .logLevel = llspam
            .dcom = 1
        End With
        SaveConfig
    Else
        Open cfg For Binary As f
        Get f, , config
        CloseFile f
    End If
    
Exit Sub
hell: MsgBox "Could not load config? " & Err.Description
End Sub

Sub SaveConfig()
    On Error GoTo hell
    
    Dim cfg As String, f As Long
    cfg = App.path & "\settings.cfg"
    f = FreeFile
    
    Open cfg For Binary As f
    Put f, , config
    CloseFile f
    
    Exit Sub
hell: MsgBox "Could not save config? " & Err.Description
End Sub
    
'somewhere we are getting a stale handle being closed and it killing a live one
'cause vbs freefile() kinda sucks like that using sequential numbers
'close 0 doesnt hurt, so we hack over the bug with this...
'double close on stale handle is happening somewhere because we have to cache
'filehandles per form because winsock stuff has to be scattered across several events..
'trackign down exactly which one is responsible for a stale close...
Sub CloseFile(handle As Long)
    On Error Resume Next
    Close handle
    handle = 0
End Sub


Sub dbg(logLevel As LogLevels, ByVal msg)
       
    Dim isOk As Boolean
    
    'If logLevel = llunk Then isOk = True
    If LogThreshHold >= logLevel Then isOk = True
    
    If DateDiff("d", CurrentDay, Now) > 1 Then
        If Not fso.FolderExists(App.path & "\Logs") Then MkDir App.path & "\Logs"
        logFile = App.path & "\Logs\" & Format(Now, "mm.dd.yy.txt")
        CurrentDay = Now
        dbg llreal, "The Dawn of a new day: " & Format(Now, "mm-dd-yy") & "      Logging to: " & logFile
    End If
    
    msg = Format(Now, "hh:nn:ss") & "     " & msg
    
    If isOk Then 'log level only applies to form logging..
        With Form1.List1
            If .ListCount > 1000 Then .Clear
            .AddItem msg
        End With
    End If
    
    fso.AppendFile logFile, Replace(Replace(msg, vbCr, " <CR> "), vbLf, "<LF>")
    
End Sub

Function HPNameFromEnum(from As dumpDirs) As String
   Select Case from
        Case eBagle:   HPNameFromEnum = "Bagle"
        Case eMyDoom:  HPNameFromEnum = "MyDoom"
        Case eOptix:   HPNameFromEnum = "Optix"
        Case esub7:    HPNameFromEnum = "Sub7"
        Case eKuang:   HPNameFromEnum = "Kuang2"
        Case eRPC445:   HPNameFromEnum = "Rpc445"
        Case everitas: HPNameFromEnum = "Veritas"
        Case EDCOM:    HPNameFromEnum = "Dcom"
    End Select
End Function

Function SSNameFromEnum(ss As SubSploits) As String
   Select Case ss
        Case ssASN:     SSNameFromEnum = "ASN"
        Case ssDcom:    SSNameFromEnum = "DCOM"
        Case sslsass:   SSNameFromEnum = "LSASS"
        Case ssPNP:     SSNameFromEnum = "PNP"
    End Select
End Function

Sub ChangeNameForSize(selli As ListItem, lv As ListView, sizeIndex, nameIndex)
    
    Dim sql As String
    Dim rs As Recordset
    Dim newName As String
    
    On Error GoTo hell
    
    sql = "Select * from knownsizes where fsize=" & selli.SubItems(sizeIndex)
    
    newName = InputBox("Enter new name for " & selli.SubItems(nameIndex), , selli.SubItems(nameIndex))
    If Len(newName) = 0 Then Exit Sub
    
    ado.Update "knownsizes", "where fsize=" & selli.SubItems(sizeIndex), "malcode", newName
    
    Dim li As ListItem
    For Each li In lv.ListItems
        If li.SubItems(sizeIndex) = selli.SubItems(sizeIndex) Then li.SubItems(nameIndex) = newName
    Next
        
Exit Sub
hell:
        MsgBox Err.Description
        
End Sub

Function IsIde() As Boolean
    On Error GoTo hell
    Debug.Print 1 \ 0
Exit Function
hell: IsIde = True
End Function

Function InitDumpDirs()
    Dim i As Integer
    
    If Not fso.FolderExists("C:\honeypot") Then
        On Error Resume Next
        MkDir "c:\honeypot"
    End If
    
    FileDumpDir(0) = "C:\honeypot\MyDoom\"
    FileDumpDir(1) = "C:\honeypot\Optix\"
    FileDumpDir(2) = "C:\honeypot\Sub7\"
    FileDumpDir(3) = "C:\honeypot\Bagle\"
    FileDumpDir(4) = "C:\honeypot\Kuang\"
    FileDumpDir(5) = "C:\honeypot\DCOM\"
    FileDumpDir(6) = "C:\honeypot\RPC445\Shellcode\"
    FileDumpDir(7) = "C:\honeypot\Veritas\Shellcode\"
    
    For i = 0 To UBound(FileDumpDir)
        If Not fso.FolderExists(FileDumpDir(i)) Then
            fso.buildPath FileDumpDir(i)
        End If
    Next
    
    If Not fso.FolderExists(archiveDir) Then
        fso.buildPath archiveDir
    End If
    
End Function

Function GetDumpDir(dd As dumpDirs) As String
    If Len(FileDumpDir(0)) = 0 Then InitDumpDirs
    GetDumpDir = FileDumpDir(dd)
End Function

Function GetParentFolder(ByVal pth As String) As String
    Dim tmp() As String
    Dim x As String
    Dim org As String
    
    org = pth
    If Right(pth, 1) = "\" Then pth = Mid(pth, 1, Len(pth) - 1)
    
    On Error GoTo hell
    
    tmp = Split(pth, "\")
    x = tmp(UBound(tmp))
    pth = Replace(pth, x, Empty)
    pth = Replace(pth, "\\", "\")
    If Right(pth, 1) = "\" Then pth = Mid(pth, 1, Len(pth) - 1)
    
    GetParentFolder = pth
    
    
    
    Exit Function
hell:
    GetParentFolder = pth
End Function

Function GetFreeFileInDumpDir(dd As dumpDirs, Optional subDir As String = Empty) As String
    
    Dim f As String
    f = GetDumpDir(dd)
    
    If Len(subDir) > 0 Then
        If InStr(subDir, "./../") > 0 Then
            f = GetParentFolder(f)
            subDir = Replace(subDir, "./../", Empty)
        End If
        If Len(subDir) > 0 Then f = f & IIf(Left(subDir, 1) = "\", "", "\") & subDir
        If Not fso.FolderExists(f) Then
            If Not fso.buildPath(f) Then
                dbg llreal, "!! GetFreeFileInDumpDir Failed to build path: " & f
                f = GetDumpDir(dd)
            End If
        End If
    End If
    
    If Right(f, 1) <> "\" Then f = f & "\"
    
    f = fso.GetFreeFileName(f, ".dat")
    GetFreeFileInDumpDir = f
    
End Function

Function MoveFileToDumpDir(filepath As String, dd As dumpDirs, Optional subDir As String = Empty) As String
    
    Dim f As String
    f = GetDumpDir(dd)
    
    If Len(subDir) > 0 Then
        If InStr(subDir, "./../") > 0 Then
            f = GetParentFolder(f)
            subDir = Replace(subDir, "./../", Empty)
        End If
        If Len(subDir) > 0 Then f = f & IIf(Left(subDir, 1) = "\", "", "\") & subDir
        If Not fso.FolderExists(f) Then
            If Not fso.buildPath(f) Then
                dbg llreal, "!! MoveFileToDumpDir Failed to build path: " & f
                f = GetDumpDir(dd)
            End If
        End If
    End If
    
    If Right(f, 1) <> "\" Then f = f & "\"
    
    Dim newPath As String
    newPath = f & fso.FileNameFromPath(filepath)
    
    If fso.FileExists(newPath) Then newPath = f & fso.GetFreeFileName(f, ".dat")
    
    On Error GoTo hell
    Name filepath As newPath
    MoveFileToDumpDir = newPath
    
hell:
    
End Function


Function safeGetTmpFile(dd As dumpDirs) As String
    On Error GoTo ihavebugs
    Dim fails As Integer
    
    If Len(FileDumpDir(0)) = 0 Then InitDumpDirs
    
tryAgain:
    safeGetTmpFile = fso.GetFreeFileName(FileDumpDir(dd), ".dat")
    
Exit Function
ihavebugs:
      fails = fails + 1
      If fails > 30 Then
        Exit Function
      Else
        GoTo tryAgain
      End If
    
End Function

Function CountOccurances(it, find) As Integer
    Dim tmp() As String
    If InStr(1, it, find, vbTextCompare) < 1 Then CountOccurances = 0: Exit Function
    tmp = Split(it, find, , vbTextCompare)
    CountOccurances = UBound(tmp)
End Function



Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Integer
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Function GetSelIndex(l As ListBox) As Long
    On Error Resume Next
    Dim i As Long
    For i = 0 To l.ListCount
        If l.selected(i) = True Then
            GetSelIndex = i
            Exit Function
        End If
    Next
End Function

Function IPinList(ip As String, list As ListBox) As Boolean
    On Error Resume Next
    Dim i As Long
    For i = 0 To list.ListCount
        If ip Like list.list(i) Then
            IPinList = True
            Exit Function
        End If
    Next
End Function

Function GetAllElements(lv As ListView) As String
    Dim ret() As String, i As Integer, tmp As String
    Dim li As ListItem

    For i = 1 To lv.ColumnHeaders.Count
        tmp = tmp & lv.ColumnHeaders(i).Text & vbTab
    Next

    push ret, tmp
    push ret, String(50, "-")

    For Each li In lv.ListItems
        tmp = li.Text & vbTab
        For i = 1 To lv.ColumnHeaders.Count - 1
            tmp = tmp & li.SubItems(i) & vbTab
        Next
        push ret, tmp
    Next

    GetAllElements = Join(ret, vbCrLf)

End Function

Function KeyExistsInCollection(C As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = C(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function HexDump(it)
    Dim my, i, C, s, A, B
    Dim lines() As String
    
    my = ""
    For i = 1 To Len(it)
        A = Asc(Mid(it, i, 1))
        C = Hex(A)
        C = IIf(Len(C) = 1, "0" & C, C)
        B = B & IIf(A > 65 And A < 120, Chr(A), ".")
        my = my & C & " "
        If i Mod 16 = 0 Then
            push lines(), my & "  [" & B & "]"
            my = Empty
            B = Empty
        End If
    Next
    
    If Len(B) > 0 Then
        If Len(my) < 48 Then
            my = my & String(48 - Len(my), " ")
        End If
        If Len(B) < 16 Then
             B = B & String(16 - Len(B), " ")
        End If
        push lines(), my & "  [" & B & "]"
    End If
        
    If Len(it) < 16 Then
        HexDump = my & "  [" & B & "]" & vbCrLf
    Else
        HexDump = Join(lines, vbCrLf)
    End If
    
    
End Function



Function toHex(str) As String
        
    Dim tmp() As String, i As Long
    tmp = Split(str, " ")
    For i = 0 To UBound(tmp)
        tmp(i) = Chr(CInt("&h" & tmp(i)))
    Next
    
    toHex = Join(tmp, "")
        
End Function



Function AnyOfTheseInstr(sIn, sCmp) As Boolean
    Dim tmp() As String, i As Integer
    tmp() = Split(sCmp, ",")
    For i = 0 To UBound(tmp)
        tmp(i) = Trim(tmp(i))
        If Len(tmp(i)) > 0 And InStr(1, sIn, tmp(i), vbTextCompare) > 0 Then
            AnyOfTheseInstr = True
            Exit Function
        End If
    Next
End Function

Function AllOfTheseInstr(sIn, sCmp) As Boolean
    Dim tmp() As String, i As Integer
    tmp() = Split(sCmp, ",")
    For i = 0 To UBound(tmp)
        tmp(i) = Trim(tmp(i))
        If Len(tmp(i)) > 0 And InStr(1, sIn, tmp(i), vbTextCompare) < 1 Then
            Exit Function
        End If
    Next
    AllOfTheseInstr = True
End Function

Public Function UserDeskTopFolder() As String
    Dim idl As Long
    Dim p As String
    Const MAX_PATH As Long = 260
      
      p = String(MAX_PATH, Chr(0))
      If SHGetSpecialFolderLocation(0, 0, idl) <> 0 Then Exit Function
      SHGetPathFromIDList idl, p
      
      UserDeskTopFolder = Left(p, InStr(p, Chr(0)) - 1)
      CoTaskMemFree idl
  
End Function

Sub RunWhois(ip As String)
    On Error GoTo hell
    Shell "cmd /k whois " & ip, vbNormalFocus
    Exit Sub
hell: MsgBox "Could not spawm cmd.exe with whois command line. Feature depends on NT based OS with Whois.exe installed.", vbInformation
End Sub

Sub CloseSocket(s As Winsock)
    On Error Resume Next
    s.Close
End Sub

Function StringFromGUID(iid As UUID) As String
    Dim tmp As String, n As Long
    tmp = Space(80)
    StringFromGUID2 iid, tmp, 80
    n = StringFromGUID2(iid, tmp, 80)
    If n > 0 Then
        StringFromGUID = Left$(StrConv(tmp, vbFromUnicode), n - 1)
    End If
End Function

Function SubSploitFromIID(GUID As String) As SubSploits
    
    Select Case GUID
        Case "{8D9F4E40-A03D-11CE-8F69-08003E30051B}": SubSploitFromIID = ssPNP
        Case "{3919286A-B10C-11D0-9BA8-00C04FD92EF5}": SubSploitFromIID = sslsass
    End Select
    
End Function
