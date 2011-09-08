VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCmdEmulator"
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

'this is used from veritas cmd so far..it may be used in others
'but this one has enough interaction to capture multiple cmds
'and possible even make an individual scratch their head for a
'little while. (1 and counting ;)

Const banner = "Microsoft Windows 2000 [Version 5.00.2195]" & vbCrLf & _
                "(C) Copyright 1985-2000 Microsoft Corp." & vbCrLf & _
                "" & vbCrLf & _
                "C:\WINNT\System32>"

Dim pwd As String

Const cList = " Volume in drive C has no label." & vbCrLf & _
                " Volume Serial Number is 48EB-E0E0" & vbCrLf & _
                "" & vbCrLf & _
                " Directory of [PWD]" & vbCrLf & _
                "" & vbCrLf & _
                "[DIRTABLE]" & vbCrLf & _
                "" & vbCrLf & _
                "[PWD]\>" & vbCrLf

Const dTable = "10/21/2004  01:21p      <DIR>          ." & vbCrLf & _
                "10/21/2004  01:21p      <DIR>          .." & vbCrLf & _
                "               0 File(s)              0 bytes" & vbCrLf & _
                "               2 Dir(s)   1,418,227,712 bytes free"

Const ha = "[ARG0] is not recognized as an internal or external command," & vbCrLf & _
           "operable program or batch file." & vbCrLf & vbCrLf & _
           "[PWD]>"

Const nodel = "Could Not Find [ARG0] " & vbCrLf & _
                "" & vbCrLf & _
                "[PWD]>"

Private cmds() As String
 

Private Sub Class_Initialize()
    pwd = "C:\WINNT\System32"
End Sub

Function GetBanner() As String
    GetBanner = banner
End Function

Sub Reset()
    Class_Initialize
    Erase cmds
End Sub

Function GetCmdList() As String
    GetCmdList = Join(cmds, vbCrLf)
End Function
    
Function GetResponse(cmd As String) As String

    Dim args() As String
    Dim tmp As String
    Dim isdrive As Boolean
    
    If Len(cmd) = 0 Then Exit Function
    
    push cmds, cmd
    
    cmd = LCase(cmd)
    args = Split(cmd, " ")
    
    If Len(args(0)) = 2 And Right(args(0), 1) = ":" Then
        If LCase(Left(args(0), 1)) = "c" Then
            GetResponse = pwd & ">"
        Else
            GetResponse = "The device is not ready." & vbCrLf & vbCrLf & pwd & ">"
        End If
        Exit Function
    End If
    
    Select Case args(0)
        Case "del":
                    tmp = Replace(nodel, "[ARG0]", args(1))
                    tmp = Replace(tmp, "[PWD]", pwd)
                    GetResponse = tmp
 
        Case "echo": GetResponse = pwd & ">"
        Case "cd"
                    If args(1) = ".." Then updir Else todir args(1)
                    GetResponse = pwd & ">"
        Case "dir"
                    tmp = Replace(cList, "[PWD]", pwd)
                    tmp = Replace(tmp, "[DIRTABLE]", dTable)
                    GetResponse = tmp
                    
        
        Case Else:
                    tmp = Replace(ha, "[ARG0]", args(0))
                    tmp = Replace(tmp, "[PWD]", pwd)
                    GetResponse = tmp
                    
    End Select

    
End Function


Private Sub updir()
    
    Dim tmp() As String
    Dim i As Long
    
    tmp = Split(pwd, "\")
    
    pwd = Empty
    For i = 0 To UBound(tmp) - 1
        pwd = pwd & tmp(i) & "\"
    Next
    
    If Len(pwd) > 3 And Right(pwd, 1) = "\" Then pwd = Mid(pwd, 1, Len(pwd) - 1)
    
End Sub


Private Sub todir(D)
    If InStr(1, D, "c:\", vbTextCompare) > 0 Then
        pwd = D
    Else
        pwd = pwd & IIf(Right(pwd, 1) = "\", "", "\") & D
    End If
End Sub




Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
