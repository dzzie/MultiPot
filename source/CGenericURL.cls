VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGenericURL"
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


'Generic Url parser for http, ftp, and tftp urls & ftp echo string cmds
'include basic generic xor decrypter

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

Public err_msg As String
Public Busy As Boolean
Public SampleFile As String
Public URL As String

'this is a dirty hack
Private useDecodedBuffer As Boolean
Private DecodedBuffer As String



Function CheckSignature(ByVal MemBuffer As String) As Boolean
    
    Dim tmp As String, start As Long
    tmp = Replace(MemBuffer, Chr(0), "")
     
    DecodedBuffer = Empty
    useDecodedBuffer = False
    
    If InStr(1, tmp, "http://", vbTextCompare) > 0 Then CheckSignature = True
    If InStr(1, tmp, "ftp://", vbTextCompare) > 0 Then CheckSignature = True
    If InStr(1, tmp, "tftp", vbTextCompare) > 0 Then CheckSignature = True
    If AllOfTheseInstr(tmp, "echo open,echo user, ftp") Then CheckSignature = True
     
    If CheckSignature Then Exit Function
    
    'this is more processor intensive so we only do if we have to..
    'note if genericdecoder returns true then it modified tmp variable as well..
    useDecodedBuffer = GenericDecoder(tmp)
    If useDecodedBuffer Then DecodedBuffer = tmp
    
    If InStr(1, tmp, "http://", vbTextCompare) > 0 Then CheckSignature = True
    If InStr(1, tmp, "ftp://", vbTextCompare) > 0 Then CheckSignature = True
    If InStr(1, tmp, "tftp", vbTextCompare) > 0 Then CheckSignature = True
    If AllOfTheseInstr(tmp, "echo open,echo user, ftp") Then CheckSignature = True

End Function

Function BlockWhileBusy()
    
    While Busy
        DoEvents
        Sleep 30
    Wend
    
End Function


Function HandleShellcode(remoteHost As String, filepath As String, MemBuffer As String, Optional calledFrom As dumpDirs = eRPC445) As Boolean
     
    Dim mode As Long '0=http,1=ftp,2=tftp,3=echo ftp string
    Dim start As Long
    Dim tmp As String
    Dim buf() As Byte
    Dim leng As Long
    
    On Error GoTo done
    
    Busy = True
    SampleFile = Empty
    URL = Empty
    
       
    If useDecodedBuffer Then
        tmp = DecodedBuffer
    Else
        If InStr(MemBuffer, Chr(90) & Chr(0)) > 0 Then 'isunicode buffer with unicode nop slide
            tmp = Replace(MemBuffer, Chr(0), "")
        Else
            tmp = MemBuffer
        End If
    End If
    
    start = InStr(1, tmp, "http://", vbTextCompare)
    If start > 0 Then
        mode = 0
    Else
        start = InStr(1, tmp, "ftp://", vbTextCompare)
        If start > 0 Then
            mode = 1
        Else
            start = InStr(1, tmp, "tftp", vbTextCompare)
            If start > 0 Then
                mode = 2
            Else
                start = InStr(1, tmp, "echo open", vbTextCompare)
                If start > 0 Then
                    mode = 3
                Else
                    err_msg = "Shouldnt get here..Could not locate any url signatures?"
                    GoTo done
                End If
            End If
        End If
    End If
    
    leng = Len(MemBuffer) - start
    If leng > 255 Then leng = 255
    
    If leng < 1 Then
        err_msg = "Truncated payload?"
        GoTo done
    End If
    
    tmp = Mid(tmp, start, leng)
    buf = StrConv(tmp, vbFromUnicode)
    tmp = asciiScan(buf)
     
    If Len(tmp) = 0 Then
        err_msg = "Could not extract url?"
        GoTo done
    End If
    
    URL = tmp
    
    Dim lFile As String
    lFile = GetFreeFileInDumpDir(calledFrom, "./../")
        
    Select Case mode
        Case 0, 1: 'http, ftp 'ie ftp support is pretty crappy..we should probably do it ourselves with cftpget...
                HandleShellcode = DownloadFile(tmp, lFile)
                
        Case 2:
                Dim tftp As CTFTPClient
                Dim r_server As String, r_file As String
                
                Set tftp = New CTFTPClient
                If tftp.ParseTftpCmd(tmp, remoteHost, r_server, r_file) Then
                    tftp.RemoteFileName = r_file
                    HandleShellcode = tftp.GetFile(r_file, lFile, r_server)
                Else
                    dbg llinfo, "Could not parse tftp cmd: " & tmp
                End If
                
                Unload tftp
                Set tftp = Nothing
        Case 3:
                Dim ftp As CFtpGet
                Set ftp = New CFtpGet
                
                If ftp.LoadEchoString(tmp, remoteHost) Then
                    ftp.saveAs = lFile
                    If Len(ftp.user) = 1 Then mode = mport
                    If ftp.GetFile(mode) Then
                        dbg llreal, "[*] FTP Downloaded File: " & ftp.saveAs
                        HandleShellcode = True
                    Else
                        dbg llinfo, "Ftp download failed for ErrorLog: " & Join(ftp.ErrorLog, vbCrLf)
                        Exit Function
                    End If
                Else
                    dbg llinfo, "FTP.LoadEchoString Failed :" & tmp
                    Exit Function
                End If
                Unload ftp
                Set ftp = Nothing
                    
    End Select
   
   If HandleShellcode Then
        err_msg = "Download successful: " & tmp & " to " & lFile
        SampleFile = lFile
   Else
        err_msg = "Download failed: " & tmp
   End If
   
done:
    
   
   
   
End Function


 Private Function asciiScan(buf() As Byte) As String
    
    Dim ret As String
    Dim i As Long
    
    'scan till first not accepted character
    For i = 0 To UBound(buf)
        'Debug.Print buf(i) & " " & Chr(buf(i))
        
        Select Case buf(i)
            Case &H61 To &H7A:  'lower alpha
            Case &H41 To &H5A:  'UPPER ALPHA
            Case &H30 To &H39:  'numbers
            Case 39, 34, 63, 47, 92, 58, 37, 64, 38, 46, 35, 61, 45, 95, &H20, 62, &HD, &HA:   '"?/\:%@&.#= -_> CR LF
            Case Else: Exit For
        End Select
        
        ret = ret & Chr(buf(i))
        
    Next
                
    asciiScan = ret
    
    
 End Function
 




Public Function DownloadFile(URL, LocalFilename) As Boolean
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, CStr(URL), CStr(LocalFilename), 0, 0)
    If lngRetVal = 0 Then DownloadFile = True
End Function



 


Function HexDump(it)
    Dim my, i, c, s, A, B
    Dim lines() As String
    
    my = ""
    For i = 1 To Len(it)
        A = Asc(Mid(it, i, 1))
        c = Hex(A)
        c = IIf(Len(c) = 1, "0" & c, c)
        B = B & IIf(A > 65 And A < 120, Chr(A), ".")
        my = my & c & " "
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

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
 
 
 'scans the file and checks for a specific kind of xor decoder.
 'I dotn want a full disasm engine in here atm, so this is it for now.
 'it will look for loops immediatly preceeded by xor byte ptr [r32], i8
 'it will also accept and ignore one inc/dec r32 just before the loop,
 'or two inc/decs before the xor..assuming the loop takes us to the xor
 'loop jump has to be > 3 bytes, < 6 we are picky..but this hits all teh
 'encoders i have seen so far. we scan the file for opcode E2 loop and will
 'try up to 30 times per buffer before goving up. if we get a hit, then we
 'return true, and modify the actual buffer parameter as another return val.
 Function GenericDecoder(buffer As String) As Boolean
    
    Dim start As Long
    Dim xor_key As Long
    Dim B() As Byte
    Dim max As Long
    Dim loops As Long
    Dim i As Long
    
    start = InStr(buffer, Chr(&HE2))
    B() = StrConv(buffer, vbFromUnicode)
    
    loops = 0 'doubly just in case..
    
    While start > 0 And loops < 30
            
        xor_key = CheckLoopForXorKey(start - 1, B)
        
        If xor_key > 0 Then
            max = UBound(B) - 1
            If max - start > 1500 Then max = start + 1500 'just in case
            For i = start To max
                B(i) = B(i) Xor xor_key
            Next
            buffer = StrConv(B(), vbUnicode) 'modify actual byRef variable
            GenericDecoder = True
            Exit Function
        Else
            start = InStr(start + 1, buffer, Chr(&HE2))
        End If
        
        loops = loops + 1
        
    Wend
    
    
 End Function
 
 
 
 
 
 
 
 
 
 
 
 Function CheckLoopForXorKey(loopStart As Long, B() As Byte) As Long
    
    Dim Length As Long
    Dim xor_start As Long
    Dim xor_end As Long
    
    On Error GoTo guessNot
    
    If B(loopStart) <> &HE2 Then
        MsgBox "You hosed the loop start :P"
        Exit Function
    End If
    
    If B(loopStart + 1) < &HF8 Or B(loopStart + 1) > &HFB Then
        'if jump < 3 bytes or > 6 then not interested
        Exit Function
    End If
    
    Length = &HFE - B(loopStart + 1)
    
    xor_start = loopStart - Length
    If CheckIncDecReg(B(xor_start)) Then
        'loop takes us to inc r8, or dec r8, so skip to next byte
        xor_start = xor_start + 1
    End If
    
    If CheckIncDecReg(B(xor_start)) Then
        'next inst in loop is also a inc r8, or dec r8, so skip to next byte
        xor_start = xor_start + 1
    End If
    
    If B(xor_start) <> &H80 Then
        'we did not find an xor byte ptr [r32], i8
        Exit Function
    End If
    
    If Length > 3 And CheckIncDecReg(B(loopStart - 1)) Then
        'inst before loop is inc r8 or dec r8 so remove from length
        Length = Length - 1
    End If
    
    xor_end = xor_start + Length - 1
    
    CheckLoopForXorKey = B(xor_end) 'last instruction of xor is our xor_key
    
Exit Function
guessNot:
            CheckLoopForXorKey = 0
End Function

Function CheckIncDecReg(opcode As Byte) As Boolean
     If opcode >= &H40 And opcode <= &H4C Then
            CheckIncDecReg = True
     End If
End Function

 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
'this function is going to be fragile, only handles a couple specific
'cases of xor decoders..possibly break other readable urls with false positive?
'Function GenericDecoder(buffer As String) As Boolean
'
'    Dim B() As Byte
'    Dim A As Long, C As Long, i As Long
'
'    Dim xloop As String
'    Dim xor_1 As String
'    Dim xor_2 As String
'    Dim xor_key As Byte
'
'    xloop = Chr(&HE2) & Chr(&HF9) 'loop opcode
'    xor_1 = Chr(&H80) & Chr(&H73) '4 byte xor opcode with key as last
'    xor_2 = Chr(&H80) & Chr(&H74) '5 byte xor opcode with key as last
'
'    Dim F As Long
'    B = StrConv(buffer, vbFromUnicode)
'
'    A = InStr(buffer, xloop)
'    If A > 5 Then
'        C = InStr(buffer, xor_1)
'        If C > 0 And C < A And (C + 8) > A Then
'            'xor found and xor before loop and xor with 8 bytes of loop
'            xor_key = B(C + 3 - 1) '-1 for zero based array offset
'        Else
'            C = InStr(buffer, xor_2)
'            If C > 0 And C < A And (C + 8) > A Then
'                xor_key = B(C + 4 - 1) '-1 for zero based array offset
'            End If
'        End If
'    End If
'
'    If xor_key = 0 Then Exit Function
'
'    For i = A To UBound(B)
'        B(i) = B(i) Xor xor_key
'    Next
'
'    buffer = StrConv(B, vbUnicode)
'    GenericDecoder = True
'
'End Function


'generic decoder function is only based on observation of a series
'of decoders i found across multiple payloads.

'rbot dcom
'
'450 : 5b                            pop     ebx
'451 : 31 c9                         xor     ecx, ecx
'453 : b1 e0                         mov     cl, 0E0h ; 'à'
'455 : 80 73  c 12                   xor     byte ptr [ebx+0Ch], 12h
'459 : 43                            inc     ebx
'45A : e2 f9                         loop    loc_455'
'
'
'multistage -shellcode - dcom / dcom_multistage2 '
'
'418 : eb 15                         jmp     short loc_42F
'41A : b9 8b e6 13 41                mov     ecx, 4113E68Bh
'41F : 81 f1 39 e6 13 41             xor     ecx, 4113E639h
'425 : 5e                            pop     esi
'426 : 80 74 31 ff 42                xor     byte ptr [ecx+esi-1], 42h
'42B : e2 f9                         loop    loc_426
'42D : eb  5                         jmp     short loc_434
'42F : e8 e6 ff ff ff                call    loc_41A
'
'multistage 2
'418 : eb 15                         jmp     short loc_42F
'41A : b9 8b e6 13 41                mov     ecx, 4113E68Bh
'41F : 81 f1 39 e6 13 41             xor     ecx, 4113E639h
'425 : 5e                            pop     esi
'426 : 80 74 31 ff 85                xor     byte ptr [ecx+esi-1], 85h
'42B : e2 f9                         loop    loc_426
'42D : eb  5                         jmp     short loc_434
'
'
' the new generic decoder now handles these two too..
'-----------------------------------------------------------
'lsass xor loop
'14 : 33 c9                         xor     ecx, ecx
'16 : 66 b9 7d  1                   mov     cx, 17Dh
'1A : 80 34  a 99                   xor     byte ptr [edx+ecx], 99h
'1E : e2 fa                         loop    loc_1A
'
'dipnet.f xor loop
'1E : 80 37 ff                      xor     byte ptr [edi], 0FFh
'21 : 47                            inc     edi
'22:  e2 fa

'Opcode      Instruction         Description
'-----------------------------------------------------
'34 ib       XOR AL,imm8         AL XOR imm8
'35 iw       XOR AX,imm16        AX XOR imm16
'35 id       XOR EAX,imm32       EAX XOR imm32
'80 /6 ib    XOR r/m8,imm8       r/m8 XOR imm8
'81 /6 iw    XOR r/m16,imm16     r/m16 XOR imm16
'81 /6 id    XOR r/m32,imm32     r/m32 XOR imm32
'83 /6 ib    XOR r/m16,imm8      r/m16 XOR imm8 (sign-extended)
'83 /6 ib    XOR r/m32,imm8      r/m32 XOR imm8 (sign-extended)
'30 /r       XOR r/m8,r8         r/m8 XOR r8
'31 /r        XOR r/m16,r16      r/m16 XOR r16
'31 /r       XOR r/m32,r32       r/m32 XOR r32
'32 /r       XOR r8,r/m8      r8 XOR r/m8
'33 /r        XOR r16,r/m16      r16 XOR r/m16
'33 /r        XOR r32,r/m32      r32 XOR r/m32
''
'
'00401843     83F3 10        XOR EBX,10
'00401846     83F0 10        XOR EAX,10
'00401849     8030 10        XOR BYTE PTR DS:[EAX],10
'0040184C     8033 10        XOR BYTE PTR DS:[EBX],10
'0040184F     8070 01 10     XOR BYTE PTR DS:[EAX+1],10
'00401853     8071 01 10     XOR BYTE PTR DS:[ECX+1],10
'00401857     807419 01 10   XOR BYTE PTR DS:[ECX+EBX+1],10

