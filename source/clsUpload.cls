VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUpload"
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
Public DumpDir As dumpDirs
Public state As oState
Public filename As String
Public FileSize As Long
Public RemoteHostIp As String
Public UploadComplete As Boolean
Public TimeOut As Boolean
Public dlIndex As Integer
Public Handled As Boolean

Public GUID As String
Public ScratchPad As Long

 
Private buffer() As Byte
Private fhnd As Long
 
Private Const MAX_FILE_SIZE As Long = 3000000 '~3mb
Private displayed_maxMsg As Boolean

Sub AppendBuffer(x() As Byte)
    If AryIsEmpty(buffer) Then
        buffer = x
    Else
        Dim oldUB As Long
        oldUB = UBound(buffer) + 1
        ReDim Preserve buffer(oldUB + UBound(x))
        CopyMemory buffer(oldUB), x(0), UBound(x) + 1
    End If
End Sub

Function GetBuffer() As Byte()
    GetBuffer = buffer
End Function
    
Function GetBufStr() As String
    GetBufStr = StrConv(buffer, vbUnicode)
End Function

Function ClearBuffer()
    Erase buffer
End Function

Sub SetFileSize(sz)
    On Error Resume Next
    FileSize = CLng(sz)
End Sub

Property Get RealFileSize() As Long
    On Error Resume Next
    RealFileSize = FileLen(filename)
End Property

Sub WriteChunk(B() As Byte)

    On Error Resume Next
    
    Dim sz As Long
    
    sz = RealFileSize()
    
    If sz > MAX_FILE_SIZE Then
        Module1.CloseFile fhnd
        If Not displayed_maxMsg Then dbg llreal, "REACHED MAX FILE SIZE CLOSING FILE EARLY " & filename
        displayed_maxMsg = True
        Exit Sub
    End If
    
    Put fhnd, , B()
     
    'Debug.Print "Chunk Size: " & UBound(b) & " FSize: " & FileSize & "  LOF: " & LOF(fhnd)
  
    If FileSize = 0 Then Exit Sub
       
    If LOF(fhnd) >= (FileSize - 10) Then
        UploadComplete = True
    End If
    
        
End Sub

Sub StartFileDump()
    On Error Resume Next
    
    filename = safeGetTmpFile(DumpDir)
    
    fhnd = FreeFile
    
    Open filename For Binary As fhnd
    
End Sub

Sub CloseFile()
    On Error Resume Next
    Module1.CloseFile fhnd
End Sub


Private Sub class_terminate()
    On Error Resume Next
    Module1.CloseFile fhnd
End Sub
