VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAntiHammer"
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

Public MinutesBlocked As Long
Public Enabled As Boolean

Private ips As New Collection
Private checks As Integer

Private Sub Class_Initialize()
    MinutesBlocked = 60
    Enabled = True
End Sub

Public Sub Reset()
    Set ips = New Collection
End Sub

Public Function shouldBlock(ip As String, blockUntil As String) As Boolean
    
    
    If Not Enabled Then Exit Function
    
    Dim Exists As Boolean, h As CHost
    
    Exists = objInCol(ips, ip)
    
    If Exists Then
        Set h = ips(ip)
        
        If DateDiff("n", h.time, Format(Now, "h:n:s")) >= MinutesBlocked Then
            ips.remove h.ip
            Exit Function
        End If
        
        shouldBlock = True
        blockUntil = DateAdd("n", MinutesBlocked, h.time)
    End If
    
End Function

Public Sub AddIp(ip As String)

    Dim h As CHost
    
    If Len(ip) = 0 Then Exit Sub
    
    checks = checks + 1
    If checks > 50 Then
        CheckExpirations
        checks = 0
    End If
    
    If Not objInCol(ips, ip) Then
        Set h = New CHost
        h.ip = ip
        h.time = Format(Now, "h:n:s")
        ips.Add h, ip
    End If
        
End Sub

Sub RemoveIp(ip As String)
    
    If objInCol(ips, ip) Then
        ips.remove ip
    End If

End Sub

Private Sub CheckExpirations()
    
    Dim h As CHost

    For Each h In ips
        If DateDiff("n", h.time, Format(Now, "h:n:s")) >= MinutesBlocked Then
            ips.remove h.ip
        End If
    Next
            
End Sub



Private Function objInCol(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t As Object
    Set t = c(val)
    objInCol = True
 Exit Function
nope:
End Function

