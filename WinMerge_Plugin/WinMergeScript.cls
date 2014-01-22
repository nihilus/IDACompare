VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinMergeScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+----------------------------------------------------------------------+
'| This is a plugin for WinMerge <www.winmerge.org>.                    |
'| It is designed for block level diffing of asm code                   |
'| Copyright (C) 2014 by David Zimmer                                   |
'+----------------------------------------------------------------------+
'| This program is free software; you can redistribute it and/or modify |
'| it under the terms of the GNU General Public License as published by |
'| the Free Software Foundation; either version 2 of the License, or    |
'| (at your option) any later version.                                  |
'|                                                                      |
'| This program is distributed in the hope that it will be useful,      |
'| but WITHOUT ANY WARRANTY; without even the implied warranty of       |
'| MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the        |
'| GNU General Public License for more details.                         |
'|                                                                      |
'| You should have received a copy of the GNU General Public License    |
'| along with this program; if not, write to the Free Software          |
'| Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.            |
'+----------------------------------------------------------------------+
Option Explicit

Public Property Get PluginEvent() As String
  PluginEvent = "BUFFER_PREDIFF"
End Property

Public Property Get PluginDescription() As String
  PluginDescription = "This plugin is for block level diffing of IDA disassembly"
End Property

Public Property Get PluginFileFilters() As String
  PluginFileFilters = ".idacompare"
End Property

Public Property Get PluginIsAutomatic() As Boolean
  PluginIsAutomatic = True
End Property

Public Function PrediffBufferW(ByRef text As String, ByRef size As Long, ByRef bChanged As Boolean) As Boolean
  'text is the full text of the file to diff..
  
  On Error Resume Next
  
  Dim sTemp As String
  
  text = Standardize(text, False, True)
  size = Len(text)
  bChanged = True
  
  If Err.Number <> 0 Then MsgBox "Error in wmIDACompare plugin: " & Err.Description, vbExclamation
  
  PrediffBufferW = True
  
End Function

'this is probably overly agressive but seems to have fairly good results on initial tests.
Private Function Standardize(text As String, Optional stdJmps As Boolean = True, Optional stripArgs As Boolean = True) As String

    Dim tmp() As String
    Dim i As Long, x, cmd
    
    tmp = Split(text, vbCrLf)
    
    For i = 0 To UBound(tmp)
        If Right(tmp(i), 1) = ":" Then
            'its a label strip it
            tmp(i) = Empty
        Else
            tmp(i) = Replace(tmp(i), vbTab, " ")
            tmp(i) = Replace(tmp(i), "  ", " ")
            tmp(i) = Trim(tmp(i))
            If Len(tmp(i)) > 0 Then
                cmd = Split(tmp(i), " ")
                
                If stdJmps Then
                    If VBA.Left(cmd(0), 1) = "j" And cmd(0) <> "jmp" Then
                        cmd(0) = "jxx" 'standardize all conditional jumps
                    End If
                    If Not stripArgs Then tmp(i) = Join(cmd, " ")
                End If
                
                If stripArgs Then tmp(i) = cmd(0)
                
                x = InStr(tmp(i), ";")
                If x > 0 Then tmp(i) = Mid(tmp(i), 1, x - 1)
                tmp(i) = Trim(tmp(i))
            End If
        End If
    Next
    
    Standardize = Join(tmp, vbCrLf)
    
End Function


'example
'  sTemp = ""
'  bChanged = False
'  bLineHasStarted = True
'  For i = 0 To size - 1 'For all chars...
'    sChar = Mid(text, i + 1, 1)
'    Select Case sChar
'      Case "N", "n", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0" 'N or number...
'        If (bLineHasStarted = True) Then 'If line has started...
'          sChar = ""
'          bChanged = True
'        End If
'      Case vbCr, vbLf 'EOL...
'        bLineHasStarted = True
'      Case Else
'        bLineHasStarted = False
'    End Select
'    sTemp = sTemp & sChar
'  Next i
'
'  If (bChanged = True) Then 'If text has changed...
'    text = sTemp
'    size = Len(sTemp)
'  End If

