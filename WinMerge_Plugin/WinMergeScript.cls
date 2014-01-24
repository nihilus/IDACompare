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
  Dim defaultFilter As Long
   
  defaultFilter = CLng(GetSetting("winmerge", "settings", "defaultFilter", 0))
  
  'MsgBox "Using filter: " & defaultFilter
  
  Select Case defaultFilter
      Case 0: text = Aggressive(text, False, True)
      Case 1: text = LessAggressive(text)
      Case 2: text = frmVisualDebug.DebugFilter(text, Me)
      Case Default: MsgBox "An unknown default filter is set for IDACompare Winmerge plugin? Index: " & defaultFilter
  End Select
  
  size = Len(text)
  bChanged = True
  
  If Err.Number <> 0 Then MsgBox "Error in wmIDACompare plugin: " & Err.Description, vbExclamation
  
  PrediffBufferW = True
  
End Function

'this is probably overly agressive but seems to have fairly good results on initial tests.
Public Function Aggressive(text As String, Optional stdJmps As Boolean = True, Optional stripArgs As Boolean = True) As String

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
    
    Aggressive = Join(tmp, vbCrLf)
    
End Function

Public Function LessAggressive(text As String) As String

    Dim tmp() As String
    Dim i As Long, x, cmd
    Const stdJmps As Boolean = True
    Const stripArgs As Boolean = True
    
    Dim baseCmd As String
    Dim a As Long
    
    tmp = Split(text, vbCrLf)
    
    For i = 0 To UBound(tmp)
        If Right(tmp(i), 1) = ":" Then
            'its a label strip after the _ if it has one
            a = InStr(tmp(i), "_")
            If a < 1 Then
                tmp(i) = Empty
            Else
                tmp(i) = Mid(tmp(i), 1, a) 'at least we leave a label was here marker better than blank space..
            End If
        Else
            tmp(i) = Replace(tmp(i), vbTab, " ")
            tmp(i) = Replace(tmp(i), "  ", " ")
            tmp(i) = Trim(tmp(i))
            
            If Len(tmp(i)) > 0 Then
                cmd = Split(tmp(i), " ")
                baseCmd = cmd(0)
                 
                'now lets build a command signature based on some markers that dont change..
                If InStr(tmp(i), "[") > 0 Then baseCmd = baseCmd & "[]"
                If InStr(tmp(i), "+") > 0 Then baseCmd = baseCmd & "+"
                If InStr(tmp(i), "-") > 0 Then baseCmd = baseCmd & "-"
                If InStr(tmp(i), "*") > 0 Then baseCmd = baseCmd & "*"
                If ExtractConstant(tmp(i), x) Then baseCmd = baseCmd & x
                    
                x = InStr(tmp(i), ";")
                If x > 0 Then baseCmd = baseCmd & Mid(tmp(i), x) 'append on any automatic comment IDA included
                tmp(i) = Trim(baseCmd)
                
            End If
            
        End If
    Next
    
    LessAggressive = Join(tmp, vbCrLf)
    
End Function

Private Function ExtractConstant(ByVal asm, ByRef ConstVal) As Boolean
    Dim tmp() As String
    Dim i As Long
    Dim a As Long
    
    a = InStr(asm, ";")
    If a > 0 Then asm = Mid(asm, 1, a)
    
    If Right(asm, 1) = "]" Then
        asm = Mid(asm, 1, Len(asm) - 1)
        asm = Replace(asm, "+", " ")
        asm = Replace(asm, "-", " ")
        asm = Replace(asm, "*", " ")
    End If
    
    asm = Replace(asm, vbTab, " ")
    
    While InStr(asm, "  ") > 0
        asm = Replace(asm, "  ", " ")
    Wend

    asm = Trim(asm)
    tmp = Split(asm, " ")
   
    tmp(i) = Trim(tmp(UBound(tmp)))
    If Right(tmp(i), 1) = "h" And Len(tmp(i)) > 3 Then
        ConstVal = tmp(i)
        ExtractConstant = True
    End If
    
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

