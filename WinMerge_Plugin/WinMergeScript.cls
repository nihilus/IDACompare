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
      Case 3: text = ComplexSignature(text)
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
                 
                'now lets reduce the command to the base elements it contains
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

Public Function ComplexSignature(text As String) As String

    Dim tmp() As String
    Dim i As Long, x, cmd, c, r
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
            
            While InStr(tmp(i), "  ")
                tmp(i) = Replace(tmp(i), "  ", " ")
            Wend
                
            tmp(i) = Trim(tmp(i))
            
            If Len(tmp(i)) > 0 Then
                cmd = Split(tmp(i), " ")
                baseCmd = cmd(0)
                 
                'now lets reduce the command to the base elements it contains (try to achieve generic asm menumonic format for instruction)
                'eg. mov reg32, reg32  or  mov reg32, [] etc..
                For Each c In cmd
                     baseCmd = baseCmd & stdReg(c)
                     baseCmd = baseCmd & PointerSignature(c)
                     If ExtractConstant(c, x) Then baseCmd = baseCmd & x
                     If c = ";" Then Exit For 'we process comments next..
                Next

                x = InStr(tmp(i), ";")
                If x > 0 Then baseCmd = baseCmd & Mid(tmp(i), x) 'append on any automatic comment IDA included
                tmp(i) = Trim(baseCmd)
                
            End If
            
        End If
    Next

    ComplexSignature = Join(tmp, vbCrLf)
End Function

Private Function PointerSignature(ByVal cmd As String)

    If Len(cmd) = 0 Then Exit Function
    
    If InStr(cmd, "[") > 0 Then PointerSignature = PointerSignature & "["
    If InStr(cmd, "+") > 0 Then PointerSignature = PointerSignature & "+"
    If InStr(cmd, "-") > 0 Then PointerSignature = PointerSignature & "-"
    If InStr(cmd, "*") > 0 Then PointerSignature = PointerSignature & "*"
    If InStr(cmd, "]") > 0 Then PointerSignature = PointerSignature & "]"
    
    If Len(PointerSignature) > 0 Then PointerSignature = " " & PointerSignature
    
End Function

Private Function stdReg(ByVal cmd As String)
    
    Dim reg32() As String, reg16() As String, reg8() As String, stack() As String, r
    Dim hasComma As Boolean
    
    If Len(cmd) = 0 Then Exit Function
    
    cmd = LCase(cmd)
    
    If Right(cmd, 1) = "," Then
        hasComma = True
        cmd = Mid(cmd, 1, Len(cmd) - 1)
    End If
    
    stack = Split("esp,ebp", ",")
    reg32 = Split("eax,ebx,ecx,edx,esi,edi", ",")
    reg16 = Split("ax,bx,cx,dx,si,di", ",")
    reg8 = Split("al,ah,bl,bh,cl,ch,dl,dh", ",")
    
    For Each r In stack
        If cmd = r Then
            stdReg = " " & r & IIf(hasComma, ",", Empty)
            Exit Function
        End If
    Next
    
    For Each r In reg32
        If cmd = r Then
            stdReg = " reg32" & IIf(hasComma, ",", Empty)
            Exit Function
        End If
    Next
    
    For Each r In reg16
        If cmd = r Then
            stdReg = " reg16" & IIf(hasComma, ",", Empty)
            Exit Function
        End If
    Next
    
    For Each r In reg8
        If cmd = r Then
            stdReg = " reg8" & IIf(hasComma, ",", Empty)
            Exit Function
        End If
    Next
    
    
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
    If Right(tmp(i), 1) = "h" And Len(tmp(i)) > 3 And tmp(i) <> "push" Then
        ConstVal = " " & tmp(i)
        Debug.Print "const found: " & asm & " -> " & tmp(i)
        ExtractConstant = True
    End If
    
End Function

