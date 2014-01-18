VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Author: david@idefense.com <david@idefense.com, dzzie@yahoo.com>
'
'License: Copyright (C) 2005 iDefense.com, A Verisign Company
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

Public Length As Long
Public Calls As Long
Public matched As Boolean
Public Name As String
Public autoid As Long
Public index As Long
Public Jumps As Long
Public Pushs As Long
Public mCRC As String
Public li As ListItem
Public esp As Long
Public MatchMethod As String
Public fxCalls As New Collection
Public Constants As New Collection
Public labels As New Collection
Public Standardized As String

Function StandardizeAsm(ByVal asm As String) As String
    
    Dim tmp() As String
    Dim cmd() As String
    Dim x As Long
    Dim j As Long
    Dim i As Long
    Dim k
    Dim t As String
    
    tmp = Split(asm, vbCrLf)
    
    'quick profile of function elements
    'hit counts for: calls, pushs, jxxs, esp size, and apis called
    For j = 0 To UBound(tmp)
            
            'If InStr(tmp(j), "28955B88h") > 0 Then Stop
            
            Select Case left(tmp(j), 3)
                Case "cmp", "and", "xor", "pus", "mov", "lea", "sub": CheckConstant tmp(j)
            End Select
            
            If InStr(tmp(j), "call") > 0 Then
                Calls = Calls + 1
                If InStr(tmp(j), "sub_") < 1 Then 'not unnamed function call
                    If InStr(tmp(j), "ds:") > 0 Then 'is api call
                        t = Replace(tmp(j), "ds:", "")
                        t = Replace(t, "call ", "")
                        If InStr(t, ";") > 0 Then t = Mid(t, 1, InStr(t, ";") - 1)
                        If Len(Trim(t)) > 3 Then 'ignore call reg32
                            fxCalls.Add Trim(t)
                            'Debug.Print Trim(t)
                        End If
                    End If
                End If
            ElseIf InStr(tmp(j), "push") > 0 Then
                Pushs = Pushs + 1
            ElseIf left(tmp(j), 1) = "j" Then
                Jumps = Jumps + 1
            ElseIf CStr(tmp(j)) Like "sub*esp,*" Then
                k = Trim(Mid(tmp(j), InStr(tmp(j), "esp,") + 4))
                If InStr(k, "h") > 0 Then
                    esp = CLng("&h" & Replace(k, "h", ""))
                Else
                    esp = CLng(k)
                End If
            ElseIf right(Trim(tmp(j)), 1) = ":" Then
                labels.Add Mid(Trim(tmp(j)), 1, Len(Trim(tmp(j))) - 1)
            End If
    Next
        
    Standardize tmp
    'tmp = sort.AlphaSort(tmp) 'this would wack natural ordering
    
    Standardized = Join(tmp, vbCrLf)
    Standardized = Replace(Standardized, vbCrLf & vbCrLf, vbCrLf) 'strip empty lines (where labels were)
    mCRC = crc(Standardized)
    StandardizeAsm = mCRC
     
End Function

Function ReHash(ByVal asm)
    Dim tmp() As String
    If IsNull(asm) Then Exit Function
    tmp = Split(asm, vbCrLf)
    Standardize tmp, , False
    ReHash = crc(Join(tmp, vbCrLf))
    mCRC = ReHash
End Function

Private Sub Standardize(tmp() As String, Optional stdJmps As Boolean = True, Optional stripArgs As Boolean = True)
    Dim i As Long, x, cmd
    
    For i = 0 To UBound(tmp)
        If right(tmp(i), 1) = ":" Then
            'its a label strip it
            tmp(i) = Empty
        Else
            tmp(i) = Replace(tmp(i), vbTab, " ")
            tmp(i) = Replace(tmp(i), "  ", " ")
            If Len(tmp(i)) > 0 Then
                cmd = Split(tmp(i), " ")
                
                If stdJmps Then
                    If VBA.left(cmd(0), 1) = "j" And cmd(0) <> "jmp" Then
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
    
End Sub

Private Sub CheckConstant(ByVal asm)
    Dim tmp() As String
    Dim i As Long
    Dim a As Long
    
    a = InStr(asm, ";")
    If a > 0 Then asm = Mid(asm, 1, a)
    
    If right(asm, 1) = "]" Then
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
    
    'If InStr(asm, ",") Then i = 2 Else i = 1
    'If UBound(tmp) < i Then Exit Sub
   
    tmp(i) = Trim(tmp(UBound(tmp)))
    If right(tmp(i), 1) = "h" And Len(tmp(i)) > 3 Then
        If KeyExistsInCollection(Constants, "key:" & tmp(i)) Then
            Constants.Add tmp(i)
        Else
           Constants.Add tmp(i), "key:" & tmp(i)
        End If
    End If
    
End Sub

Function ConstantExists(key) As Boolean
    ConstantExists = KeyExistsInCollection(Constants, "key:" & key)
End Function
