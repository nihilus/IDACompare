VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAlphaSort"
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

'this class is currently unused in this project

'quick -n-dirty..seems to miss z's
'
'
'    Dim tmp() As String
'
'    push tmp, "a"
'    push tmp, "bc"
'    push tmp, "ba"
'    push tmp, "a"
'    push tmp, "c"
'    push tmp, "a"
'    push tmp, "aaa"
'    push tmp, "c"
'    push tmp, "z"
'    push tmp, "g"
'
'    tmp = sort.AlphaSort(tmp)
'
'    Text1 = Join(tmp, vbCrLf)
'
'
'    Exit Sub

Function AlphaSort(list() As String) As String()

    Dim i As Long
    Dim tmp() As String
    Dim letter As Byte
    Dim index As Integer
    Dim ti As Long
    Dim x
    
    Dim letterlist()
    
    Dim cBase As New Collection
    Dim cOut As New Collection

    For i = 0 To UBound(list)
        list(i) = Trim(list(i))
        If Len(list(i)) > 0 Then cBase.Add LCase(list(i))
    Next
    
    letter = Asc("a")
    
    Do While cBase.Count > 0
            
        ti = cBase.Count
        
        For x = cBase.Count To 1 Step -1
            If Left(cBase(x), 1) = Chr(letter) Then
                push letterlist, cBase(x)
                cBase.Remove x
            End If
        Next
        
        If cBase.Count = ti Then 'no more of this letter
            
            If Not AryIsEmpty(letterlist) Then
                
                letterlist() = SubSort(letterlist, 2)
                
                For Each x In letterlist
                    cOut.Add x
                Next
                
                Erase letterlist
            
            End If
            
            
            letter = letter + 1
            If letter > Asc("z") Then Exit Do
        End If
        
    Loop
        
    For Each x In cOut
        push tmp, x
    Next
    
    If cBase.Count > 0 Then MsgBox cBase.Count & " " & cBase(1)

    AlphaSort = tmp


End Function

Private Function SubSort(llist(), index As Integer) As Variant()

    Dim cBase As New Collection
    'Dim cOut As New Collection
    Dim i, ti, x
    Dim letter As Byte
    Dim letterlist()
    
    For Each x In llist
        cBase.Add x
    Next
    
    letter = Asc("a")
    
    Do While cBase.Count > 0
            
        ti = cBase.Count
        
        For x = cBase.Count To 1 Step -1
            If Len(cBase(x)) < index Then
                push letterlist, cBase(x)
                cBase.Remove x
            End If
        Next
        
        For x = cBase.Count To 1 Step -1
            If Mid(cBase(x), index, 1) = Chr(letter) Then
                push letterlist, cBase(x)
                cBase.Remove x
            End If
        Next
        
        If cBase.Count = ti Then 'no more of this letter
            letter = letter + 1
            If letter > Asc("z") Then
                If index < 4 Then
                    index = index + 1
                    letterlist() = SubSort(letterlist, index)
                Else
                    Exit Do
                End If
            End If
            
        End If
        
    Loop
    
    SubSort = letterlist
    
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



Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

'Function AlphaSort(list() As String) As String()
'
'    Dim i As Long
'    Dim tmp() As String
'    Dim letter As Byte
'    Dim index As Integer
'    Dim ti As Long
'    Dim x
'
'    Set cBase = New Collection
'    Set cOut = New Collection
'
'    For i = 0 To UBound(list)
'        cBase.Add LCase(list(i))
'    Next
'
'    letter = Asc("a")
'
'    Do While cBase.Count > 0
'
'        ti = cBase.Count
'
'        For x = cBase.Count To 1 Step -1
'            If Left(cBase(x), 1) = Chr(letter) Then
'                cOut.Add cBase(x)
'                cBase.Remove x
'            End If
'        Next
'
'        If cBase.Count = ti Then 'no more of this letter
'            letter = letter + 1
'            If letter > Asc("z") Then Exit Do
'        End If
'
'    Loop
'
'    For Each x In cOut
'        push tmp, x
'    Next
'
'    If cBase.Count > 0 Then MsgBox cBase.Count & " " & cBase(1)
'
'    AlphaSort = tmp
'
'
'End Function
