Attribute VB_Name = "Module1"
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



Global crc As New clsCrc
Global sort As New CAlphaSort

Sub rtfHighlightAsm(asm As String, c As CFunction, tb As RichTextBox)
    
    On Error Resume Next
    
    Dim tmp() As String
    Dim x, i As Long
    Const indentLen = 2
    
    tmp() = Split(asm, vbCrLf)
    
    'first we add line breaks for comments and indents for code..
    For i = 0 To UBound(tmp)
        x = Trim(tmp(i))
        If right(x, 1) = ":" Then 'label
            tmp(i) = vbCrLf & x
        Else
            tmp(i) = Space(indentLen) & x
        End If
    Next
    
    tb.Text = Join(tmp, vbCrLf) 'save to textbox..
    
    rtf.SetWindowUpdate tb
    
    Dim curPos As Long
    Dim a As Long
    
   'now we highlight
    For i = 0 To UBound(tmp)
        x = Trim(tmp(i))
        
        If left(x, 1) = "j" Then 'isjump
            tb.SelStart = curPos
            tb.SelLength = Len(tmp(i))
            tb.SelColor = vbRed
        ElseIf left(x, 4) = "call" Then 'iscall
            tb.SelStart = curPos
            tb.SelLength = Len(tmp(i))
            tb.SelColor = vbBlue
        End If
        
        a = InStr(tmp(i), ";")
        If a > 0 Then 'comment
            tb.SelStart = curPos + a
            tb.SelLength = Len(tmp(i)) - a
            tb.SelColor = &H8000&
        End If
        
        If right(x, 1) = ":" Then 'is label
            tb.SelStart = curPos
            tb.SelLength = Len(tmp(i))
            tb.SelColor = &H404000
            tb.SelBold = True
        End If
        
        curPos = curPos + Len(tmp(i)) + 2 'for crlf
    Next
            
            
    'now we search for and highlight all constants from the function..
    Dim k
    Dim eol As Long
    Dim nextSpace As Long
    
    For Each k In c.Constants
        a = 0
        Do
            a = tb.Find(k, a)
            If a > 0 Then
                eol = InStr(a, tb.Text, vbCrLf)
                nextSpace = InStr(a + 1, tb.Text, " ")
                If nextSpace < eol And nextSpace > 0 Then eol = nextSpace
                tb.SelStart = a
                tb.SelLength = eol - a
                tb.SelBold = True
                a = a + tb.SelLength
            End If
        Loop While a > 0
    Next
          
'    For Each k In c.labels 'they are already red we dont need them bold to, to much processing
'        a = 0
'        Do
'            a = tb.Find(k, a)
'            If a > 0 Then
'                eol = InStr(a, tb.Text, vbCrLf)
'                nextSpace = InStr(a + 1, tb.Text, " ")
'                If nextSpace < eol And nextSpace > 0 Then eol = nextSpace
'                tb.SelStart = a
'                tb.SelLength = eol - a
'                tb.SelBold = True
'                a = a + tb.SelLength
'            End If
'        Loop While a > 0
'    Next
    
    tb.SelStart = 0
    
    rtf.SetWindowUpdate tb, False
    
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
 
Sub FormPos(fform As Form, Optional andSize As Boolean = False, Optional save_mode As Boolean = False)
    
    On Error Resume Next
    
    Dim f, sz, i, ff, def
    f = Split(",Left,Top,Height,Width", ",")
    
    If fform.WindowState = vbMinimized Then Exit Sub
    If andSize = False Then sz = 2 Else sz = 4
    
    For i = 1 To sz
        If save_mode Then
            ff = CallByName(fform, f(i), VbGet)
            SaveSetting App.EXEName, fform.Name & ".FormPos", f(i), ff
        Else
            def = CallByName(fform, f(i), VbGet)
            ff = GetSetting(App.EXEName, fform.Name & ".FormPos", f(i), def)
            CallByName fform, f(i), VbLet, ff
        End If
    Next
    
End Sub

Function RandomNum() As Long
    Dim tmp As Long
    On Error GoTo hell
hell:
    Randomize
    tmp = Round(Timer * Now * Rnd(), 0)
    RandomNum = tmp
End Function

Function isWithin(cnt As Integer, v1, v2, Optional min As Integer = 0) As Boolean
    
    Dim low As Long
    Dim high As Long
    
    If v1 <= min Or v2 <= min Then Exit Function
    
    If v1 = v2 Then
        isWithin = True
        Exit Function
    End If
    
    low = IIf(v1 < v2, v1, v2)
    
    high = v1
    If low = v1 Then high = v2
    
    If low + cnt >= high Then isWithin = True
    
    
End Function


Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    If Len(val) = 0 Then Exit Function
    If IsObject(c(val)) Then
        KeyExistsInCollection = True
    Else
        t = c(val)
        KeyExistsInCollection = True
    End If
 Exit Function
nope: KeyExistsInCollection = False
End Function

 



Function FileExists(path) As Boolean
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Function ado(sql) As Recordset
    Set ado = Form1.cn.Execute(sql)
End Function


Sub ResetPB(newMax, caption)
    With Form1
        If newMax < 1 Then newMax = 1
        .pb.Max = newMax
        .pb.value = 0
        .Label1.caption = caption
        .Label1.Refresh
    End With
End Sub

Function GetAllElements(lv As ListView, Optional selOnly As Boolean = False) As String
    Dim ret() As String, i As Integer, tmp As String
    Dim li As ListItem

    For i = 1 To lv.ColumnHeaders.Count
        tmp = tmp & lv.ColumnHeaders(i).Text & vbTab
    Next

    push ret, tmp
    push ret, String(50, "-")

    For Each li In lv.ListItems
        tmp = li.Text & vbTab
        For i = 1 To lv.ColumnHeaders.Count - 1
            If selOnly Then
                If li.Selected Then tmp = tmp & li.SubItems(i) & vbTab
            Else
                tmp = tmp & li.SubItems(i) & vbTab
            End If
        Next
        push ret, tmp
    Next

    GetAllElements = Join(ret, vbCrLf)

End Function

Function ReadFile(filename)
  Dim f, temp
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

'
'Sub ExactCrcMatch()
'
'    Dim li As ListItem
'    Dim lit As ListItem
'    Dim t, u
'    Dim x, i1, i2
'
'    'For Each x In c1
'    For Each lit In lv1.ListItems
'
'        If KeyExistsInCollection(c2, CStr(x)) Then 'exact crc match
'            Set c = c1(x)
'            i1 = c.li.index
'            'i1 = c1i(x)
'            i2 = c2(x)
'
'            Set li = lvExact.ListItems.Add
'            'li.Tag = t1(x) & "," & t2(x)
'
'            li.Text = lv1.ListItems(i1).SubItems(2)
'            li.SubItems(1) = lv2.ListItems(i2).SubItems(2)
'            li.SubItems(4) = "Exact CRC"
'
'            t = lv1.ListItems(i1).SubItems(1)
'            u = lv2.ListItems(i2).SubItems(1)
'
'            lv1.ListItems(i1).Tag = "delete"
'            lv2.ListItems(i2).Tag = "delete"
'            a.Remove lv1.ListItems(i1).SubItems(3)
'            b.Remove lv2.ListItems(i2).SubItems(3)
'
'            If t = u Then
'                li.SubItems(2) = "yes"
'            Else
'                li.SubItems(2) = t & "," & u
'            End If
'
'            li.SubItems(3) = lv1.ListItems(i1).SubItems(3)
'
'        End If
'
'        pb.value = pb.value + 1
'    Next
'
'     Dim i As Long
'
'    'remove matchs from top two list views
'    For i = lv1.ListItems.Count To 1 Step -1
'        Set li = lv1.ListItems(i)
'        If li.Tag = "delete" Then lv1.ListItems.Remove i
'    Next
'
'    For i = lv2.ListItems.Count To 1 Step -1
'        Set li = lv2.ListItems(i)
'        If li.Tag = "delete" Then lv2.ListItems.Remove i
'    Next
'
'End Sub

