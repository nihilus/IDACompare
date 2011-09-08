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



Global crc As New clsCRC
Global sort As New CAlphaSort


Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
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

