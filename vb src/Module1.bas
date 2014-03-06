Attribute VB_Name = "Module1"
Option Explicit
'Author:   David Zimmer <dzzie@yahoo.com> - Copyright 2004
'Site:     http://www.geocities.com/dzzie

'License:
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


'this is a little stupid, but it is what it is..if the dll extension isnt dll, it must be declared explicitly
'since we want to work with both the plw and the p64..we need declares for both, and choose the appropirate one at runtime.
'at least the prototypes are the same between the two thanks to making the C shim layer universal..

Private Declare Function IDAFilePath32 Lib "ida_compare.plw" Alias "FilePath" (ByVal buf_maxpath As String) As Long
Private Declare Sub MessageUI32 Lib "ida_compare.plw" Alias "MessageUI" (ByVal Msg As String)
Private Declare Sub GetName32 Lib "ida_compare.plw" Alias "GetName" (ByVal addr As Currency, ByVal buf As String, ByVal bufSize As Long)
Private Declare Function NumFuncs32 Lib "ida_compare.plw" Alias "NumFuncs" () As Long
Private Declare Function FunctionStart32 Lib "ida_compare.plw" Alias "FunctionStart" (ByVal functionIndex As Long) As Currency
Private Declare Function FunctionEnd32 Lib "ida_compare.plw" Alias "FunctionEnd" (ByVal functionIndex As Long) As Currency
Private Declare Sub FuncName32 Lib "ida_compare.plw" Alias "FuncName" (ByVal offset As Currency, ByVal buf As String, ByVal bufSize As Long)
Private Declare Sub Setname32 Lib "ida_compare.plw" Alias "SetName" (ByVal offset As Currency, ByVal name As String)
Private Declare Sub Refresh32 Lib "ida_compare.plw" Alias "Refresh" ()
Private Declare Function GetBytes32 Lib "ida_compare.plw" Alias "GetBytes" (ByVal offset As Currency, buf As Byte, ByVal length As Long) As Long
Private Declare Function GetAsm32 Lib "ida_compare.plw" Alias "GetAsm" (ByVal offset As Currency, ByVal buf As String, ByVal length As Long) As Long
Private Declare Function Addx32 Lib "ida_compare.plw" Alias "Addx64" (ByVal offset As Currency, ByVal val As Long) As Currency
Private Declare Function Subx32 Lib "ida_compare.plw" Alias "Subx64" (ByVal v0 As Currency, ByVal v1 As Currency) As Currency

Private Declare Function IDAFilePath64 Lib "ida_compare.p64" Alias "FilePath" (ByVal buf_maxpath As String) As Long
Private Declare Sub MessageUI64 Lib "ida_compare.p64" Alias "MessageUI" (ByVal Msg As String)
Private Declare Sub GetName64 Lib "ida_compare.p64" Alias "GetName" (ByVal addr As Currency, ByVal buf As String, ByVal bufSize As Long)
Private Declare Function NumFuncs64 Lib "ida_compare.p64" Alias "NumFuncs" () As Long
Private Declare Function FunctionStart64 Lib "ida_compare.p64" Alias "FunctionStart" (ByVal functionIndex As Long) As Currency
Private Declare Function FunctionEnd64 Lib "ida_compare.p64" Alias "FunctionEnd" (ByVal functionIndex As Long) As Currency
Private Declare Sub FuncName64 Lib "ida_compare.p64" Alias "FuncName" (ByVal offset As Currency, ByVal buf As String, ByVal bufSize As Long)
Private Declare Sub Setname64 Lib "ida_compare.p64" Alias "Setname" (ByVal offset As Currency, ByVal name As String)
Private Declare Sub Refresh64 Lib "ida_compare.p64" Alias "Refresh" ()
Private Declare Function GetBytes64 Lib "ida_compare.p64" Alias "GetBytes" (ByVal offset As Currency, buf As Byte, ByVal length As Long) As Long
Private Declare Function GetAsm64 Lib "ida_compare.p64" Alias "GetAsm" (ByVal offset As Currency, ByVal buf As String, ByVal length As Long) As Long
Private Declare Function Addx64 Lib "ida_compare.p64" (ByVal offset As Currency, ByVal val As Long) As Currency
Private Declare Function Subx64 Lib "ida_compare.p64" (ByVal v0 As Currency, ByVal v1 As Currency) As Currency




Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Type MungeCurr
    Value As Currency
End Type

Private Type Munge2Long
    LoValue As Long
    HiValue As Long
End Type

Global x64Mode As Boolean

Function x64ToHex(v As Currency) As String
    Dim c As MungeCurr
    Dim l As Munge2Long
    c.Value = v
    LSet l = c
    If l.HiValue = 0 Then
        x64ToHex = Hex(l.LoValue)
    Else
        x64ToHex = Hex(l.HiValue) & Right("00000000" & Hex(l.LoValue), 8)
    End If
End Function

'handles hex strings for 32bit and 64 bit numbers, leading 00's on high part not required, of course they are on lo if there is a high..
Function HextoX64(s As String) As Currency
    Dim c As MungeCurr
    Dim l As Munge2Long
    
    Dim lo As String, hi As String
    If Len(s) <= 8 Then
        l.LoValue = CLng("&h" & s)
    Else
        lo = Right(s, 8)
        hi = Left(s, Len(s) - 8)
        l.LoValue = CLng("&h" & lo)
        l.HiValue = CLng("&h" & hi)
    End If
    
    LSet c = l
    HextoX64 = c.Value
    
End Function

Function isIde() As Boolean
    On Error GoTo hell
    Debug.Print 1 / 0
    isIde = False
    Exit Function
hell:
    isIde = True
End Function

Function NumFuncs() As Long
    If x64Mode Then
        NumFuncs = NumFuncs64()
    Else
        NumFuncs = NumFuncs32()
    End If
End Function

Function Refresh()
    If x64Mode Then
        Refresh64
    Else
        Refresh32
    End If
End Function


Function FunctionStart(index As Long) As String
    Dim c As Currency
    If x64Mode Then
        c = FunctionStart64(index)
    Else
        c = FunctionStart32(index)
    End If
    FunctionStart = x64ToHex(c)
End Function

Function FunctionEnd(index As Long) As String
    Dim c As Currency
    If x64Mode Then
        c = FunctionEnd64(index)
    Else
        c = FunctionEnd32(index)
    End If
    FunctionEnd = x64ToHex(c)
End Function

Sub SetName(offset As String, newName As String)
    Dim addr As Currency
    
    addr = HextoX64(offset)
    
    If x64Mode Then
        Setname64 addr, newName
    Else
        Setname32 addr, newName
    End If

End Sub

Function GetFName(offset As String) As String
    Dim buf As String
    Dim l As Long
    Dim addr As Currency
    
    addr = HextoX64(offset)
    buf = String(257, Chr(0))
    
    If x64Mode Then
        FuncName64 addr, buf, Len(buf)
    Else
        FuncName32 addr, buf, Len(buf)
    End If
    
    l = InStr(buf, Chr(0))
    If l > 2 Then buf = Mid(buf, 1, l - 1)
    
    GetFName = buf
    
End Function

Function GetHex(x As Byte) As String
    Dim Y As String
    Y = Hex(x)
    If x < &H10 Then Y = "0" & Y
    GetHex = Y
End Function

Private Function GetAsmCode(addr As Currency) As String
    Dim buf As String
    Dim sLen As Long
    
    buf = String(500, Chr(0))
    
    If x64Mode Then
        sLen = GetAsm64(addr, buf, Len(buf))
    Else
        sLen = GetAsm32(addr, buf, Len(buf))
    End If
    
    If sLen > 1 Then
        GetAsmCode = Mid(buf, 1, sLen)
    End If
    
End Function

 Function GetAsmByOffset(offset As String) As String
    Dim buf As String
    Dim sLen As Long
    Dim addr As Currency

    addr = HextoX64(offset)
    buf = String(257, Chr(0))

    If x64Mode Then
        sLen = GetAsm64(addr, buf, Len(buf))
    Else
        sLen = GetAsm32(addr, buf, Len(buf))
    End If

    If sLen > 1 Then
        GetAsmByOffset = Mid(buf, 1, sLen)
    End If

End Function

Function GetAsmRange(start As String, leng As Long) As String
    Dim x As String, tmp As String, i As Long, n As String
    Dim addr As Currency

    addr = HextoX64(start)
    
    For i = 0 To leng - 1
        tmp = GetAsmCode(addr)
        If Len(tmp) > 0 Then
            If i <> 0 Then 'add in local labels...but not the function name (offset 0)
                n = GetName(addr)
                If Len(n) > 0 Then x = x & n & ":" & vbCrLf
            End If
            x = x & tmp & vbCrLf
        End If
        If x64Mode Then
            addr = Addx64(addr, 1)
        Else
            addr = Addx32(addr, 1)
        End If
    Next

    GetAsmRange = x
    
End Function

'takes in two 64 bit hex number strings, outputs their difference as a long..
Public Function SubX(v0 As String, v1 As String) As Long
    Dim a As Currency, b As Currency, c As Currency, tmp
    a = HextoX64(v0)
    b = HextoX64(v1)
    If x64Mode Then
        c = Subx64(a, b)
    Else
        c = Subx32(a, b)
    End If
    tmp = x64ToHex(c)
    If Len(tmp <= 8) Then
        SubX = CLng("&h" & tmp)
    Else
        SubX = -1 'error not used as intended..
    End If
End Function

Function HexDumpBytes(start As String, leng As Long) As String
    Dim buf() As Byte, i As Long, x As String
    Dim addr As Currency
    
    addr = HextoX64(start)
    
    ReDim buf(1 To leng)
    
    If x64Mode Then
        GetBytes64 addr, buf(1), leng
    Else
        GetBytes32 addr, buf(1), leng
    End If
    
    For i = 1 To leng
        x = x & GetHex(buf(i)) & " "
    Next
    
    HexDumpBytes = x
    
End Function

Private Function GetName(addr As Currency) As String
    Dim buf As String, x
    
    buf = String(300, Chr(0))
    
    If x64Mode Then
        GetName64 addr, buf, 257
    Else
        GetName32 addr, buf, 257
    End If
    
    x = InStr(buf, Chr(0))
    If x = 1 Then
        buf = ""
    ElseIf x > 2 Then
        buf = Mid(buf, 1, x - 1)
    End If
    
    GetName = buf

End Function

'Private Function GetName(offset As String) As String
'    Dim buf As String, x
'    Dim addr As Currency
'
'    addr = HextoX64(offset)
'    buf = String(257, Chr(0))
'
'    aGetName addr, buf, 257
'
'    x = InStr(buf, Chr(0))
'    If x = 1 Then
'        buf = ""
'    ElseIf x > 2 Then
'        buf = Mid(buf, 1, x - 1)
'    End If
'
'    GetName = buf
'
'End Function


Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
init:     ReDim ary(0): ary(0) = Value
End Sub

Function FileExists(path) As Boolean
  On Error GoTo hell
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  Exit Function
hell:
     MsgBox "Malformed Path FileExists(" & path & ")", vbInformation
End Function


Function LoadedFile() As String
    Dim buf As String
    Dim retlen As Long
    
    buf = String(261, Chr(0))
    
    If x64Mode Then
        retlen = IDAFilePath64(buf)
    Else
        retlen = IDAFilePath32(buf)
    End If
    
    LoadedFile = Mid(buf, 1, retlen)
    
End Function

Function GetHextxt(t As TextBox, v As Long) As Boolean
    
    On Error Resume Next
    v = CLng("&h" & t)
    If Err.Number > 0 Then
        MsgBox "Error " & t.Text & " is not valid hex number", vbInformation
        Exit Function
    End If
    
    GetHextxt = True
    
End Function


Function GetParentFolder(ByVal path) As String
    Dim tmp, ub
    If Len(path) < 1 Then Exit Function
    If Right(path, 1) = "\" Then path = Left(path, Len(path) - 1)
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
End Function


Sub Enable(t As TextBox, Optional enabled = True)
    t.BackColor = IIf(enabled, vbWhite, &H80000004)
    t.enabled = enabled
    t.Text = Empty
End Sub


Sub Insert(cn As Connection, tblName, fields, ParamArray params())
    Dim sql As String, i As Integer, values(), tn As String
    
    values() = params() 'force byval
    
    For i = 0 To UBound(values)
        tn = LCase(TypeName(values(i)))
        If tn = "string" Or tn = "textbox" Or tn = "date" Then
            'international users may on occasion get an sql insert error when
            'exporting the disasm text to the access db. It is probably caused
            'by this replace function below. Either its not properly escaping
            'all of the single quotes, or mabey the ado engine is translating
            'the unicode string to a an ascii one and another ' is being injected
            'there? I have not been able to reproduce yet but the bug should be
            'located here.
            values(i) = "'" & Replace(values(i), "'", "''") & "'"
        End If
    Next

    sql = "Insert into " & tblName & " (" & fields & ") VALUES(____)"
    sql = Replace(sql, "____", Join(values, ","))
    cn.Execute sql
    
End Sub

Function DllPath() As String
    Dim f As String
    Dim n As Long
    
    f = Space(260)
    n = GetModuleFileName(GetModuleHandle("IDACompare.dll"), f, 260)
    If n > 0 Then f = Mid(f, 1, n)
    DllPath = Replace(Trim(f), "IDACompare.dll", "", , , vbTextCompare)
    
End Function

Function OpenDB(cn As Connection, pth) As Boolean
    On Error Resume Next
    
    cn.Close
    Err.Clear
    
    If Len(pth) > 0 Then
        Const cnStr = "Provider=MSDASQL;Driver={Microsoft Access Driver (*.mdb)};DBQ=____;"
        cn.ConnectionString = Replace(cnStr, "____", pth)
    End If
    
    cn.Open
    OpenDB = IIf(Err.Number > 0, False, True)
    
End Function



Function FileNameFromPath(fullpath) As String
    Dim tmp() As String
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function
