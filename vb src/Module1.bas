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

Public Declare Sub HideEA Lib "ida_compare.plw" (ByVal addr As Long)
Public Declare Sub ShowEA Lib "ida_compare.plw" (ByVal addr As Long)
Public Declare Function NextAddr Lib "ida_compare.plw" (ByVal addr As Long) As Long
Public Declare Function PrevAddr Lib "ida_compare.plw" (ByVal addr As Long) As Long
Public Declare Function OriginalByte Lib "ida_compare.plw" (ByVal addr As Long) As Byte
Public Declare Function IDAFilePath Lib "ida_compare.plw" Alias "FilePath" (ByVal buf_maxpath As String) As Long
Public Declare Function RootFileName Lib "ida_compare.plw" (ByVal buf_maxpath As String) As Long
Public Declare Function ProcessState Lib "ida_compare.plw" () As Long
Public Declare Function FuncIndex Lib "ida_compare.plw" (ByVal addr As Long) As Long
Public Declare Function FuncArgSize Lib "ida_compare.plw" (ByVal Index As Long) As Long
Public Declare Function FuncColor Lib "ida_compare.plw" (ByVal Index As Long) As Colors
Public Declare Sub PatchByte Lib "ida_compare.plw" (ByVal addr As Long, ByVal valu As Byte)
Public Declare Sub PatchWord Lib "ida_compare.plw" (ByVal addr As Long, ByVal valu As Long)
Public Declare Sub DelFunc Lib "ida_compare.plw" (ByVal addr As Long)
Public Declare Sub AddComment Lib "ida_compare.plw" (ByVal cmt As String, ByVal clr As Byte)
Public Declare Sub AddProgramComment Lib "ida_compare.plw" (ByVal cmt As String)
Public Declare Sub AddCodeXRef Lib "ida_compare.plw" (ByVal start As Long, ByVal endd As Long)
Public Declare Sub DelCodeXRef Lib "ida_compare.plw" (ByVal start As Long, ByVal endd As Long)
Public Declare Sub AddDataXRef Lib "ida_compare.plw" (ByVal start As Long, ByVal endd As Long)
Public Declare Sub DelDataXRef Lib "ida_compare.plw" (ByVal start As Long, ByVal endd As Long)
Public Declare Sub MessageUI Lib "ida_compare.plw" (ByVal Msg As String)
Public Declare Sub MakeCode Lib "ida_compare.plw" (ByVal addr As Long)
Public Declare Sub Undefine Lib "ida_compare.plw" (ByVal addr As Long)
Public Declare Sub AnalyzeArea Lib "ida_compare.plw" (ByVal startat As Long, ByVal endat As Long)
Public Declare Sub aGetName Lib "ida_compare.plw" Alias "GetName" (ByVal addr As Long, ByVal buf As String, ByVal bufSize As Long)
Public Declare Sub SetComment Lib "ida_compare.plw" (ByVal addr As Long, ByVal comment As String)
Public Declare Sub GetComment Lib "ida_compare.plw" (ByVal addr As Long, ByVal comment As String)
Public Declare Function NumFuncs Lib "ida_compare.plw" () As Long
Public Declare Function FunctionStart Lib "ida_compare.plw" (ByVal functionIndex As Long) As Long
Public Declare Function FunctionEnd Lib "ida_compare.plw" (ByVal functionIndex As Long) As Long
Public Declare Sub Jump Lib "ida_compare.plw" (ByVal offset As Long)
Public Declare Sub RemvName Lib "ida_compare.plw" (ByVal offset As Long)
Public Declare Sub Setname Lib "ida_compare.plw" (ByVal offset As Long, ByVal name As String)
Public Declare Sub aRefresh Lib "ida_compare.plw" Alias "Refresh" ()
Public Declare Function ScreenEA Lib "ida_compare.plw" () As Long
Public Declare Sub SelBounds Lib "ida_compare.plw" (selstart As Long, selend As Long)
Public Declare Function GetBytes Lib "ida_compare.plw" (ByVal offset As Long, buf As Byte, ByVal length As Long) As Long
Private Declare Sub FuncName Lib "ida_compare.plw" (ByVal offset As Long, ByVal buf As String, ByVal bufSize As Long)
Private Declare Function GetAsm Lib "ida_compare.plw" (ByVal offset As Long, ByVal buf As String, ByVal length As Long) As Long

     
 
Enum Colors
        COLOR_DEFAULT = &H1           ' Default
        COLOR_REGCMT = &H2            ' Regular comment
        COLOR_RPTCMT = &H3            ' Repeatable comment (comment defined somewhere else)
        COLOR_AUTOCMT = &H4           ' Automatic comment
        COLOR_INSN = &H5              ' Instruction
        'COLOR_DATNAME = &H6           ' Dummy Data Name
        'COLOR_DNAME = &H7             ' Regular Data Name
        'COLOR_DEMNAME = &H8           ' Demangled Name
        'COLOR_SYMBOL = &H9            ' Punctuation
        'COLOR_CHAR = &HA              ' Char constant in instruction
        'COLOR_STRING = &HB            ' String constant in instruction
        'COLOR_NUMBER = &HC            ' Numeric constant in instruction
        'COLOR_VOIDOP = &HD            ' Void operand
        'COLOR_CREF = &HE              ' Code reference
        'COLOR_DREF = &HF              ' Data reference
        'COLOR_CREFTAIL = &H10         ' Code reference to tail byte
        'COLOR_DREFTAIL = &H11         ' Data reference to tail byte
        COLOR_ERROR = &H12            ' Error or problem
        COLOR_PREFIX = &H13           ' Line prefix
        COLOR_BINPREF = &H14          ' Binary line prefix bytes
        COLOR_EXTRA = &H15            ' Extra line
        COLOR_ALTOP = &H16            ' Alternative operand
        'COLOR_HIDNAME = &H17          ' Hidden name
        COLOR_LIBNAME = &H18          ' Library function name
        COLOR_LOCNAME = &H19          ' Local variable name
        COLOR_CODNAME = &H1A          ' Dummy code name
        COLOR_ASMDIR = &H1B           ' Assembler directive
        'COLOR_MACRO = &H1C            ' Macro
        COLOR_DSTR = &H1D             ' String constant in data directive
        COLOR_DCHAR = &H1E            ' Char constant in data directive
        COLOR_DNUM = &H1F             ' Numeric constant in data directive
        COLOR_KEYWORD = &H20          ' Keywords
        'COLOR_REG = &H21              ' Register name
        COLOR_IMPNAME = &H22          ' Imported name
        'COLOR_SEGNAME = &H23          ' Segment name
        'COLOR_UNKNAME = &H24          ' Dummy unknown name
        COLOR_CNAME = &H25            ' Regular code name
        'COLOR_UNAME = &H26            ' Regular unknown name
        'COLOR_COLLAPSED = &H27        ' Collapsed line
        'COLOR_FG_MAX = &H28           ' Max color number
End Enum

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Function isIde() As Boolean
    On Error GoTo hell
    Debug.Print 1 / 0
    isIde = False
    Exit Function
hell:
    isIde = True
End Function

Function GetFName(offset As Long) As String
    Dim buf As String
    Dim l As Long
    
    buf = String(257, Chr(0))
    
    FuncName offset, buf, Len(buf)
    
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

Function GetAsmCode(offset) As String
    Dim buf As String
    Dim sLen As Long
    
    buf = String(257, Chr(0))
    
    sLen = GetAsm(offset, buf, Len(buf))
    
    If sLen > 1 Then
        GetAsmCode = Mid(buf, 1, sLen)
    End If
    
End Function

'Function GetAsmRange(start As Long, leng As Long) As String
'    Dim x As String, tmp As String, i As Long, n As String
'
'    For i = 0 To leng - 1
'
'        tmp = GetAsmCode(start + i)
'        If Len(tmp) > 0 Then x = x & tmp & vbCrLf
'
'    Next
'
'    GetAsmRange = x
'
'End Function

Function GetAsmRange(start As Long, leng As Long) As String
    Dim x As String, tmp As String, i As Long, n As String
    
    For i = 0 To leng - 1
        tmp = GetAsmCode(start + i)
        If Len(tmp) > 0 Then
            
            If i <> 0 Then 'add in local labels...but not the function name (offset 0)
                n = GetName(start + i)
                If Len(n) > 0 Then x = x & n & ":" & vbCrLf
            End If
            
            x = x & tmp & vbCrLf
            
        End If
    Next
    
    GetAsmRange = x
    
End Function

Function HexDumpBytes(start As Long, leng As Long) As String
    Dim buf() As Byte, i As Long, x As String
    
    ReDim buf(1 To leng)
    GetBytes start, buf(1), leng
    
    For i = 1 To leng
        x = x & GetHex(buf(i)) & " "
    Next
    
    HexDumpBytes = x
    
End Function

Function GetName(offset) As String
    Dim buf As String, x
    buf = String(257, Chr(0))
    
    aGetName CLng(offset), buf, 257
    
    x = InStr(buf, Chr(0))
    If x = 1 Then
        buf = ""
    ElseIf x > 2 Then
        buf = Mid(buf, 1, x - 1)
    End If
    
    GetName = buf

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

Function FileExists(path) As Boolean
  On Error GoTo hell
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  Exit Function
hell:
     MsgBox "Malformed Path FileExists(" & path & ")", vbInformation
End Function


Function loadedFile() As String
    Dim buf As String
    Dim retlen As Long
    buf = String(261, Chr(0))
    
    retlen = IDAFilePath(buf)
    loadedFile = Mid(buf, 1, retlen)
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
