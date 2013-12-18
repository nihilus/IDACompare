VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Form1 
   Caption         =   "IDACompare"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3795
      Left            =   0
      TabIndex        =   5
      Top             =   3960
      Width           =   8955
      Begin VB.CommandButton cmdCurrentProfile 
         Caption         =   "P"
         Height          =   255
         Left            =   4980
         TabIndex        =   20
         Top             =   1020
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ú"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4200
         TabIndex        =   19
         Top             =   60
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "Ù"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   4560
         TabIndex        =   18
         Top             =   60
         Width           =   375
      End
      Begin VB.TextBox txtReport 
         Height          =   735
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   240
         Width           =   8775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Load && Compare"
         Height          =   375
         Left            =   7380
         TabIndex        =   11
         Top             =   3300
         Width           =   1515
      End
      Begin VB.CommandButton cmdManualMatch 
         Caption         =   "Manual Match"
         Height          =   255
         Left            =   6420
         TabIndex        =   9
         Top             =   1020
         Width           =   1395
      End
      Begin VB.CommandButton cmdBreakMatch 
         Caption         =   "Break Match"
         Enabled         =   0   'False
         Height          =   255
         Left            =   7860
         TabIndex        =   8
         Top             =   1020
         Width           =   1035
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   255
         Left            =   5340
         TabIndex        =   7
         Top             =   1020
         Width           =   975
      End
      Begin VB.CheckBox chkExternalMatchs 
         Caption         =   "Check1"
         Height          =   255
         Left            =   3300
         TabIndex        =   6
         Top             =   1020
         Width           =   255
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   315
         Left            =   2400
         TabIndex        =   10
         Top             =   3360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ListView lvExact 
         Height          =   1935
         Left            =   60
         TabIndex        =   13
         Top             =   1320
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name 1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name 2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Len Match"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Stats"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Match Method"
            Object.Width           =   6174
         EndProperty
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   60
         TabIndex        =   17
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Matched Functions"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblTransform 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rename Tools"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1740
         TabIndex        =   15
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Label lblExternal 
         BackColor       =   &H8000000A&
         Caption         =   "External Matchs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3600
         TabIndex        =   14
         Top             =   1020
         Width           =   1395
      End
   End
   Begin MSScriptControlCtl.ScriptControl sc 
      Index           =   0
      Left            =   8400
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.TextBox txtB 
      Height          =   1515
      Left            =   4620
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2400
      Width           =   4335
   End
   Begin VB.TextBox txtA 
      Height          =   1515
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2400
      Width           =   4515
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "i"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "sz"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "crc"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lv2 
      Height          =   2055
      Left            =   4620
      TabIndex        =   1
      Top             =   300
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "i"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "sz"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "crc"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Unmatched Sample 1                                                                  Unmatched sample 2"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   6975
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuCheckAll 
         Caption         =   "Select All"
         Index           =   0
      End
      Begin VB.Menu mnuCheckAll 
         Caption         =   "Select None"
         Index           =   1
      End
      Begin VB.Menu mnuCheckAll 
         Caption         =   "Invert Selection"
         Index           =   2
      End
      Begin VB.Menu mnuCheckAll 
         Caption         =   "Remove Selected"
         Index           =   3
      End
      Begin VB.Menu mnuCheckAll 
         Caption         =   "Remove UnSelected"
         Index           =   4
      End
      Begin VB.Menu mnuCheckAll 
         Caption         =   "Select all w/Default names"
         Index           =   5
      End
   End
   Begin VB.Menu mnuPopupRename 
      Caption         =   "mnuPopupRename"
      Visible         =   0   'False
      Begin VB.Menu mnuRename 
         Caption         =   "Sequentially Rename Matchs"
         Index           =   0
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Port names from 1 to 2"
         Index           =   1
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Port names from 2 to 1"
         Index           =   2
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Help"
         Index           =   3
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
'
'
'
'this code was created quite quickly to test out the idea and data matching engine.
'there was allot of UI code to generate to wire in all teh desired features, so I
'split the horsepower required to generate this app between the two.
'
'the main data parsing engine could be more robust and more finely tuned, however
'for an initial release with all the desired features 3/4 of the way there for it
'was as far as was warrented for now.
'
'this should all be functional now and usable. I have done moderate testing. Future
'developements on it will depend on how heavily I end up using it.
'
'this is implemented as a standalone exe for debugging sake, developing complex
'features and functionality within a plugin can be a very painful experience, it didnt
'really hurt this app much seeing how we need data across several plugin instances
'anyway
'
'I should also say that this code has been multitasked some, supporting both signature
'scannign mode as well as compare version/variant mode. These features were hacked into
'this existing interface/codebase because it is so similar to avoid tons of repetive code.
'The downside of this is added complexity. Code is now littered with obscure special case
'clauses and there are probably bugs just because of this (same interface supporting 2
'bits of functionality)
'
'Anyway...its free andopen source and provides a good framework to see what works and
'what doesnt. UI should present enough info you can fine tune the code as you want and
'determine its strengths/weaknesses without too much mroe work.

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public cmndlg1 As New clsCmnDlg
Public cn As New Connection

'parallel function match collections, m1(index) matched with m2(index)
Dim m1 As New Collection 'of matched cfunction from ibd1
Dim m2 As New Collection 'of matched cfunction from ibd2
    
Dim a As New Collection 'of cfunction, all funcs for idb 1
Dim b As New Collection 'of cfunction, all funcs for idb 2

Dim c As CFunction
Dim h As CFunction
    
Public currentMDB As String
Public SigMode As Boolean
Dim sel_1 As ListItem
Dim sel_2 As ListItem
Dim sel_exact As ListItem

Enum CompareModes
    compare1 = 0
    compare2 = 1
    SignatureScan = 2
    TmpMode = 3
End Enum



 
Private Sub cmdBreakMatch_Click()
   
   Dim x, li As ListItem
   On Error Resume Next
   
   If sel_exact Is Nothing Then Exit Sub
   
   x = Split(sel_exact.Tag, ",")
   Set c = GetClassFromAutoID(a, x(0))
   Set h = GetClassFromAutoID(b, x(1))
   
   Set li = lv1.ListItems.Add(, "id:" & c.autoid)
   li.Tag = c.autoid
   li.Text = c.index
   li.SubItems(1) = c.Length
   li.SubItems(2) = c.Name
   li.SubItems(3) = c.mCRC
   
   Set li = lv2.ListItems.Add(, "id:" & h.autoid)
   li.Tag = h.autoid
   li.Text = h.index
   li.SubItems(1) = h.Length
   li.SubItems(2) = h.Name
   li.SubItems(3) = h.mCRC
   
   lvExact.ListItems.Remove sel_exact.index
   Set sel_exact = Nothing
   cmdBreakMatch.Enabled = False
            
End Sub

Private Sub cmdCurrentProfile_Click()
    On Error Resume Next
    Dim f As frmProfile
    
    Set c = Nothing
    Set h = Nothing
    
    If Not sel_exact Is Nothing Then
        lvExact_DblClick
        Exit Sub
    End If
    
    If sel_1 Is Nothing And sel_2 Is Nothing Then Exit Sub
    
    If Not sel_1 Is Nothing And Not sel_2 Is Nothing Then
        Set c = a(sel_1.ListSubItems(3))
        Set h = b(sel_2.ListSubItems(3))
    ElseIf Not sel_1 Is Nothing Then
        Set c = a(sel_1.ListSubItems(3))
    ElseIf Not sel_2 Is Nothing Then
        Set c = b(sel_2.ListSubItems(3))
    End If
    
    Set f = New frmProfile
    f.ShowProfile c, h
    
End Sub

Private Sub cmdFind_Click()
    If cn.State = 0 Then
        MsgBox "You must open a database first.", vbInformation
        Exit Sub
    End If
    frmFind.Show
End Sub

Private Sub Command1_Click(index As Integer)
        
    ScrollPage txtA, txtB, CBool(index)
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    If Me.Height < 8190 Then Me.Height = 8190
    If Me.Width < 9120 Then Me.Width = 9120
    
    Frame1.Top = Me.Height - Frame1.Height - 400
    txtA.Height = Frame1.Top - txtA.Top
    txtB.Height = txtA.Height
    
    txtA.Width = ((Me.Width / 2) - 40) - txtA.left
    txtB.left = txtA.left + txtA.Width + 20
    txtB.Width = Me.Width - txtB.left - 120
    
    lv1.Width = txtA.Width
    lv2.Width = txtB.Width
    lv2.left = txtB.left
    
    Frame1.Width = Me.Width - 120
    txtReport.Width = Frame1.Width - 120
    lvExact.Width = txtReport.Width
    
    Command1(1).left = txtB.left - Command1(1).Width
    Command1(0).left = txtB.left
    
End Sub

Private Sub lblExternal_Click()

    MsgBox "This will use the match functions in compare.vbs to " & vbCrLf & _
           "determine matchs rather than the default compiled in " & vbCrLf & _
           "logic." & vbCrLf & vbCrLf & _
           "This gives you a way to fine tune the match logic" & vbCrLf & _
           "to your needs without having to have VB or recompile", vbInformation
           
End Sub

Private Sub lv1_DblClick()
    On Error Resume Next
    Dim f As frmProfile
    If sel_1 Is Nothing Then Exit Sub
    Set f = New frmProfile
    Set c = a(sel_1.ListSubItems(3))
    f.ShowProfile c
End Sub

Private Sub lv2_DblClick()
    On Error Resume Next
    Dim f As frmProfile
    If sel_2 Is Nothing Then Exit Sub
    Set f = New frmProfile
    Set c = b(sel_2.ListSubItems(3))
    f.ShowProfile c
End Sub

Private Sub lvExact_DblClick()
   Dim x
   On Error Resume Next
   Dim f As frmProfile
   If sel_exact Is Nothing Then Exit Sub
   Set f = New frmProfile
   x = Split(sel_exact.Tag, ",")
   Set c = GetClassFromAutoID(a, x(0))
   Set h = GetClassFromAutoID(b, x(1))
   f.ShowProfile c, h
End Sub

Private Function GetClassFromAutoID(x As Collection, autoid) As CFunction
    Dim y As CFunction
    For Each y In x
        If y.autoid = autoid Then
            Set GetClassFromAutoID = y
            Exit Function
        End If
    Next
End Function

Private Sub cmdManualMatch_Click()
    Dim li As ListItem
    Dim t, u
    
    If sel_1 Is Nothing Then
        MsgBox "Select a function from list A to match"
        Exit Sub
    End If
    
    If sel_2 Is Nothing Then
        MsgBox "Select a function from list B to match"
        Exit Sub
    End If
    
    Set c = a(sel_1.ListSubItems(3))
    Set h = b(sel_2.ListSubItems(3))
       
    c.matched = True
    h.matched = True
    c.MatchMethod = "Manual Match"
    h.MatchMethod = "Manual Match"
    
    m1.Add c
    m2.Add h
    
    lv1.ListItems.Remove sel_1.index
    lv2.ListItems.Remove sel_2.index

    Set li = lvExact.ListItems.Add
    li.Tag = c.li.Tag & "," & h.li.Tag
    
    li.Text = c.Name
    li.SubItems(1) = h.Name
    li.SubItems(4) = c.MatchMethod
    
    t = c.Length
    u = h.Length
    
    If t = u Then
        li.SubItems(2) = "yes"
    Else
        li.SubItems(2) = t & "," & u
    End If
    
    li.SubItems(3) = c.Calls & "/" & h.Calls & " " & Hex(c.esp) & "/" & Hex(h.esp)
    
    Set sel_1 = Nothing
    Set sel_2 = Nothing
    cmdManualMatch.Enabled = False

End Sub

'if we have a function calling out to a matched function
'and it was the only way there..then if we only had one
'other unmatched on calling to it them we could relate them.

Private Sub Form_Load()
    
    If Len(Command) > 0 Then
        currentMDB = Replace(Command, """", Empty)
        If InStr(1, currentMDB, "/sigscan", vbTextCompare) > 1 Then
            SigMode = True
            mnuRename(0).Enabled = False
            mnuRename(1).Enabled = False
            Command2.Visible = False
            cmdManualMatch.Visible = False
            cmdBreakMatch.value = False
            currentMDB = Trim(Replace(currentMDB, "/sigscan", Empty))
        End If
        If Not FileExists(currentMDB) Then
            MsgBox "Usage: ida_compare.exe <mdb path to analyze>" & vbCrLf & vbCrLf & currentMDB
            currentMDB = Empty
        Else
            Me.Visible = True
            Command2_Click
        End If
    End If
    
    On Error Resume Next
    With lv1.ColumnHeaders(4)
        .Width = lv1.Width - .left - 100
    End With
    With lv2.ColumnHeaders(4)
        .Width = lv2.Width - .left - 100
    End With
    With lvExact.ColumnHeaders(5)
        .Width = lvExact.Width - .left - 100
    End With
    
End Sub


Sub LoadList(lv As ListView, mode As CompareModes, Optional minLen As Long = 30, Optional clause = "")
    Dim r() As String
    Dim rs As Recordset
    Dim li As ListItem
    Dim asm As String
   
    Dim t, u
    Dim tbl
    Dim isTableA As Boolean
    
    On Error Resume Next
    
    Select Case mode
        Case compare1:      tbl = "a":   isTableA = True
        Case TmpMode:       tbl = "tmp": isTableA = True
        Case compare2:      tbl = "b"
        Case SignatureScan: tbl = "signatures"
    End Select
    
    Set rs = ado("Select autoid,index,leng,fname,disasm from " & tbl & " where leng > " & minLen & clause)
    
    If rs Is Nothing Then
        MsgBox "Sql Query Failed could not load data from table: " & tbl & " min func len must be > " & minLen, vbCritical
        Exit Sub
    End If
    
    pb.Max = ado("Select count(autoid) as cnt from " & tbl & " where leng > " & minLen & clause)!cnt
    pb.value = 0
    Label1.caption = "Loading Table " & tbl
    Label1.Refresh
    
    While Not rs.EOF
        Set c = New CFunction
        asm = rs!disasm
        c.StandardizeAsm asm

        If KeyExistsInCollection(IIf(isTableA, a, b), c.mCRC) Then
            c.ReHash asm
            If KeyExistsInCollection(IIf(isTableA, a, b), c.mCRC) Then
                While KeyExistsInCollection(IIf(isTableA, a, b), c.mCRC)
                    c.mCRC = "rand:" & RandomNum
                Wend
            End If
        End If

        If Len(c.mCRC) > 0 Then
            Set li = lv.ListItems.Add(, "id:" & rs!autoid)
            Set c.li = li

            c.Length = rs!leng
            c.autoid = rs!autoid
            c.Name = rs!fname
            c.index = rs!index

            li.Tag = c.autoid
            li.Text = rs!index
            li.SubItems(1) = c.Length
            li.SubItems(2) = c.Name
            li.SubItems(3) = c.mCRC

            Err.Clear

            If mode = compare1 Or mode = TmpMode Then
                a.Add c, li.SubItems(3)               'collection "a" = function with crc as key
            Else
                b.Add c, li.SubItems(3)
            End If

            If Err.Number > 0 Then
                Debug.Print "Length:" & li.SubItems(1) & " CRC:" & li.SubItems(3) & " Name: " & c.Name & " Err:" & Err.Description
                Err.Clear
            End If

        End If

        rs.MoveNext
        pb.value = pb.value + 1
    Wend
    
    
End Sub

Sub ExactCrcMatch()
    
    Dim lit As ListItem
      
    Label1 = "CRC Matching"
    
    For Each lit In lv1.ListItems
        If KeyExistsInCollection(b, lit.SubItems(3)) Then
            Set c = a(lit.SubItems(3))
            Set h = b(lit.SubItems(3))
            AddToMatchCollection c, h, "Exact CRC"
        End If
        pb.value = pb.value + 1
    Next
    
End Sub

Sub CallPushMatch()

    pb.value = 0
    Label1 = "Call/Push Matching"
    
    For Each c In a
        For Each h In b
            If Not c.matched And Not h.matched Then
                If c.Calls = h.Calls And c.Pushs = h.Pushs Then  'same num of calls and pushs
                    If isWithin(60, c.Length, h.Length, 80) Then     'and length is close
                        If isWithin(4, c.Jumps, h.Jumps) Then    'and num jmps is close
                           AddToMatchCollection c, h, "Call/Push Match"
                        End If
                    End If
                 End If
            End If
        Next
        pb.value = pb.value + 1
        DoEvents
    Next
    
    pb.value = 0
    
End Sub

Sub EspMatch()

      pb.value = 0
      Label1 = "ESP Matching"
      
      For Each c In a
            For Each h In b
                If Not c.matched And Not h.matched Then
                     If isWithin(80, c.Length, h.Length, 80) Then
                        If c.esp <> 0 And c.esp = h.esp And isWithin(40, c.Length, h.Length) Then
                            AddToMatchCollection c, h, "ESP Match"
                        End If
                     End If
                End If
            Next
            pb.value = pb.value + 1
            DoEvents
      Next
      
      pb.value = 0
      
End Sub


Sub APIMatch()
    Dim i, j, t
    
    pb.value = 0
    Label1 = "API Matching"
    
     For Each c In a
        For Each h In b
            'not matched, same num of apicalls, within 15 bytes sizewise, and api called in same order
            If Not c.matched And Not h.matched Then
                If h.fxCalls.Count = c.fxCalls.Count And h.fxCalls.Count > 0 Then
                    'If isWithin(15, c.Length, h.Length) Then
                        j = 0
                        i = 0
                        For Each t In h.fxCalls
                            i = i + 1
                            If t = c.fxCalls(i) Then
                                j = j + 1
                            End If
                        Next
                        If j = h.fxCalls.Count Then
                            AddToMatchCollection c, h, "API Profile Match"
                        End If
                    'End If
                End If
            End If
            DoEvents
        Next
        pb.value = pb.value + 1
        DoEvents
    Next
    
    pb.value = 0
    
End Sub


Sub APIMatch2()
    Dim i, j, t, k
    
    pb.value = 0
    Label1 = "API2 Matching"
    
     For Each c In a
        For Each h In b
            If Not c.matched And Not h.matched Then
                If isWithin(4, h.fxCalls.Count, c.fxCalls.Count, 4) And _
                     isWithin(100, c.Length, h.Length) Then
                        j = 0
                        For Each t In h.fxCalls
                            For Each i In c.fxCalls
                                If t = i Then j = j + 1
                            Next
                        Next
                        If isWithin(4, j, h.fxCalls.Count, 3) Then
                            AddToMatchCollection c, h, "API Profile Match 2"
                        End If
                End If
            End If
            DoEvents
        Next
        pb.value = pb.value + 1
        DoEvents
    Next
    
    pb.value = 0
    
End Sub

Sub ConstMatch()
    Dim x, j
    
      pb.value = 0
      Label1 = "Const Matching"
      
      For Each c In a
            For Each h In b
                If Not c.matched And Not h.matched Then
                     If isWithin(3, c.Constants.Count, h.Constants.Count, 1) And _
                          isWithin(60, c.Length, h.Length) Then
                                j = 0
                                For Each x In c.Constants
                                   If h.ConstantExists(x) Then j = j + 1
                                Next
                                
                                If isWithin(3, c.Constants.Count, j, 2) Then
                                    AddToMatchCollection c, h, "Const Match"
                                End If
                                
                     End If
                End If
            Next
            pb.value = pb.value + 1
            DoEvents
      Next
      
      pb.value = 0
      
End Sub

Sub RunMatchSubs()

    Dim identifier
    Dim i As Long
    
    For i = 1 To 4
      pb.value = 0
      Label1 = "Running External Matchs"
      
      For Each c In a
            For Each h In b
                If Not c.matched And Not h.matched Then
                     If sc(1).Run("Match_" & i, c, h, identifier) = True Then
                         AddToMatchCollection c, h, CStr(identifier)
                     End If
                End If
            Next
            pb.value = pb.value + 1
            DoEvents
      Next
      
      pb.value = 0
    Next
      
End Sub


Sub AddToMatchCollection(match1 As CFunction, match2 As CFunction, method As String)
    m1.Add match1
    m2.Add match2
    match1.matched = True
    match2.matched = True
    match2.MatchMethod = method
    match1.MatchMethod = method
End Sub


Sub AddMatchs()
    Dim j As Long
    Dim t, u
    Dim li As ListItem
    
    For Each c In m1
            j = j + 1
            Set li = lvExact.ListItems.Add
            li.Tag = c.li.Tag & "," & m2(j).li.Tag
            
            li.Text = c.Name
            li.SubItems(1) = m2(j).Name
            li.SubItems(4) = c.MatchMethod
            
            t = c.Length
            u = m2(j).Length
            
            If t = u Then
                li.SubItems(2) = "yes"
            Else
                li.SubItems(2) = t & "," & u
            End If
            
            pb.value = pb.value + 1
            li.SubItems(3) = c.Calls & "/" & m2(j).Calls & " " & Hex(c.esp) & "/" & Hex(m2(j).esp)
    Next
     
   ResetPB a.Count, "Trimming A"
    
    For Each h In a
        If h.matched Then
            lv1.ListItems.Remove "id:" & h.autoid
        End If
        pb.value = pb.value + 1
    Next
    
    If Not SigMode Then
        ResetPB b.Count, "Trimming B"
        
        For Each h In b
            If h.matched Then
                lv2.ListItems.Remove "id:" & h.autoid
            End If
            pb.value = pb.value + 1
        Next
    End If
    
End Sub

Function LoadScript() As Boolean
    On Error GoTo hell
    If sc.Count = 2 Then Unload sc(1)
    Load sc(1)
    sc(1).AddCode ReadFile(App.path & "\compare.vbs")
    LoadScript = True
    Exit Function
hell:
    MsgBox "Error Loading script Line:" & sc(1).Error.Line & "Desc:" & vbCrLf & vbCrLf & sc(1).Error.Description
End Function

Private Sub Command2_Click()

    On Error Resume Next
    Dim pth As String
    Dim minFunctions As Long
    Dim j As Long
    Dim li As ListItem
    Dim t, u
    Dim r()
    
    Dim startTime As Long
    Dim endTime As Long
    
    GlobalResets
    startTime = GetTickCount
    
    If chkExternalMatchs.value = 1 Then
        If Not FileExists(App.path & "\compare.vbs") Then
            MsgBox "Could not locate compare.vbs for external match checks!", vbInformation
            Exit Sub
        End If
    End If
    
    If Len(currentMDB) = 0 Then
        cmndlg1.SetCustomFilter "Access Databases", "*.mdb"
        pth = cmndlg1.OpenDialog(CustomFilter)
    Else
        pth = currentMDB
        currentMDB = Empty
    End If
    
    If Len(pth) = 0 Then Exit Sub
    
    If Not FileExists(pth) Then
        MsgBox "Could not load: " & pth
        Exit Sub
    End If
    
    If cn.State <> 0 Then cn.Close
                         
    cn.Open "Provider=MSDASQL;Driver={Microsoft " & _
            "Access Driver (*.mdb)};DBQ=" & pth & ";"
    
    LoadList lv1, IIf(SigMode, TmpMode, compare1)  ', , " and index=16"
    LoadList lv2, IIf(SigMode, SignatureScan, compare2)  ', , " and index=21"
    
    push r, "Total functions " & lv1.ListItems.Count & ":" & lv2.ListItems.Count
    minFunctions = IIf(lv1.ListItems.Count > lv2.ListItems.Count, lv2.ListItems.Count, lv1.ListItems.Count)
    
    ResetPB lv1.ListItems.Count, "Comparing..."
    
    ExactCrcMatch
    
    If chkExternalMatchs.value = 1 Then
        If Not LoadScript() Then Exit Sub
        RunMatchSubs
    Else
        APIMatch
        EspMatch
        CallPushMatch
        APIMatch2
        ConstMatch
    End If

    ResetPB m1.Count, "Adding Matchs"
    AddMatchs
    
    If SigMode Then
        Label3 = "Current DB Functions (" & lv1.ListItems.Count & " Unmatched)                                       " & _
                 "Known Signatures "
    Else
        Label3 = "Unmatched Sample A  (" & lv1.ListItems.Count & " Unmatched)                                " & _
                 "Unmatched sample B (" & lv2.ListItems.Count & " Remaining)"
    End If
    
    Label1 = Empty
    pb.value = 0
    endTime = GetTickCount
    
    On Error Resume Next
    r(UBound(r)) = r(UBound(r)) & "  - made " & lvExact.ListItems.Count & " matchs"
    push r, "Percent:  " & CInt((lvExact.ListItems.Count / minFunctions) * 100) & "%"
    push r, "Elapsed Time: " & (endTime - startTime) \ 1000 & "secs"
    
    txtReport = Join(r, vbCrLf)
    
    Unload sc(1)
    
End Sub
 













Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    cn.Close
    Set cmndlg1 = Nothing
End Sub

Private Sub lblTransform_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    PopupMenu mnuPopupRename
End Sub



Public Sub lv1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rs As Recordset
    Dim asm As String
    
    Item.Selected = True
    Item.EnsureVisible
    
    Set rs = ado("Select * from a where autoid=" & Item.Tag)
    asm = rs!disasm
    txtA = asm
    
    Set sel_exact = Nothing
    Set sel_1 = Item
    If Not sel_2 Is Nothing Then
        cmdManualMatch.Enabled = True
        cmdBreakMatch.Enabled = False
    Else
        cmdBreakMatch.Enabled = False
    End If
    
    
    
    Me.caption = "Function list 1 " & lv1.ListItems.Count & " entries"
End Sub

Public Sub lv2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rs As Recordset
    Dim asm As String
    
    Item.Selected = True
    Item.EnsureVisible
    
    Set rs = ado("Select disasm from b where autoid=" & Item.Tag)
    asm = rs!disasm
    txtB = asm
    
    Set sel_exact = Nothing
    Set sel_2 = Item
    If Not sel_1 Is Nothing Then
        cmdManualMatch.Enabled = True
        cmdBreakMatch.Enabled = False
    Else
        cmdBreakMatch.Enabled = False
    End If
    
    Me.caption = "Function list 1 " & lv2.ListItems.Count & " entries"
    
End Sub



Private Sub lvExact_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Dim x
   On Error Resume Next
   
   x = Split(Item.Tag, ",")
   txtA = ado("Select disasm from a where autoid=" & x(0))!disasm
   txtB = ado("Select disasm from b where autoid=" & x(1))!disasm
   
   Set sel_exact = Item
   Set sel_1 = Nothing
   Set sel_2 = Nothing
   cmdManualMatch.Enabled = False
   cmdBreakMatch.Enabled = True
    
End Sub

Sub GlobalResets()
  
    Set m1 = New Collection
    Set m2 = New Collection
        
    Set a = New Collection
    Set b = New Collection
    
    lv1.ListItems.Clear
    lv2.ListItems.Clear
    txtA = Empty
    txtB = Empty
    txtReport = Empty
    lvExact.ListItems.Clear
    
End Sub

Private Sub lvExact_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuCheckAll_Click(index As Integer)
    
    On Error Resume Next
    Dim li As ListItem
    
Top:
    For Each li In lvExact.ListItems
        Select Case index
            Case 0: li.Selected = True
            Case 1: li.Selected = False
            Case 2: li.Selected = Not li.Selected
            Case 3: If li.Selected Then lvExact.ListItems.Remove li.index: GoTo Top
            Case 4: If Not li.Selected Then lvExact.ListItems.Remove li.index: GoTo Top
            Case 5:
                    If Len(li.Text) < 4 Then
                        li.Selected = False
                    Else
                        li.Selected = IIf(VBA.left(li.Text, 4) = "sub_", True, False)
                    End If
        End Select
    Next
    
End Sub

Private Sub mnuRename_Click(index As Integer)
    
    If index = 3 Then GoTo helpmsg
    
    If lvExact.ListItems.Count < 1 Then
        MsgBox "There are no matchs to port!", vbInformation
        Exit Sub
    End If
    
    Dim li As ListItem
    Dim tags() As String
    Dim i As Long
    Dim newName As String
    
    For Each li In lvExact.ListItems
        tags = Split(li.Tag, ",") 'autoid1, autoid2
        Select Case index
            Case 0: 'sequential rename of matchs - disabled for sigscan mode
                i = i + 1
                cn.Execute "Update a set newName='match_" & i & "' where autoid=" & tags(0)
                cn.Execute "Update b set newName='match_" & i & "' where autoid=" & tags(1)
            Case 1: 'port fx names from a->b - disabled for sigscan mode
                newName = li.Text
                If left(newName, 3) = "sub" Then newName = "imported_" & newName 'reserved
                cn.Execute "Update b set newName='" & newName & "' where autoid=" & tags(1)
            Case 2: 'port fx names from b->a
                newName = li.SubItems(1)
                If left(newName, 3) = "sub" Then newName = "imported_" & newName 'reserved
                cn.Execute "Update a set newName='" & newName & "' where autoid=" & tags(0)
        End Select
    Next
    
    MsgBox "Ok your mdb signature database has been updated with the changes." & vbCrLf & _
            "to apply the changes to the IDB disasm, launch the ida_compare plugin" & vbCrLf & _
            "and tell it to import the new names to the idb", vbInformation
            
    
    Exit Sub
helpmsg:

        MsgBox "These menu functions allow you to port names of matchs across dbs. To use, " & vbCrLf & _
                "trim the lower list using the check boxes and its right click menu until it contains" & vbCrLf & _
                "only the functions you want to see renamed." & vbCrLf & _
                "" & vbCrLf & _
                "For sequential renaming, all entries from both lists will be renamed match1, match2 etc" & vbCrLf & _
                "any user generated names will be overwritten. " & vbCrLf & _
                "" & vbCrLf & _
                "When you select to port the names, the corrosponding database record in the mdb" & vbCrLf & _
                "signature database will be updated with the new name to use. " & vbCrLf & _
                "" & vbCrLf & _
                "To apply the changes to the actual idb database, you will have to launch the IDA " & vbCrLf & _
                "compare plugin, and choose the import match names option." & vbCrLf
    
End Sub
