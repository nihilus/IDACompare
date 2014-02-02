VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "IDACompare"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   9195
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox frmConfigMatches 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Left            =   1860
      ScaleHeight     =   3615
      ScaleWidth      =   6735
      TabIndex        =   12
      Top             =   720
      Width           =   6795
      Begin VB.Frame frmConfigMatchesInner 
         Caption         =   " Configure Match Engine "
         Height          =   3465
         Left            =   60
         TabIndex        =   13
         Top             =   60
         Width           =   6585
         Begin VB.CheckBox chkEnforceMinSize 
            Caption         =   "Ignore functions < 30 Bytes"
            Height          =   285
            Left            =   2820
            TabIndex        =   33
            Top             =   2190
            Width           =   3675
         End
         Begin VB.Frame Frame2 
            Caption         =   " WinMerge Plugin Match Engine"
            Height          =   1605
            Left            =   1800
            TabIndex        =   28
            Top             =   420
            Width           =   4725
            Begin VB.OptionButton optWinMergeFilter 
               Caption         =   "Advanced"
               Height          =   435
               Index           =   3
               Left            =   240
               TabIndex        =   32
               Top             =   1050
               Width           =   2415
            End
            Begin VB.OptionButton optWinMergeFilter 
               Caption         =   "Debug Interface"
               Height          =   315
               Index           =   2
               Left            =   2460
               TabIndex        =   31
               Top             =   330
               Width           =   2205
            End
            Begin VB.OptionButton optWinMergeFilter 
               Caption         =   "Intermediate"
               Height          =   345
               Index           =   1
               Left            =   240
               TabIndex        =   30
               Top             =   690
               Width           =   1845
            End
            Begin VB.OptionButton optWinMergeFilter 
               Caption         =   "Basic"
               Height          =   315
               Index           =   0
               Left            =   240
               TabIndex        =   29
               Top             =   330
               Width           =   1005
            End
         End
         Begin VB.CheckBox chkExternalMatchScript 
            Caption         =   "External Match Script ( see compare.vbs )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   270
            TabIndex        =   27
            Top             =   2880
            Value           =   1  'Checked
            Width           =   3495
         End
         Begin VB.CheckBox chkConstMatch 
            Caption         =   "ConstMatch"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   270
            TabIndex        =   20
            Top             =   2490
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkApiMatch2 
            Caption         =   "ApiMatch2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   270
            TabIndex        =   19
            Top             =   2130
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkExactCRC 
            Caption         =   "ExactCRC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   270
            TabIndex        =   18
            Top             =   330
            Value           =   1  'Checked
            Width           =   1185
         End
         Begin VB.CheckBox chkCallPushMatch 
            Caption         =   "CallPush Match"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   270
            TabIndex        =   17
            Top             =   1770
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkEspMatch 
            Caption         =   "EspMatch"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   270
            TabIndex        =   16
            Top             =   1410
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkApiMatch 
            Caption         =   "ApiMatch"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   270
            TabIndex        =   15
            Top             =   1050
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox chkNameMatch 
            Caption         =   "Name Match"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   270
            TabIndex        =   14
            Top             =   720
            Value           =   1  'Checked
            Width           =   1245
         End
         Begin VB.Label lblCloseConfig 
            BackColor       =   &H00FFFFFF&
            Caption         =   " X"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5970
            TabIndex        =   21
            Top             =   30
            Width           =   615
         End
      End
   End
   Begin VB.Frame splitter 
      BackColor       =   &H00808080&
      Height          =   75
      Left            =   60
      MousePointer    =   7  'Size N S
      TabIndex        =   26
      Top             =   2760
      Width           =   9855
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4605
      Left            =   0
      TabIndex        =   3
      Top             =   4440
      Width           =   9885
      Begin VB.TextBox txtData 
         Height          =   3885
         Left            =   6030
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   34
         Top             =   630
         Width           =   3795
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
         Left            =   4590
         TabIndex        =   10
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
         Left            =   4950
         TabIndex        =   9
         Top             =   60
         Width           =   375
      End
      Begin VB.CommandButton cmdManualMatch 
         Caption         =   "Manual Match"
         Height          =   255
         Left            =   6270
         TabIndex        =   5
         Top             =   300
         Width           =   1815
      End
      Begin VB.CommandButton cmdBreakMatch 
         Caption         =   "Break Match"
         Enabled         =   0   'False
         Height          =   255
         Left            =   8190
         TabIndex        =   4
         Top             =   300
         Width           =   1485
      End
      Begin MSComctlLib.ListView lvExact 
         Height          =   3915
         Left            =   90
         TabIndex        =   6
         Top             =   600
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   6906
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name 1"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name 2"
            Object.Width           =   3528
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
      Begin VB.Label lblMatched 
         Caption         =   "Matched Functions"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   330
         Width           =   1665
      End
      Begin VB.Label lblTransform 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rename Tools"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3300
         TabIndex        =   7
         Top             =   300
         Width           =   2355
      End
   End
   Begin MSScriptControlCtl.ScriptControl sc 
      Index           =   0
      Left            =   8400
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   780
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   3413
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "i"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "sz"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "crc"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lv2 
      Height          =   1935
      Left            =   4950
      TabIndex        =   1
      Top             =   780
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3413
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "i"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "sz"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "crc"
         Object.Width           =   2540
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtB 
      Height          =   1545
      Left            =   4980
      TabIndex        =   23
      Top             =   2910
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   2725
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtA 
      Height          =   1545
      Left            =   30
      TabIndex        =   22
      Top             =   2880
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   2725
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":007C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   315
      Left            =   5580
      TabIndex        =   24
      Top             =   120
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   240
      TabIndex        =   25
      Top             =   120
      Width           =   5235
   End
   Begin VB.Label lblDBB 
      Caption         =   "Unmatched sample 2"
      Height          =   285
      Left            =   5010
      TabIndex        =   11
      Top             =   510
      Width           =   4935
   End
   Begin VB.Label lblDBA 
      Caption         =   "Unmatched Sample 1"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   540
      Width           =   4695
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuLoadDatabase 
         Caption         =   "Load New"
      End
      Begin VB.Menu mnuRescanCurrent 
         Caption         =   "Rescan Current DB"
      End
      Begin VB.Menu mnuLoadLast 
         Caption         =   "Open Last"
      End
      Begin VB.Menu mnuSpacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCfgEngine 
         Caption         =   "Configure Match Engine"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchFor 
         Caption         =   "Search For..."
      End
      Begin VB.Menu mnuProfileSelected 
         Caption         =   "Profile Selected Functions"
      End
      Begin VB.Menu mnuDecompileSelected 
         Caption         =   "Decompile Selected Functions"
      End
      Begin VB.Menu mnuSpacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCurExportForDiff 
         Caption         =   "WinMerge - Diff Disasm "
      End
      Begin VB.Menu mnuReconnectIDASrvr 
         Caption         =   "Reconnect To IDASrvr"
      End
   End
   Begin VB.Menu mnuLVPopup 
      Caption         =   "mnuLVPopup"
      Begin VB.Menu mnuTopCopyFuncNames 
         Caption         =   "Copy Function Names"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
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
      Begin VB.Menu mnuCopySelection 
         Caption         =   "Copy Selection"
      End
   End
   Begin VB.Menu mnuPopupRename 
      Caption         =   "mnuPopupRename"
      Begin VB.Menu mnuRename 
         Caption         =   "Sequentially Rename Matchs"
         Index           =   0
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Port user names from 1 to 2"
         Index           =   1
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Port user names from 2 to 1"
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

'note if we could remove matched entries from a/b collections after a match, then subsequent match
'  checks would have fewer functions to iteriate over (they arent checked again but they still have to be looped)
'  not sure the complexity is worth the optimization...
'
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


Dim selLV As ListView
Dim sel_1 As ListItem
Dim sel_2 As ListItem
Dim sel_exact As ListItem

Dim idaClient As New cIDAClient
Dim idaHwndA As Long
Dim idaHwndB As Long

Dim fullIDB_A As String
Dim fullIDB_B As String

Dim idb_a As String
Dim idb_b As String

Enum CompareModes
    compare1 = 0
    compare2 = 1
    SignatureScan = 2
    TmpMode = 3
End Enum

Private Capturing As Boolean
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long


'Private Sub frmConfigMatchesInner_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Capturing Then
'        ReleaseCapture
'        Capturing = False
'    End If
'End Sub
'
'Private Function CfgMoveOk(X&, Y&) As Boolean 'Put in any limiters you desire
'    CfgMoveOk = False
'    If Y > 0 And Y < Me.Height - Frame1.Height And _
'       X > 0 And X < Me.Width - frmConfigMatches.Width _
'    Then
'        CfgMoveOk = True
'    End If
'End Function
'
'Private Sub frmConfigMatchesInner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim x1 As Long, y1 As Long
'
'    If Button = 1 Then 'The mouse is down
'        If Capturing = False Then
'            SetCapture frmConfigMatches.hwnd
'            Capturing = True
'        End If
'        With frmConfigMatches
'            y1 = .Top + Y
'            x1 = .left + X
'            If CfgMoveOk(x1, y1) Then
'                .Top = y1
'                .left = x1
'            End If
'        End With
'    End If
'
'End Sub

Private Sub optWinMergeFilter_Click(index As Integer)
    SaveSetting "winmerge", "settings", "defaultFilter", index
End Sub

'splitter code
'------------------------------------------------
Private Sub splitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim a1&

    If Button = 1 Then 'The mouse is down
        If Capturing = False Then
            'splitter.ZOrder
            SetCapture splitter.hwnd
            Capturing = True
        End If
        With splitter
            a1 = .Top + y
            If MoveOk(a1) Then
                .Top = a1
            End If
        End With
    End If
End Sub

Private Sub splitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Capturing Then
        ReleaseCapture
        Capturing = False
        DoMove
    End If
End Sub


Private Sub DoMove()
    On Error Resume Next
    Const buf = 30
    txtA.Top = splitter.Top + splitter.Height + buf
    txtA.Height = Frame1.Top - txtA.Top
    txtB.Top = txtA.Top
    txtB.Height = txtA.Height
    lv1.Height = splitter.Top - lv1.Top - buf
    lv2.Height = lv1.Height
End Sub


Private Function MoveOk(y&) As Boolean  'Put in any limiters you desire
    MoveOk = False
    If y > lv1.Top + 1000 And y < Me.Height - (Frame1.Height * 1.5) Then
        MoveOk = True
    End If
End Function

'------------------------------------------------
'end splitter code


 
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

Private Sub Form_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim f As String
    f = data.Files(1)
    If fso.GetExtension(f) = ".mdb" Then
        currentMDB = f
        LoadDataBase f
    End If
End Sub

Private Sub lv1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set selLV = lv1
    If Button = 2 Then PopupMenu mnuLVPopup
End Sub

Private Sub lv2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set selLV = lv2
    If Button = 2 Then PopupMenu mnuLVPopup
End Sub


Private Sub lv2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LV_ColumnSort lv2, ColumnHeader
End Sub

Private Sub lv1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LV_ColumnSort lv1, ColumnHeader
End Sub

Private Sub lv1_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "p" Then
        mnuProfileSelected_Click
        KeyAscii = 0 'have to eat the keypress so that it doesnt auto find the first function starting with p
    End If
    If Chr(KeyAscii) = "d" Then
        mnuDecompileSelected_Click
        KeyAscii = 0
    End If
End Sub

Private Sub lv2_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "p" Then
        mnuProfileSelected_Click
        KeyAscii = 0 'have to eat the keypress so that it doesnt auto find the first function starting with p
    End If
    If Chr(KeyAscii) = "d" Then
        mnuDecompileSelected_Click
        KeyAscii = 0
    End If
End Sub

Private Sub lvExact_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     LV_ColumnSort lvExact, ColumnHeader
End Sub

Private Sub lvExact_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "p" Then
        mnuProfileSelected_Click
        KeyAscii = 0 'have to eat the keypress so that it doesnt auto find the first function starting with p
    End If
    If Chr(KeyAscii) = "d" Then
        mnuDecompileSelected_Click
        KeyAscii = 0
    End If
End Sub

Private Sub mnuLoadLast_Click()
    currentMDB = GetSetting("IDACompare", "settings", "lastMDB", currentMDB)
    LoadDataBase currentMDB
End Sub

Private Sub mnuProfileSelected_Click()
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

Private Sub mnuDecompileSelected_Click()

    On Error Resume Next
    Dim a As String
    Dim b As String
    Dim va As Long
    Dim funcA As String, funcB As String
     
    On Error Resume Next
    
    If Not sel_exact Is Nothing Then
    
        With idaClient
        
            If idaHwndA <> 0 Then
                 .ActiveIDA = idaHwndA
                 va = .FuncVAByName(sel_exact.Text)
                 If va <> 0 Then a = .Decompile(va)
            End If
        
            If idaHwndB <> 0 Then
                 .ActiveIDA = idaHwndB
                 va = .FuncVAByName(sel_exact.SubItems(1))
                 If va <> 0 Then b = .Decompile(va)
            End If
            
            If Len(a) > 0 Then txtA.Text = a
            If Len(b) > 0 Then txtB.Text = b
    
        End With
        Exit Sub
        
    End If
    
    If sel_1 Is Nothing And sel_2 Is Nothing Then Exit Sub
    If idaHwndA = 0 And idaHwndB = 0 Then Exit Sub
    
    If Not sel_1 Is Nothing Then funcA = sel_1.ListSubItems(2)
    If Not sel_2 Is Nothing Then funcB = sel_2.ListSubItems(2)
    
    With idaClient
        If idaHwndA <> 0 Then
             .ActiveIDA = idaHwndA
             va = .FuncVAByName(funcA)
             If va <> 0 Then a = .Decompile(va)
        End If
        
        If idaHwndB <> 0 Then
             .ActiveIDA = idaHwndB
             va = .FuncVAByName(funcB)
             If va <> 0 Then b = .Decompile(va)
        End If
        
    End With
    
    If Len(a) > 0 Then txtA.Text = a
    If Len(b) > 0 Then txtB.Text = b
    
        
    
End Sub

Private Sub mnuReconnectIDASrvr_Click()

    idaClient.EnumIDAWindows
    idaHwndA = idaClient.FindHwndForIDB(fullIDB_A)
    idaHwndB = idaClient.FindHwndForIDB(fullIDB_B)
    
    MsgBox "Connected to IDA for: " & fullIDB_A & "      " & IIf(idaHwndA = 0, "FAIL", "OK") & vbCrLf & _
           "Connected to IDA for: " & fullIDB_B & "      " & IIf(idaHwndB = 0, "FAIL", "OK") & vbCrLf & vbCrLf & _
           "If you are having problems make sure you manually installed the IDASrvr.plw" & vbCrLf & _
           "plugin to IDA and that you have the correct databases already open.", vbInformation
            
End Sub

Private Sub mnuRescanCurrent_Click()
    LoadDataBase currentMDB
End Sub

Private Sub mnuSearchFor_Click()
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
    If Me.Height < 9030 Then Me.Height = 9030
    If Me.Width < 10110 Then Me.Width = 10110
    
    Frame1.Top = Me.Height - Frame1.Height - 800
    txtA.Height = Frame1.Top - txtA.Top
    txtB.Height = txtA.Height
    
    txtA.Width = ((Me.Width / 2) - 40) - txtA.left
    txtB.left = txtA.left + txtA.Width + 20
    txtB.Width = Me.Width - txtB.left - 120
    lblDBB.left = txtB.left
    
    lv1.Width = txtA.Width
    lv2.Width = txtB.Width
    lv2.left = txtB.left
    lblDBA.Width = lv1.Width
    lblDBB.Width = lv2.Width
    
    lv1.ColumnHeaders(lv1.ColumnHeaders.Count - 1).Width = lv1.Width - lv1.ColumnHeaders(lv1.ColumnHeaders.Count - 1).left - 100 - lv1.ColumnHeaders(lv1.ColumnHeaders.Count).Width
    lv2.ColumnHeaders(lv2.ColumnHeaders.Count - 1).Width = lv2.Width - lv2.ColumnHeaders(lv2.ColumnHeaders.Count - 1).left - 100 - lv2.ColumnHeaders(lv2.ColumnHeaders.Count).Width
    
    Frame1.Width = Me.Width - 120
    splitter.Width = Frame1.Width
    lvExact.Width = Frame1.Width - 120 - txtData.Width - 120
    txtData.left = Frame1.Width - txtData.Width - 120 - Frame1.left
    pb.Width = Me.Width - pb.left - 200
    
    Command1(1).left = txtB.left - Command1(1).Width
    Command1(0).left = txtB.left
    
    If splitter.Top < lv1.Top + 1000 Then
        splitter.Top = lv1.Top + 1000
        DoMove
    ElseIf splitter.Top > Me.Height - (Frame1.Height * 1.5) Then
        splitter.Top = Me.Height - (Frame1.Height * 1.5)
        DoMove
    End If
    
    
End Sub

Private Sub mnuCfgEngine_Click()
    
    On Error Resume Next
    
    With frmConfigMatches
        .ZOrder
        .left = Me.Width / 2 - .Width / 2
        .Top = Me.Height / 2 - .Height / 2
        .Visible = True
    End With
    
End Sub

Private Sub lblCloseConfig_Click()
    frmConfigMatches.Visible = False
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

Sub LoadChkSettings(Optional load As Boolean = True)
    
    Dim cc As CheckBox
    Dim c As Control
    Dim r As Long
    Dim defVal As Long
    
    On Error Resume Next
    For Each c In Me.Controls
        If TypeName(c) = "CheckBox" Then
            Set cc = c
            defVal = 1
            If cc.Name = chkExternalMatchScript.Name Then defVal = 0
            If load Then
                r = GetSetting("IDACompare", "settings", cc.Name, defVal)
                cc.value = r
            Else
                Call SaveSetting("IDACompare", "settings", cc.Name, cc.value)
            End If
        End If
    Next
        
End Sub

Function isIDE() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    isIDE = CBool(Err.Number)
End Function

'if we have a function calling out to a matched function
'and it was the only way there..then if we only had one
'other unmatched on calling to it them we could relate them.

Private Sub Form_Load()
    Dim cmd As String
    Dim filtIndex As Long
    
    On Error Resume Next
    
    idaClient.Listen Me.hwnd
    mnuPopup.Visible = False
    mnuPopupRename.Visible = False
    mnuLVPopup.Visible = False
    frmConfigMatches.Visible = False
    FormPos Me, True
    splitter.Top = GetSetting("IDACompare", "settings", "SplitterTop", splitter.Top)
    DoMove
    Form_Resize
    
    filtIndex = GetSetting("winmerge", "settings", "defaultFilter", 1)
    optWinMergeFilter(filtIndex).value = True
    LoadChkSettings
    
    cmd = Command
    'If isIDE() Then cmd = App.path & "\..\mydoom_example.mdb"
    
    If Len(cmd) > 0 Then
        currentMDB = Replace(cmd, """", Empty)
        If InStr(1, currentMDB, "/sigscan", vbTextCompare) > 1 Then
            SigMode = True
            mnuRename(0).Enabled = False
            mnuRename(1).Enabled = False
            mnuLoadDatabase.Enabled = False
            cmdManualMatch.Visible = False
            cmdBreakMatch.value = False
            currentMDB = Trim(Replace(currentMDB, "/sigscan", Empty))
        End If
        If Not FileExists(currentMDB) Then
            MsgBox "Usage: ida_compare.exe <mdb path to analyze>" & vbCrLf & vbCrLf & currentMDB
            currentMDB = Empty
        Else
            Me.Visible = True
            LoadDataBase currentMDB
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
    Dim idb As String
    
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
    
    If mode = compare1 Then
        idb_a = ado("Select top 1 idb from " & tbl)!idb
        fullIDB_A = idb_a
        If Len(idb_a) > 12 Then idb_a = Mid(idb_a, 1, 10) & "..."
        lblDBA = "1: " & idb_a
    ElseIf mode = compare2 Then
        idb_b = ado("Select top 1 idb from " & tbl)!idb
        fullIDB_B = idb_b
        If Len(idb_b) > 12 Then idb_b = Mid(idb_b, 1, 10) & "..."
        lblDBB = "2: " & idb_b
    End If
    
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
            'c.ReHash asm
            'If KeyExistsInCollection(IIf(isTableA, a, b), c.mCRC) Then
                c.mCRC = HandleCRCDuplicate(IIf(isTableA, a, b), c.mCRC) 'using just this instead of rehash is much much better..
                'While KeyExistsInCollection(IIf(isTableA, a, b), c.mCRC)
                '    c.mCRC = "rand:" & RandomNum
                'Wend
            'End If
        End If

        If Len(c.mCRC) > 0 Then
            Set li = lv.ListItems.Add(, "id:" & rs!autoid)
            Set c.li = li

            c.Length = rs!leng
            c.autoid = rs!autoid
            c.Name = rs!fname
            c.index = rs!index

            li.Tag = c.autoid
            li.Text = pad(rs!index, 3)
            li.SubItems(1) = pad(c.Length)
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

Private Function HandleCRCDuplicate(c As Collection, baseCrc As String) As String
    
    Dim tmp As String
    Dim i As Long
    
    Do
        i = i + 1
        If i > 3000 Then
            tmp = "rand:" & RandomNum
        Else
            tmp = baseCrc & "_" & i
        End If
    Loop While KeyExistsInCollection(c, tmp)
    
    HandleCRCDuplicate = tmp
    
End Function

Function ExactCrcMatch() As Long
    
    Dim lit As ListItem
    Dim ret As Long
    
    Label1 = "CRC Matching"
    
    For Each lit In lv1.ListItems
        If KeyExistsInCollection(b, lit.SubItems(3)) Then
            Set c = a(lit.SubItems(3))
            Set h = b(lit.SubItems(3))
            AddToMatchCollection c, h, "Exact CRC"
            ret = ret + 1
        End If
        pb.value = pb.value + 1
    Next
    
    ExactCrcMatch = ret
    
End Function

Function NameMatch() As Long

    Dim ret As Long
    pb.value = 0
    Label1 = "Public Name Matching"
    
    For Each c In a
        For Each h In b
            If Not c.matched And Not h.matched Then
                 If c.Name = h.Name _
                    And VBA.left(c.Name, 4) <> "sub_" _
                    And VBA.left(c.Name, 4) <> "unkn" _
                    And c.Name <> "start" _
                 Then
                    AddToMatchCollection c, h, "Name Match"
                    ret = ret + 1
                 End If
            End If
        Next
        pb.value = pb.value + 1
        DoEvents
    Next
    
    pb.value = 0
    NameMatch = ret
    
End Function

Function CallPushMatch() As Long

    Dim ret As Long
    pb.value = 0
    Label1 = "Call/Push Matching"
    
    For Each c In a
        For Each h In b
            If Not c.matched And Not h.matched Then
                If c.Calls = h.Calls And c.Pushs = h.Pushs Then  'same num of calls and pushs
                    If isWithin(60, c.Length, h.Length, 80) Then     'and length is close
                        If isWithin(4, c.Jumps, h.Jumps) Then    'and num jmps is close
                           AddToMatchCollection c, h, "Call/Push Match"
                           ret = ret + 1
                        End If
                    End If
                 End If
            End If
        Next
        pb.value = pb.value + 1
        DoEvents
    Next
    
    pb.value = 0
    CallPushMatch = ret
    
End Function

Function EspMatch() As Long

      Dim ret As Long
      pb.value = 0
      Label1 = "ESP Matching"
      
      For Each c In a
            For Each h In b
                If Not c.matched And Not h.matched Then
                     If isWithin(80, c.Length, h.Length, 80) Then
                        If c.esp <> 0 And c.esp = h.esp And isWithin(40, c.Length, h.Length) Then
                            AddToMatchCollection c, h, "ESP Match"
                            ret = ret + 1
                        End If
                     End If
                End If
            Next
            pb.value = pb.value + 1
            DoEvents
      Next
      
      pb.value = 0
      EspMatch = ret
      
End Function


Function APIMatch() As Long
    Dim i, j, t
    
    Dim ret As Long
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
                            ret = ret + 1
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
    APIMatch = ret
    
End Function


Function APIMatch2() As Long
    Dim i, j, t, k
    
    Dim ret As Long
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
                            ret = ret + 1
                        End If
                End If
            End If
            DoEvents
        Next
        pb.value = pb.value + 1
        DoEvents
    Next
    
    pb.value = 0
    APIMatch2 = ret
    
End Function

Function ConstMatch() As Long
    Dim x, j
    
      Dim ret As Long
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
                                    ret = ret + 1
                                End If
                                
                     End If
                End If
            Next
            pb.value = pb.value + 1
            DoEvents
      Next
      
      pb.value = 0
      ConstMatch = ret
      
End Function

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
    load sc(1)
    sc(1).AddCode ReadFile(App.path & "\compare.vbs")
    LoadScript = True
    Exit Function
hell:
    MsgBox "Error Loading script Line:" & sc(1).Error.Line & "Desc:" & vbCrLf & vbCrLf & sc(1).Error.Description
End Function

Private Sub mnuLoadDatabase_click()
    Dim pth As String
    cmndlg1.SetCustomFilter "Access Databases", "*.mdb"
    pth = cmndlg1.OpenDialog(CustomFilter, , , Me.hwnd)
    If Len(pth) = 0 Then Exit Sub
    currentMDB = pth
    LoadDataBase currentMDB
End Sub

Sub LoadDataBase(pth As String)

    On Error Resume Next
    Dim minFunctions As Long
    Dim j As Long
    Dim li As ListItem
    Dim t, u
    Dim r()
    
    Dim startTime As Long
    Dim endTime As Long
    
    GlobalResets
    startTime = GetTickCount
    lblMatched = "Matched"
    
    If chkExternalMatchScript.value = 1 Then
        If Not FileExists(App.path & "\compare.vbs") Then
            MsgBox "Could not locate compare.vbs for external match checks!", vbInformation
            Exit Sub
        End If
    End If
    
    If Len(pth) = 0 Then Exit Sub
    
    If Not FileExists(pth) Then
        MsgBox "Could not load: " & pth
        Exit Sub
    End If
    
    If cn.State <> 0 Then cn.Close
                         
    cn.Open "Provider=MSDASQL;Driver={Microsoft " & _
            "Access Driver (*.mdb)};DBQ=" & pth & ";"
    
    Dim minLength As Long
    minLength = IIf(chkEnforceMinSize.value = 1, 30, 0)
    
    LoadList lv1, IIf(SigMode, TmpMode, compare1), minLength ', , " and index=16"
    LoadList lv2, IIf(SigMode, SignatureScan, compare2), minLength ', , " and index=21"
    
    push r, "Total functions " & lv1.ListItems.Count & ":" & lv2.ListItems.Count
    minFunctions = IIf(lv1.ListItems.Count > lv2.ListItems.Count, lv2.ListItems.Count, lv1.ListItems.Count)
    
    ResetPB lv1.ListItems.Count, "Comparing..."
    
    Dim matches As Long
    Dim stats() As String
    
    If chkExactCRC.value = 1 Then
        matches = ExactCrcMatch()
        push stats(), "ExactCrc Matches: " & matches
    End If
    
    If chkNameMatch.value = 1 Then
        matches = NameMatch()
        push stats(), "Name Matches: " & matches
    End If
    
    If chkApiMatch.value = 1 Then
        matches = APIMatch()
        push stats(), "API Matches: " & matches
    End If
    
    If chkEspMatch.value = 1 Then
        matches = EspMatch()
        push stats(), "ESP Matches: " & matches
    End If
    
    If chkCallPushMatch.value = 1 Then
        matches = CallPushMatch()
        push stats(), "CallPush Matches: " & matches
    End If
    
    If chkApiMatch2.value = 1 Then
        matches = APIMatch2()
        push stats(), "API2 Matches: " & matches
    End If
    
    If chkConstMatch.value = 1 Then
        matches = ConstMatch()
        push stats(), "Const Matches: " & matches
    End If
        
        
    If chkExternalMatchScript.value = 1 Then
        If LoadScript() Then RunMatchSubs
    End If
 

    ResetPB m1.Count, "Adding Matchs"
    AddMatchs
    
    If SigMode Then
        lblDBA = idb_a & " Funcs (" & lv1.ListItems.Count & " Unmatched)"
        lblDBB = "Known Signatures "
    Else
        lblDBA = "Unmatched 1: " & idb_a & " (" & lv1.ListItems.Count & " Unmatched) "
        lblDBB = "Unmatched 2: " & idb_b & " (" & lv2.ListItems.Count & " Remaining)"
    End If
    
    Label1 = Empty
    pb.value = 0
    endTime = GetTickCount
    
    Dim pcent As String
    
    On Error Resume Next
    pcent = CInt((lvExact.ListItems.Count / minFunctions) * 100) & "%"
    Label1 = pcent & " similarity. See stats for details"
    
    r(UBound(r)) = r(UBound(r)) & vbCrLf & "Total Matches: " & lvExact.ListItems.Count
    push r, "Percent:  " & pcent
    push r, "Elapsed Time: " & (endTime - startTime) \ 1000 & "secs"
    
    txtData = Join(r, vbCrLf) & vbCrLf & Join(stats, vbCrLf)
    
    idaClient.EnumIDAWindows
    idaHwndA = idaClient.FindHwndForIDB(fullIDB_A)
    idaHwndB = idaClient.FindHwndForIDB(fullIDB_B)
    mnuDecompileSelected.Enabled = idaClient.DecompilerActive(idaHwndA)
        
    lblMatched = "Matched: " & lvExact.ListItems.Count
    
    Unload sc(1)
    
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    cn.Close
    FormPos Me, True, True
    LoadChkSettings False
    SaveSetting "IDACompare", "settings", "SplitterTop", splitter.Top
    SaveSetting "IDACompare", "settings", "lastMDB", currentMDB
    Set cmndlg1 = Nothing
    
    Dim f
    For Each f In Forms
        Unload f
    Next
End Sub

Private Sub lblTransform_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    PopupMenu mnuPopupRename
End Sub



Public Sub lv1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rs As Recordset
    Dim asm As String
    Dim c As CFunction
    
    Item.Selected = True
    Item.EnsureVisible
    
    Set rs = ado("Select * from a where autoid=" & Item.Tag)
    asm = rs!disasm
    'txtA = asm
    
    Set c = a(Item.ListSubItems(3))
    rtfHighlightAsm asm, c, txtA
    
    Set sel_exact = Nothing
    Set sel_1 = Item
    If Not sel_2 Is Nothing Then
        cmdManualMatch.Enabled = True
        cmdBreakMatch.Enabled = False
    Else
        cmdBreakMatch.Enabled = False
    End If
    
    If idaHwndA <> 0 Then
        idaClient.JumpName Item.SubItems(2), idaHwndA
    End If
    
    Me.caption = "Function list 1 " & lv1.ListItems.Count & " entries"
End Sub

Public Sub lv2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rs As Recordset
    Dim asm As String
    Dim c As CFunction
    
    Item.Selected = True
    Item.EnsureVisible
    
    Set rs = ado("Select disasm from b where autoid=" & Item.Tag)
    asm = rs!disasm
    'txtB = asm
    
    Set c = b(Item.ListSubItems(3))
    rtfHighlightAsm asm, c, txtB
    
    Set sel_exact = Nothing
    Set sel_2 = Item
    If Not sel_1 Is Nothing Then
        cmdManualMatch.Enabled = True
        cmdBreakMatch.Enabled = False
    Else
        cmdBreakMatch.Enabled = False
    End If
    
    If idaHwndB <> 0 Then
        idaClient.JumpName Item.SubItems(2), idaHwndB
    End If
    
    Me.caption = "Function list 1 " & lv2.ListItems.Count & " entries"
    
End Sub



Private Sub lvExact_ItemClick(ByVal Item As MSComctlLib.ListItem)
   Dim x, asmA As String, asmB As String
   On Error Resume Next
   
   If Not sel_exact Is Nothing Then
        If sel_exact = Item Then Exit Sub
   End If
   
   x = Split(Item.Tag, ",")
   asmA = ado("Select disasm from a where autoid=" & x(0))!disasm
   asmB = ado("Select disasm from b where autoid=" & x(1))!disasm
   
   'Set c = a(Item.ListSubItems(3))
   rtfHighlightAsm asmA, Nothing, txtA
   rtfHighlightAsm asmB, Nothing, txtB
    
   Set sel_exact = Item
   Set sel_1 = Nothing
   Set sel_2 = Nothing
   cmdManualMatch.Enabled = False
   cmdBreakMatch.Enabled = True
    
    If idaHwndA <> 0 Then
        idaClient.JumpName Item.Text, idaHwndA
    End If
    
    If idaHwndB <> 0 Then
        idaClient.JumpName Item.SubItems(1), idaHwndB
    End If
    
End Sub

Function FindMatchAutoID(funcName As String, isTableA As Boolean) As Long
    
    Dim li As ListItem
    Dim fn As String
    Dim x
    
    On Error Resume Next
    
    For Each li In lvExact.ListItems
        If isTableA Then fn = li.Text Else fn = li.SubItems(1)
        If fn = funcName Then
            x = Split(li.Tag, ",")
            FindMatchAutoID = IIf(isTableA, x(0), x(1))
            Exit Function
        End If
    Next
    
    
End Function

Sub GlobalResets()
  
    Set m1 = New Collection
    Set m2 = New Collection
        
    Set a = New Collection
    Set b = New Collection
    
    lv1.ListItems.Clear
    lv2.ListItems.Clear
    txtA = Empty
    txtB = Empty
    txtData = Empty
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
    
    If index = 3 Or index = 4 Then 'removes
        MsgBox lvExact.ListItems.Count & " matches left", vbInformation
    End If
    
End Sub

Private Sub mnuCopySelection_Click()
    Clipboard.Clear
    Clipboard.SetText GetAllElements(lvExact, True)
End Sub

Private Sub mnuCurExportForDiff_Click()
    
    Dim exe As String
    Dim dll As String
    Dim srcDll As String

    Dim tmp As String
    Dim a As String
    Dim b As String
   
    On Error GoTo hell
   
    tmp = Environ("temp")
   
    If Not fso.FolderExists(tmp) Then
        MsgBox "Temp envirnoment variable not set", vbInformation
        Exit Sub
    End If
   
    a = tmp & "\a.idacompare"
    b = tmp & "\b.idacompare"
   
    dll = "C:\Program Files\WinMerge\MergePlugins\wmIDACompare.dll"
    exe = "C:\Program Files\WinMerge\WinMergeU.exe"
    srcDll = App.path & IIf(isIDE(), "\..\", "") & "\WinMerge_Plugin\wmIDACompare.dll"
   
    If Len(txtA.Text) = 0 Or Len(txtB.Text) = 0 Then
        MsgBox "You must select two functions to diff first."
        Exit Sub
    End If
   
    If Not fso.FileExists(exe) Then
        MsgBox "WinMerge not found. Please download from winmerge.org to use this feature." & _
                vbCrLf & vbCrLf & "Expected path: " & exe, vbInformation
        Exit Sub
    End If
   
    If Not fso.FileExists(dll) Then
        FileCopy srcDll, dll
        MsgBox "The IDACompare Winmerge plugin has been automatically installed. You will have to enable plugins and then run it manually from the prediffer menu.", vbInformation
    End If
   
    If wHash.HashFile(dll) <> wHash.HashFile(srcDll) Then
        FileCopy srcDll, dll
        Me.caption = "The IDACompare Winmerge plugin has been updated."
    End If
   
    writeFile a, txtA.Text
    writeFile b, txtB.Text
    Shell exe & " " & a & " " & b, vbNormalFocus
     
    'To active prediffer automatically send the following keys:
    '(did i mention how much i love vb6 for its simplicity? :)
    AppActivate "WinMerge - [a.idacompare - b.idacompare]"
    SendKeys "%p"         'alt-p
    SendKeys "{DOWN 4}"   'down 4
    SendKeys "{RIGHT}"    'right 1
    SendKeys "{DOWN 2}"   'down two
    SendKeys vbCr         'return
    
    Exit Sub
hell:
        MsgBox Err.Description
        
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

 
Private Sub mnuTopCopyFuncNames_Click()
    On Error Resume Next
    Dim li As ListItem
    Dim tmp As String
    
    If selLV Is Nothing Then Exit Sub
    
    For Each li In selLV.ListItems
        tmp = tmp & li.SubItems(2) & vbCrLf
    Next
    
    Clipboard.Clear
    Clipboard.SetText tmp
End Sub
