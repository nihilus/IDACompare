VERSION 5.00
Begin VB.Form frmProfile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Function  Profile Details"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
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
      Left            =   3900
      TabIndex        =   3
      Top             =   4260
      Width           =   375
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
      Left            =   3540
      TabIndex        =   2
      Top             =   4260
      Width           =   375
   End
   Begin VB.TextBox txtDetails2 
      Height          =   4155
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   60
      Width           =   3735
   End
   Begin VB.TextBox txtDetails1 
      Height          =   4155
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   3735
   End
End
Attribute VB_Name = "frmProfile"
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

Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_TOPMOST = -1



Sub SetWindowTopMost(f As Form)
   SetWindowPos f.hWnd, HWND_TOPMOST, f.left / 15, _
        f.Top / 15, f.Width / 15, _
        f.Height / 15, Empty
End Sub


Sub ShowProfile(a As CFunction, Optional b As CFunction = Nothing)
    
    If b Is Nothing Then
        Me.Width = 3870  'single view mode
        Me.Height = 4635 'hide dual scroll buttons
    End If
    
    txtDetails1 = BuildReport(a)
    
    If Not b Is Nothing Then
        txtDetails2 = BuildReport(b)
    End If
    
    Me.Visible = True
    SetWindowTopMost Me
    
End Sub

Private Sub Command1_Click(index As Integer)
        
    ScrollPage txtDetails1, txtDetails2, CBool(index)
    
End Sub

Private Function BuildReport(x As CFunction) As String
    
    Dim r() As String
    Dim l As String
    
    l = vbCrLf & String(30, "-") & vbCrLf
    
    With x
        push r, "Function Name: " & x.Name
        push r, "Length: " & x.Length
        push r, "Calls: " & x.Calls
        push r, "Jmps: " & x.Jumps
        push r, "ESP: " & Hex(x.esp) & "h"
        push r, "CRC: " & x.mCRC & vbCrLf
        push r, "Consts:" & l & DumpCollection(x.Constants) & vbCrLf
        push r, "Fx Calls:" & l & DumpCollection(x.fxCalls) & vbCrLf
    End With
        
    BuildReport = Join(r, vbCrLf)
    
End Function

Private Function DumpCollection(c As Collection, Optional delimiter As String = vbCrLf)
    Dim x
    Dim ret As String
    For Each x In c
        ret = ret & x & delimiter
    Next
    DumpCollection = ret
End Function

