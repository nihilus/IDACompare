VERSION 5.00
Begin VB.Form frmVisualDebug 
   Caption         =   "Visual Debug Winmerge Filter"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   13110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMoreComplex 
      Caption         =   "More Complex Signature"
      Height          =   405
      Left            =   4260
      TabIndex        =   7
      Top             =   8070
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Accept and continue"
      Height          =   435
      Left            =   10590
      TabIndex        =   5
      Top             =   8100
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Less Agressive / Signature"
      Height          =   405
      Left            =   1680
      TabIndex        =   4
      Top             =   8070
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agressive / basic"
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Top             =   8070
      Width           =   1365
   End
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7545
      Left            =   6600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   450
      Width           =   6405
   End
   Begin VB.TextBox txtIn 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7545
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   420
      Width           =   6405
   End
   Begin VB.Label Label2 
      Caption         =   "Transformed return buffer (do not change line numbers!)"
      Height          =   255
      Left            =   6630
      TabIndex        =   6
      Top             =   150
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Incoming Buffer"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   1935
   End
End
Attribute VB_Name = "frmVisualDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim parent As WinMergeScript

Function DebugFilter(bufIn As String, c As WinMergeScript) As String
    Set parent = c
    
    txtIn.text = bufIn
    Me.Show 1 'modal function doesnt return until hidden
    
    DebugFilter = txtOut.text
    Unload Me
End Function

Private Sub cmdMoreComplex_Click()
    txtOut = parent.ComplexSignature(txtIn)
End Sub

Private Sub Command1_Click()
    txtOut = parent.Aggressive(txtIn, False, True)
End Sub

Private Sub Command2_Click()
    txtOut = parent.LessAggressive(txtIn)
End Sub

Private Sub Command3_Click()
    Me.Hide 'breaks modal show
End Sub
