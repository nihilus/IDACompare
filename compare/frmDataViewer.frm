VERSION 5.00
Begin VB.Form frmDataViewer 
   Caption         =   "Data Viewer"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   LinkTopic       =   "Form2"
   ScaleHeight     =   4065
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   90
      Width           =   6225
   End
End
Attribute VB_Name = "frmDataViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function ShowData(x As String)
    txtData.Text = x
    Me.Visible = True
End Function

