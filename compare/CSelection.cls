VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public selStart As Long
Public selLength As Long
Public wasBold As Long
Public orgSelColor As Long

Function LoadSel(rtf As RichTextBox)
    On Error Resume Next
    
    Me.orgSelColor = rtf.SelColor
    Me.selLength = rtf.selLength
    Me.selStart = rtf.selStart
    Me.wasBold = rtf.SelBold

End Function

Function Undo(rtf As RichTextBox)
     rtf.selStart = Me.selStart
     rtf.selLength = Me.selLength
     rtf.SelColor = Me.orgSelColor
     rtf.SelBold = Me.wasBold
     SelBackcolor rtf, vbWhite
End Function
