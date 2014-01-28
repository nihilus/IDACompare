Attribute VB_Name = "Module1"
Dim fso As New CFileSystem2

'this is for debugging as an executable

Sub main()
    
    Dim w As New WinMergeScript
    Dim f As New frmVisualDebug
    Dim tmp As String
    
    tmp = fso.ReadFile(App.Path & "\test.txt")
    
    f.DebugFilter tmp, w
    
End Sub
