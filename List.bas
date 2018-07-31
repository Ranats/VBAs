Attribute VB_Name = "List"
Private Sub List()
    Dim buf As String, cnt As Long
    Dim str As String
    Dim sld As Slide
    Dim sh As Shape
    Dim extension() As String
    extension = Split("ppt*,pdf", ",")
    Dim ext As Variant
        
    Const Path As String = "./"
    For Each ext In extension
        buf = Dir(Path & "*." & ext)
        Do While buf <> ""
            str = str + buf + vbCrLf
            buf = Dir()
        Loop
    Next
    
    For Each sld In ActivePresentation.Slides
        For Each sh In sld.Shapes
            '空のプレースホルダに入力
            If Not sh.TextFrame.HasText Then
                sh.TextFrame.TextRange.Text = str
            End If
        Next
    Next
    
End Sub

