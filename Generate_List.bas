'Written by Satoshi Coda.

Attribute VB_Name = "Generate_List"
Sub Generate_List()
    Dim buf As String, cnt As Long
    Dim str As String
    Dim sld As Slide
    Dim sh As Shape
    Dim extension() As String
    extension = Split("ppt*,pdf", ",")
    Dim ext As Variant
        
    Dim Path As String: Path = ActivePresentation.Path + "\"
    
    For Each ext In extension
        buf = Dir(Path & "*." & ext)
        Do While buf <> ""
            str = str + buf + vbCrLf
            buf = Dir()
        Loop
    Next
    
    Dim count As Integer
    Dim title As Variant
    count = 1
    For Each sld In ActivePresentation.Slides
        For Each sh In sld.Shapes
            '空のプレースホルダに入力
            If Not sh.TextFrame.HasText Then
                With sh.TextFrame.TextRange
                    With .ParagraphFormat
                        '.Bullet.Type = ppBulletUnnumbered
                        .SpaceAfter = 2
                    End With
                    .Text = str
                End With
                    
                For Each title In Split(str, vbCrLf)
                    If title = "" Then
                        Exit For
                    End If
                        
                    With sh.TextFrame.TextRange.Sentences(count)
                        .Text = Left(title, InStrRev(title, ".") - 1)
                        With .ActionSettings(ppMouseClick)
                            .Action = ppActionHyperlink
                            .Hyperlink.Address = Path & title
                        End With
                    End With
                    count = count + 1
                Next
                
                
            End If
        Next
    Next
    
End Sub
