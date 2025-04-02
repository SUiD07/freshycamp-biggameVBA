Sub Togglenode1()
    Dim slide As slide
    Dim shp As Shape
    Dim maxZ As Integer

    Set slide = ActivePresentation.Slides(2)
    On Error Resume Next
    Set shp = slide.Shapes("node 1")
    On Error GoTo 0

    If Not shp Is Nothing Then
        maxZ = 0
        Dim s As Shape
        For Each s In slide.Shapes
            If s.ZOrderPosition > maxZ Then maxZ = s.ZOrderPosition
        Next s

        If shp.ZOrderPosition = maxZ Then
            shp.ZOrder msoSendToBack
        Else
            shp.ZOrder msoBringToFront
        End If
    Else
        MsgBox "cannot find node 1", vbExclamation
    End If
End Sub


