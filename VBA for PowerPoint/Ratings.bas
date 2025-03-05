Attribute VB_Name = "Ratings"
Enum Rating
    Undef = 0
    Low = 1
    Medium = 2
    High = 3
End Enum

Dim colorDict As Object

Private Sub CreateColorDict()
    Set colorDict = CreateObject("Scripting.Dictionary")
    colorDict.Add Rating.Low, rgb(54, 164, 29)
    colorDict.Add Rating.Medium, rgb(255, 192, 0)
    colorDict.Add Rating.High, rgb(255, 0, 0)
    colorDict.Add Rating.Undef, rgb(255, 255, 0) ' error
End Sub

Function ColorRGB(rat As Rating)
    If colorDict Is Nothing Then
        CreateColorDict
    End If
    ColorRGB = colorDict(rat)
End Function
