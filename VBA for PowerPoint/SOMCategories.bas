Attribute VB_Name = "SOMCategories"
Dim colorDict As Object

Sub CreateColorDictionary()
    Set colorDict = CreateObject("Scripting.Dictionary")
    ' Add the categories and their corresponding RGB colors
    colorDict.Add "Awareness", Array(118, 10, 133)
    colorDict.Add "Security Governance", Array(118, 10, 133)
    colorDict.Add "Risk Management", Array(118, 10, 133)
    colorDict.Add "Regulatory Process Compliance", Array(240, 171, 0)
    colorDict.Add "Data Privacy & Protection", Array(240, 171, 0)
    colorDict.Add "Audit & Fraud Management", Array(240, 171, 0)
    colorDict.Add "User & Identity Management", Array(227, 85, 0)
    colorDict.Add "Custom Code Security", Array(227, 85, 0)
    colorDict.Add "Roles & Authorizations", Array(227, 85, 0)
    colorDict.Add "Authentication & Single Sign-On", Array(227, 85, 0)
    colorDict.Add "Security Hardening", Array(78, 184, 28)
    colorDict.Add "Secure SAP Code", Array(79, 184, 28)
    colorDict.Add "Security Monitoring & Forensics", Array(79, 184, 28)
    colorDict.Add "Network Security", Array(102, 102, 102)
    colorDict.Add "Operating System & Database Security", Array(102, 102, 102)
    colorDict.Add "Client Security", Array(102, 102, 102)
End Sub

Function ColorRGB(categoryText As String)
    If colorDict Is Nothing Then
        CreateColorDictionary
    End If

    ColorRGB = colorDict(categoryText)
End Function


Sub GetColors()
    Dim slide As slide
    Dim shape As shape
    Dim fillColor As Long
    Dim shapeText As String
    
    ' Get the specific slide (slide 17 in this case)
    Set slide = ActivePresentation.Slides(17)
    
    ' Loop through all shapes on the slide
    For Each shape In slide.Shapes
        ' Initialize variables
        fillColor = -1 ' Default value if no fill is found
        shapeText = "" ' Default value if no text is found
        
        ' Check if the shape has a fill
        If shape.Fill.Visible Then
            fillColor = shape.Fill.ForeColor.rgb
        End If
        
        ' Check if the shape has text
        If shape.HasTextFrame Then
            If shape.TextFrame.HasText Then
                shapeText = shape.TextFrame.TextRange.text
            End If
        End If
        
        ' Output the results to the Immediate Window
        Debug.Print GetRGBComponents(fillColor) & vbTab & vbTab & shapeText
    Next shape
End Sub

Function GetRGBComponents(rgbValue As Long) As String
    Dim r As Long
    Dim g As Long
    Dim b As Long
    
    ' Extract the red, green, and blue components
    r = rgbValue Mod 256
    g = (rgbValue \ 256) Mod 256
    b = (rgbValue \ 65536) Mod 256
    
    ' Format as "RGB(r, g, b)"
    GetRGBComponents = "RGB(" & r & ", " & g & ", " & b & ")"
End Function
