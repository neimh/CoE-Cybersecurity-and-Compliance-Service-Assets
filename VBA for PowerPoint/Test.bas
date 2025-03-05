Attribute VB_Name = "Test"
'Sub Test()
'    Dim find As finding
'    Dim instring As String
'    Set find = New finding
'    Debug.Print find.priorityPretty()
'    find.priorityParse ("what High")
'    Debug.Print find.priorityPretty()
'End Sub
'
'Sub TestRegExpLookbehind()
'    Dim testString As String
'    Static re As Object
'    testString = "lookbehind #01"
'
'    Set re = CreateObject("VBScript.RegExp")
'    re.Global = True
'    re.Multiline = True
'    re.Pattern = "lookbehind #(\d\d)"
'
'    Set matches = re.Execute(testString)
'    Debug.Print matches(0).SubMatches(0)
'End Sub
'


Sub TestParseBodyForFinding()
    Dim inputText As String
    inputText = "Description:" & vbCrLf & _
                "test description" & vbCrLf & _
                "Business impact:" & vbCrLf & _
                "test business impact" & vbCrLf & _
                "Recommended actions:" & vbCrLf & _
                "test recommended actions"
    
    ' Call the subroutine
    Dim fnd As finding
    Set fnd = New finding
    fnd.parseBodyForFinding inputText
End Sub
