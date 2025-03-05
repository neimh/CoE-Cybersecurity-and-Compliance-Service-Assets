Attribute VB_Name = "Txt"
Public Function SplitRe(text As String, Pattern As String, Optional IgnoreCase As Boolean) As String()
    Static re As Object

    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.Multiline = True
    End If

    re.IgnoreCase = IgnoreCase
    re.Pattern = Pattern
    SplitRe = Strings.split(re.Replace(text, ChrW(-1)), ChrW(-1))
End Function
