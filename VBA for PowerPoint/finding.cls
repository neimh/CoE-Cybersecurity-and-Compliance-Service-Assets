VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Finding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Private pPriority As Rating
Private pEffort As Rating
Private pNumberText As String
Private pCategoryText
Private issueTitleText As String
Private pDescription As String
Public ready As Boolean
Public titleReady As Boolean
Public bodyReady As Boolean
Public Property Get description() As String
    description = pDescription
End Property
Public Property Let description(rhs As String)
    pDescription = rhs
End Property
Public Property Get effort() As Rating
    effort = pEffort
End Property
Public Property Get priority() As Rating
    priority = pPriority
End Property
Public Property Get issueTitle() As String
    issueTitle = issueTitleText
End Property
Public Property Let number(ByVal rhs As Integer)
    pNumberText = Format(rhs, "00")
End Property
Public Property Get numberText() As String
    numberText = pNumberText
End Property
Public Property Let effortText(ByRef rhs As String)
    Select Case rhs
        Case "Long"
            pEffort = Rating.High
        Case "Medium"
            pEffort = Rating.Medium
        Case "Short"
            pEffort = Rating.Low
        Case Else
            pEffort = Rating.Undef
    End Select
End Property
Public Property Let categoryText(ByVal rhs As String)
    pCategoryText = rhs
End Property
Public Property Get categoryText() As String
    categoryText = pCategoryText
End Property
Public Property Get effortText() As String
    Select Case pEffort
        Case Rating.High
            effortText = "Long"
        Case Rating.Medium
            effortText = "Medium"
        Case Rating.Low
            effortText = "Short"
        Case Else
            effortText = "Undefined"
    End Select
End Property
Public Property Let priorityText(ByRef rhs As String)
    Select Case rhs
        Case "High"
            pPriority = Rating.High
        Case "Medium"
            pPriority = Rating.Medium
        Case "Low"
            pPriority = Rating.Low
        Case Else
            pPriority = Rating.Undef
    End Select
End Property
Public Property Get priorityText() As String
    Select Case pPriority
        Case Rating.High
            priorityText = "High"
        Case Rating.Medium
            priorityText = "Medium"
        Case Rating.Low
            priorityText = "Low"
        Case Else
            priorityText = "Undefined"
    End Select
End Property
Public Function pretty() As String
    pretty = "Priority/Severity: " & vbTab & Me.priorityText & vbNewLine & _
        "Effort Est.: " & vbTab & vbTab & Me.effortText & vbNewLine & _
        "Issue Title: " & vbTab & vbTab & issueTitleText & vbNewLine & _
        "SOM Category: " & vbTab & vbTab & Me.categoryText
End Function
Public Sub priorityParse(ByVal priorityLine As String) ' TODO: don't need this sub, replaced by full match with submatches in FindingText.testFindingTitle
    Static re As Object

    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.Multiline = True
        re.Pattern = "(High|Medium|Low)"
    End If
    
    Set matches = re.Execute(priorityLine)
    If matches.Count = 1 Then
        Me.priorityText = matches
    Else
        MsgBox "Encountered multiple matches of RegExp /(High|Medium|Low)/ in priority line, expecting only one match."
    End If
End Sub
Public Sub numberParse(ByVal titleLine As String)
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBSCript.RegExp")
        re.Global = True
        re.Multiline = True
        re.Pattern = "^Issue #" ' TODO
    End If
End Sub
Function NewlineRE()
    NewlineRE = "(\r|\v|\n|\r\n)"
End Function
Public Sub parseTitleForFinding(ByRef pptShape As shape, numberOfFindings As Integer)
    Static re As Object
    Static reDigit As Object
    Set tr = pptShape.TextFrame.TextRange
    text = tr.text
    If re Is Nothing Then
        testFindingTitlePattern = "^Severity/Priority: (High|Medium|Low)" & NewlineRE() & _
            "Issue \#(\d\d)...([^" & NewlineRE() & "]*)" & NewlineRE() & _
            "Effort estimation: (Long|Medium|Short)" & NewlineRE() & _
            "Category: ([^" & NewlineRE() & "]*)$"
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.IgnoreCase = False
        re.Multiline = True
        re.Pattern = testFindingTitlePattern
    End If
    
    If reDigit Is Nothing Then
        testFindingTitlePattern = "^Severity/Priority: (High|Medium|Low)" & NewlineRE() & _
            "Issue \#"
        Set reDigit = CreateObject("VBScript.RegExp")
        reDigit.Global = True
        reDigit.IgnoreCase = False
        reDigit.Multiline = True
        reDigit.Pattern = testFindingTitlePattern
    End If
    
    
    If re.Test(text) Then
        testFindingTitle = True
    End If
    Set matches = re.Execute(text)
    If Not matches Is Nothing And matches.Count = 1 Then
        With matches(0)
            Me.priorityText = .SubMatches(0)
            Me.number = numberOfFindings + 1
            issueTitleText = .SubMatches(3)
            Me.effortText = .SubMatches(5)
            Me.categoryText = .SubMatches(7)
            Me.titleReady = True
            
            Set digitMatches = reDigit.Execute(text)
            tr.Characters(digitMatches(0).Length + 1, 2) = Me.numberText
            
            'tr.Characters(.SubMatches(2).Start, .Sub
        End With
    End If
End Sub

'Public Sub parseBodyForFinding(text As String)
'    Static reDescription As Object
'    Static reBusinessImpact As Object
'    Static reRecommendedActions As Object
'
'    If re Is Nothing Then
'        testDescriptionPattern = "^\s*Description:\s*" & NewlineRE()
'
'        Set reDescription = CreateObject("VBScript.RegExp")
'        reDescription.Global = True
'        reDescription.IgnoreCase = False
'        reDescription.Multiline = True
'        reDescription.Pattern = testBodyPattern
'
'        testBusinessImpactPattern = NewlineRE() & "\s*Business impact:\s*" & "NewLineRE()"
'        Set reBusinessImpact = CreateObject("VBScript.RegExp")
'        reBusinessImpact.Global = True
'        reBusinessImpact.IgnoreCase = False
'        reBusinessImpact.Multiline = True
'        reBusinessImpact.Pattern = testBodyPattern
'
'        testRecommendedActionsPatters = NewlineRE() & "\s*Recommended actions:\s*" & "NewLineRE()"
'        Set reRecommendedActions = CreateObject("VBScript.RegExp")
'        reRecommendedActions.Global = True
'        reRecommendedActions.IgnoreCase = False
'        reRecommendedActions.Multiline = True
'        reRecommendedActions.Pattern = testBodyPattern
'    End If
'
'End Sub

Public Sub parseBodyForFinding(text As String)
    Static reDescription As Object
    Static reBusinessImpact As Object
    Static reRecommendedActions As Object

    ' Initialize the regex objects if they haven't been initialized yet
    If reDescription Is Nothing Then
        ' Regex for Description
        Set reDescription = CreateObject("VBScript.RegExp")
        reDescription.Global = False
        reDescription.IgnoreCase = False
        reDescription.Multiline = True
        reDescription.Pattern = "^Description:\r"
        
        ' Regex for Business Impact
        Set reBusinessImpact = CreateObject("VBScript.RegExp")
        reBusinessImpact.Global = False
        reBusinessImpact.IgnoreCase = False
        reBusinessImpact.Multiline = True
        reBusinessImpact.Pattern = "^Business impact:\r"
        
        ' Regex for Recommended Actions
        Set reRecommendedActions = CreateObject("VBScript.RegExp")
        reRecommendedActions.Global = False
        reRecommendedActions.IgnoreCase = False
        reRecommendedActions.Multiline = True
        reRecommendedActions.Pattern = "^Recommended actions:\r"
    End If

    ' Check if all three regexes match the input text in the correct order
    Dim descMatch As Object
    Dim impactMatch As Object
    Dim actionsMatch As Object
    
    Set descMatch = reDescription.Execute(text)
    If Not descMatch.Count = 1 Then
        Exit Sub
    End If
    Set impactMatch = reBusinessImpact.Execute(text)
    If Not impactMatch.Count = 1 Then
        Exit Sub
    End If
    Set actionsMatch = reRecommendedActions.Execute(text)
    If Not actionsMatch.Count = 1 Then
        Exit Sub
    End If
    
    ' Ensure all three matches are found and in the correct order
    
    If descMatch.Count = 1 And impactMatch.Count = 1 And actionsMatch.Count = 1 Then
        Dim dm As Object
        Dim im As Object
        Dim am As Object
        Set dm = descMatch(0)
        Set im = impactMatch(0)
        Set am = actionsMatch(0)
        If dm.FirstIndex + dm.Length < im.FirstIndex And im.FirstIndex + im.Length < am.FirstIndex Then
            ' Extract the matched groups
            Dim description As String
            Dim businessImpact As String
            Dim recommendedActions As String
            
            Dim descStart As Integer
            Dim descLength As Integer

            descStart = dm.FirstIndex + dm.Length + 1
            descLength = im.FirstIndex - descStart
            description = Mid(text, descStart, descLength)
            
            descriptionLines = split(description, vbCr)
            Me.description = descriptionLines(0)
            Me.bodyReady = True
'            Dim impStart As Integer
'            Dim impLength As Integer
'            Dim actStart As Integer
'            Dim actLength As Integer
'            impStart = im.FirstIndex + im.Length + 1
'            impLength = am.FirstIndex - impStart
'            businessImpact = Mid(text, impStart, impLength)
'
'            actStart = am.FirstIndex + am.Length + 1
'            actLength = Len(text) - actStart
'            recommendedActions = Mid(text, actStart, actLength)
            
            'businessImpact = impactMatch(0).SubMatches(0)
            'recommendedActions = actionsMatch(0).SubMatches(0)
            
            ' Output the results
            'Debug.Print "Description: " & description
            'Debug.Print "Business impact: " & businessImpact
            'Debug.Print "Recommended actions: " & recommendedActions
        End If
    End If
End Sub


