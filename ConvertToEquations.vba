Option Explicit

Public Sub ConvertDollarLatexToEquations()
    Dim count As Long

    Application.ScreenUpdating = False
    On Error GoTo Fail

    count = count + ConvertDelimitedEquations("$$", wdOMathDisplay)
    count = count + ConvertDelimitedEquations("$", wdOMathInline)

    Application.ScreenUpdating = True
    MsgBox count & " equation(s) converted.", vbInformation
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    MsgBox "Conversion stopped: " & Err.Description, vbExclamation
End Sub

Private Function ConvertDelimitedEquations(ByVal delimiter As String, _
                                           ByVal equationType As WdOMathType) As Long
    Dim doc As Document
    Dim scan As Range
    Dim closing As Range
    Dim whole As Range
    Dim eqRange As Range
    Dim equationText As String
    Dim startPos As Long
    Dim converted As Long

    Set doc = ActiveDocument
    Set scan = doc.Content

    Do While FindNext(scan, delimiter)
        startPos = scan.Start

        Set closing = doc.Range(scan.End, doc.Content.End)
        If Not FindNext(closing, delimiter) Then Exit Do

        equationText = doc.Range(scan.End, closing.Start).Text
        equationText = NormalizeEquationText(equationText)

        Set whole = doc.Range(startPos, closing.End)

        If Len(equationText) > 0 Then
            whole.Text = equationText

            Set eqRange = doc.Range(startPos, startPos + Len(equationText))
            Set eqRange = eqRange.OMaths.Add(eqRange)

            With eqRange.OMaths(1)
                .BuildUp
                .Type = equationType
            End With

            converted = converted + 1
            Set scan = doc.Range(eqRange.End, doc.Content.End)
        Else
            Set scan = doc.Range(closing.End, doc.Content.End)
        End If
    Loop

    ConvertDelimitedEquations = converted
End Function

Private Function FindNext(ByVal target As Range, ByVal textToFind As String) As Boolean
    With target.Find
        .ClearFormatting
        .Text = textToFind
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
    End With

    FindNext = target.Find.Execute
End Function

Private Function NormalizeEquationText(ByVal text As String) As String
    text = Replace(text, vbCr, " ")
    text = Replace(text, vbLf, " ")
    text = Replace(text, vbTab, " ")
    NormalizeEquationText = Trim$(text)
End Function
