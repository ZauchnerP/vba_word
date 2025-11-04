Attribute VB_Name = "RemoveExtraParagraphMarks"

Sub RemoveExtraParagraphMarks()
'   Removes double paragraph marks (^p^p) from the current
'   selection and replaces them with single paragraph marks (^p).
'   This helps clean up text that contains unnecessary blank lines.

    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    
End Sub