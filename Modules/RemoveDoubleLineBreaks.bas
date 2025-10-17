Attribute VB_Name = "RemoveDoubleLineBreaks"

Sub RemoveDoubleLineBreaks()
    '   Removes double paragraph marks (^p^p) in the
    '   current selection, replacing them with single
    '   paragraph marks (^p). Useful for cleaning up
    '   text with extra blank lines.

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
    Selection.Find.Execute

End Sub
