Attribute VB_Name = "SplitSentences"

' See shortcut assignment below!

Sub SplitSentences()

    ' Ensure the entire paragraph is selected
    Dim rng As Range
    Set rng = Selection.Paragraphs(1).Range


    ' Apply Normal style to that paragraph
    rng.Style = ActiveDocument.Styles(wdStyleNormal)
    
    ' Protect some patterns temporarily
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Wrap = wdFindStop
        .Forward = True
        .MatchCase = TRUE

        ' Protect "et al."
        .Text = "et al."
        .Replacement.Text = "et_al_TEMP"
        .Execute Replace:=wdReplaceAll

        ' Protect "(p."
        .Text = "(p."
        .Replacement.Text = "(p_TEMP"
        .Execute Replace:=wdReplaceAll
        
        ' Protect " p."
        .Text = " p."
        .Replacement.Text = " p__TEMP"
        .Execute Replace:=wdReplaceAll
        
        ' Protect "z.B."
        .Text = "z.B."
        .Replacement.Text = "zB_TEMP"
        .Execute Replace:=wdReplaceAll
        
        ' Protect "z. B."
        .Text = "z. B."
        .Replacement.Text = "z_B_TEMP"
        .Execute Replace:=wdReplaceAll

        ' Protect "vs."
        .Text = " vs."
        .Replacement.Text = " v_s_TEMP"
        .Execute Replace:=wdReplaceAll

        ' Protect single-letter initials like "A."
        .MatchWildcards = True
        .Text = "([A-Z])\. "
        .Replacement.Text = "\1_TEMP "
        .Execute Replace:=wdReplaceAll
        .MatchWildcards = False
    End With

    ' Define arrays for find and replace operations
    Dim FindArray As Variant
    Dim ReplaceArray As Variant
    FindArray = Array(". ", "? ", "! ")
    ReplaceArray = Array(".^p", "?^p", "!^p")

    ' Loop over all punctuation marks and replace them
    Dim i As Integer
    For i = 0 To UBound(FindArray)
        With rng.Find
            .Text = FindArray(i)
            .Replacement.Text = ReplaceArray(i)
            .Forward = True
            .Wrap = wdFindStop
            .Execute Replace:=wdReplaceAll
        End With
    Next i
    
    ' Restore the protected citation markers
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Wrap = wdFindStop
        .Forward = True

        .Text = "et_al_TEMP"
        .Replacement.Text = "et al."
        .Execute Replace:=wdReplaceAll

        .Text = "(p_TEMP"
        .Replacement.Text = "(p."
        .Execute Replace:=wdReplaceAll
        
        ' " p."
        .Text = " p__TEMP"
        .Replacement.Text = " p."
        .Execute Replace:=wdReplaceAll
        
        ' "z.B."
        .Text = "zB_TEMP"
        .Replacement.Text = "z.B."        
        .MatchCase = False
        .Execute Replace:=wdReplaceAll
        
        ' "z. B."
        .Text = "z_B_TEMP"
        .Replacement.Text = "z. B."        
        .MatchCase = False
        .Execute Replace:=wdReplaceAll

        ' "vs."
        .Text = " v_s_TEMP"
        .Replacement.Text = " vs."        
        .MatchCase = False
        .Execute Replace:=wdReplaceAll

        ' Single-letter initials like "A."
        .MatchWildcards = True
        .Text = "([A-Z])_TEMP "
        .Replacement.Text = "\1. "
        .Execute Replace:=wdReplaceAll
        .MatchWildcards = False

    End With

End Sub


Sub Shortcut_SplitSentence()
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyQ, wdKeyAlt), _
        KeyCategory:=wdKeyCategoryMacro, Command:="SplitSentences"

    ' Test shortcut assignment
    Debug.Print "Shortcut created in NormalTemplate: Alt+Q runs " & KeyBindings.Key(KeyCode:=BuildKeyCode(wdKeyQ, wdKeyAlt)).Command
End Sub
