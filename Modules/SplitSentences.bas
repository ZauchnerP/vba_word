Attribute VB_Name = "SplitSentences"

' See shortcut assignment below!

Sub SplitSentences()

    Dim rng As Range
    Dim i As Long
    Dim FindArray As Variant, ReplaceArray As Variant
    Dim ProtectFind As Variant, ProtectReplace As Variant

    ' Ensure the entire paragraph is selected
    Set rng = Selection.Paragraphs(1).Range

    ' Apply Normal style to that paragraph
    rng.Style = ActiveDocument.Styles(wdStyleNormal)

    ' Define protection patterns (to avoid false splits)
    ProtectFind = Array("et al.", _
                        "(p.", " p.", "pp.", _
                        "i.e.", "e.g.", "o.J.", _
                        "z.B.", "z. B.", " vs.", _
                        "Fig.", "Sect.", _
                        "Chap.", "Ch.", _
                        "a.m.", "p.m.", _
                        "etc." _
                        )
    ProtectReplace = Array("et_al_TEMP", _
                           "(p_TEMP", " p__TEMP", "pp_TEMP", _
                           "ie_TEMP", "eg_TEMP", "oJ_TEMP", _
                           "zB_TEMP", "z_B_TEMP", " v_s_TEMP", _
                           "FIG_TEMP", "sect_TEMP", _
                           "Chap_TEMP", "Ch_TEMP", _
                           "am_TEMP", "pm_TEMP", _
                           "etc_TEMP" _
                           )

    ' Protect fixed patterns
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Wrap = wdFindStop
        .Forward = True
        .MatchCase = True

        For i = LBound(ProtectFind) To UBound(ProtectFind)
            .Text = ProtectFind(i)
            .Replacement.Text = ProtectReplace(i)
            .Execute Replace:=wdReplaceAll
        Next i

        ' Protect single-letter initials like "A."
        .MatchWildcards = True
        .Text = "([A-Z])\. "
        .Replacement.Text = "\1_TEMP "
        .Execute Replace:=wdReplaceAll
        .MatchWildcards = False
    End With

    ' Split sentences after . ? ! 
    FindArray = Array(". ", "? ", "! ")
    ReplaceArray = Array(".^p", "?^p", "!^p")

    For i = LBound(FindArray) To UBound(FindArray)
        With rng.Find
            .Text = FindArray(i)
            .Replacement.Text = ReplaceArray(i)
            .Forward = True
            .Wrap = wdFindStop
            .Execute Replace:=wdReplaceAll
        End With
    Next i

    ' Restore protected patterns
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Wrap = wdFindStop
        .Forward = True
        .MatchCase = False

        For i = LBound(ProtectFind) To UBound(ProtectFind)
            .Text = ProtectReplace(i)
            .Replacement.Text = ProtectFind(i)
            .Execute Replace:=wdReplaceAll
        Next i

        ' Restore single-letter initials
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
