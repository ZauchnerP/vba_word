Attribute VB_Name = "ReplaceParagraphMarks"

' See shortcut assignment below!

Sub ReplaceParagraphMarks()
    ' This VBA macro removes paragraph marks (^13) from the selected text in
    ' MS Word and replaces them with spaces.
    ' It is useful when pasted text contains unwanted hard returns
    ' that mark the end of lines but not actual paragraph breaks in the original source.
    '
    ' The function then sets the selected text to the "Normal" style.
    ' Caution! If some paragraphs contain parts with a different style
    ' than the rest of the paragraph, those styles won't be changed to the "Normal Style".

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

    With Selection.Find
        .Text = "^13"  '^p or ^13
        .Replacement.Text = " "
        .Forward = True
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' Apply "Normal" style ("Standard" in German) to the new paragraph
    With Selection
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        Selection.Expand Unit:=wdParagraph
        Selection.Style = wdStyleNormal
    End With

End Sub

Sub Shortcut_ReplaceParagraphMarks()
    ' Create a shortcut for the function ReplaceParagraphMarks().

    CustomizationContext = NormalTemplate
    KeyBindings.Add _
        KeyCode:=BuildKeyCode(wdKeyW, wdKeyAlt), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="ReplaceParagraphMarks"

    ' Test shortcut assignment
    Debug.Print "Shortcut created in NormalTemplate: Alt+W runs " & KeyBindings.Key(KeyCode:=BuildKeyCode(wdKeyW, wdKeyAlt)).Command

End Sub