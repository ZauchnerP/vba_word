Attribute VB_Name = "InsertWithoutBreak"

' See shortcut assignment below!

Sub InsertWithoutBreak()
    '   Paste the current clipboard content into the document,
    '   then remove all paragraph breaks (^13) from the pasted
    '   text and replace them with spaces. The cleanup is applied
    '   only to the freshly pasted block, not the entire document.
    
    Dim startPos As Long
    Dim pastedRange As Range
    
    ' Remember position before pasting
    startPos = Selection.Start
    
    ' Paste
    Selection.Paste
    
    ' Define range covering only the pasted content
    Set pastedRange = ActiveDocument.Range(startPos, Selection.End)
    
    ' Replace paragraph marks with spaces
    With pastedRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^13"
        .Replacement.Text = " "
        .Forward = True
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    
End Sub

Sub Shortcut_InsertWithoutBreak()
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyY, wdKeyAlt), _
        KeyCategory:=wdKeyCategoryMacro, Command:="InsertWithoutBreak"

    ' Test shortcut assignment
    Debug.Print "Shortcut created in NormalTemplate: Alt+Y runs " & KeyBindings.Key(KeyCode:=BuildKeyCode(wdKeyY, wdKeyAlt)).Command
End Sub