Attribute VB_Name = "RemoveParagraphMarks"

Sub ShortcutsHeadings()
    ' This VBA macro creates shortcuts for applying different
    ' heading styles in MS Word and saves is in the "Normal Template."
    ' It assigns Alt + 1 to Alt + 9 to the corresponding heading
    ' styles (Heading 1 to Heading 9).

    CustomizationContext = NormalTemplate

    ' Heading 1
    KeyBindings.Add _
        KeyCode:=BuildKeyCode(wdKey1, wdKeyAlt), _
        KeyCategory:=wdKeyCategoryStyle, _
        Command:=ActiveDocument.Styles(wdStyleHeading1).NameLocal

    Debug.Print "Test Heading 1: " & KeyBindings(1).Command

    ' Heading 2
    KeyBindings.Add _
        KeyCode:=BuildKeyCode(wdKey2, wdKeyAlt), _
        KeyCategory:=wdKeyCategoryStyle, _
        Command:=ActiveDocument.Styles(wdStyleHeading2).NameLocal

    Debug.Print "Test Heading 2: " & KeyBindings.Key(KeyCode:=BuildKeyCode(wdKey2, wdKeyAlt)).Command

    ' Heading 3
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKey3, wdKeyAlt), KeyCategory:=wdKeyCategoryStyle, _
    Command:=ActiveDocument.Styles(wdStyleHeading3).NameLocal
    Debug.Print "Test Heading 3: " & KeyBindings.Key(KeyCode:=BuildKeyCode(wdKey3, wdKeyAlt)).Command

    ' Heading 4
    KeyBindings.Add _
        KeyCode:=BuildKeyCode(wdKey4, wdKeyAlt), _
        KeyCategory:=wdKeyCategoryStyle, _
        Command:=ActiveDocument.Styles(wdStyleHeading4).NameLocal
    Debug.Print "Test Heading 4: " & KeyBindings.Key(KeyCode:=BuildKeyCode(wdKey4, wdKeyAlt)).Command

    ' Heading 5
    KeyBindings.Add _
        KeyCode:=BuildKeyCode(wdKey5, wdKeyAlt), _
        KeyCategory:=wdKeyCategoryStyle, _
        Command:=ActiveDocument.Styles(wdStyleHeading5).NameLocal
    Debug.Print "Test Heading 5: " & KeyBindings.Key(KeyCode:=BuildKeyCode(wdKey5, wdKeyAlt)).Command

    ' Heading 6
    KeyBindings.Add _
        KeyCode:=BuildKeyCode(wdKey6, wdKeyAlt), _
        KeyCategory:=wdKeyCategoryStyle, _
        Command:=ActiveDocument.Styles(wdStyleHeading6).NameLocal
    Debug.Print "Test Heading 6: " & KeyBindings.Key(KeyCode:=BuildKeyCode(wdKey6, wdKeyAlt)).Command

    ' Heading 7
    KeyBindings.Add _
        KeyCode:=BuildKeyCode(wdKey7, wdKeyAlt), _
        KeyCategory:=wdKeyCategoryStyle, _
        Command:=ActiveDocument.Styles(wdStyleHeading7).NameLocal
    Debug.Print "Test Heading 7: " & KeyBindings.Key(KeyCode:=BuildKeyCode(wdKey7, wdKeyAlt)).Command

    ' Heading 8
    KeyBindings.Add _
        KeyCode:=BuildKeyCode(wdKey8, wdKeyAlt), _
        KeyCategory:=wdKeyCategoryStyle, _
        Command:=ActiveDocument.Styles(wdStyleHeading8).NameLocal
    Debug.Print "Test Heading 8: " & KeyBindings.Key(KeyCode:=BuildKeyCode(wdKey8, wdKeyAlt)).Command

    ' Heading 9
    KeyBindings.Add _
        KeyCode:=BuildKeyCode(wdKey9, wdKeyAlt), _
        KeyCategory:=wdKeyCategoryStyle, _
        Command:=ActiveDocument.Styles(wdStyleHeading9).NameLocal
    Debug.Print "Test Heading 9: " & KeyBindings.Key(KeyCode:=BuildKeyCode(wdKey9, wdKeyAlt)).Command

End Sub

