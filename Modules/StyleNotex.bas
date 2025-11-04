Attribute VB_Name = "StyleNotex"
' This module contains the macros
' - Run CreateNotexStyle() once to create and save the "note" style.
' - Then run AssignNotex() to format all "notex" markers in the active document.

Sub CreateNotexStyle()
    ' Creates a custom paragraph style named "note" in the
    ' active Word template. The style uses a small, superscript
    ' Times New Roman font and is stored as a Quick Style.

    Set myTemplate = ActiveDocument.AttachedTemplate.OpenAsDocument
    Debug.Print myTemplate

    On Error Resume Next
        myTemplate.Styles.Add "note"

    ' Define font properties
    With myTemplate.Styles("note").Font
        .Name = "Times New Roman"
        .Size = 9
        .Color = wdColorBlack
        .Engrave = False
        .Superscript = True
    End With

    ' Define style behavior
    With myTemplate.Styles("note")
        .AutomaticallyUpdate = False
        .BaseStyle = "Standard"
        .NextParagraphStyle = "note"
    End With

   ' Save as quickstyle
    myTemplate.Styles("note").QuickStyle = True
        
    ' Save changes
    myTemplate.Close SaveChanges:=wdSaveChanges

    ' Update styles
    ActiveDocument.UpdateStyles
End Sub

Sub AssignNotex()
    ' Searches through the document and applies the "note"
    ' style to every word that exactly matches "notex" (case-insensitive).

    Dim rng As Range

    For Each rng In ActiveDocument.Words

        rng.Select
            If "notex" = LCase(Trim(rng.Text)) Then
                Selection.Style = ActiveDocument.Styles("note")
            End If

    Next rng

End Sub

