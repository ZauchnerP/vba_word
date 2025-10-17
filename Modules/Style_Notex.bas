Attribute VB_Name = "Style_Notex"
' This module contains the macros
' - Run CreateNotexStyle() once to create and save the "note" style.
' - Then run Notex1() to format all "notex" markers in the active document.

Sub CreateNotexStyle()
    ' Creates a custom paragraph style named "note" in the
    ' active Word template. The style uses a small, superscript
    ' Times New Roman font and is stored as a Quick Style.

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

   ' Save Quickstyle
    myTemplate.Styles("note").QuickStyle = True
End Sub

Sub Notex1()
    ' Searches through the document and applies the "note"
    ' style to every word that exactly matches "notex" (case-insensitive).

    Dim rng As Range

    For Each rng In ActiveDocument.Words
        Debug.Print (rng)

        rng.Select
            If "notex" = LCase(Trim(rng.Text)) Then
                Selection.Style = ActiveDocument.Styles("note")
            End If

    Next rng


End Sub
