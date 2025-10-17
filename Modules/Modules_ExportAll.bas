Attribute VB_Name = "Modules_ExportAll"

Sub ExportAll()
    ' Exports all VBA modules from Normal.dotm into separate .bas files.
    ' Please change FolderName before running this macro!
    '
    ' If you later want to import them again, simply drag and drop the files 
    ' into the Modules section of the VBA editor.

    Dim x As Variant

    ' Set folder name
    FolderName = "C:\Users\patri\Documents\word_makros\backup_" & Format(Now, "yyyymmdd_hhnnss")

    'Create directory if it does not exist
    If Dir$(FolderName, vbDirectory) = "" Then
        MkDir FolderName
    End If

    ' Export each module
    For Each x In NormalTemplate.VBProject.VBComponents

        Debug.Print x.Name
        If Not x.Name = "ThisDocument" Then
            Debug.Print FolderName & "\" & x.Name & ".bas"
            x.Export FolderName & "\" & x.Name & ".bas"
        End If
            
    Next

End Sub
