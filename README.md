# VBA scripts for MS Word

While working on my PhD, I developed a collection of Microsoft Word macros that helped me write my thesis and manage literature references more efficiently.
In this repository, I’m gradually uploading these macros as I revisit or reuse them.
They’re small but practical tools that made my writing process smoother — perhaps they’ll be useful to you, too.

# Modules and macros in this repository

## ShortcutsHeadings()

This macro assigns the shortcuts Alt + 1 to Alt + 9 to the corresponding headings.

## RemoveDoubleLineBreaks()

This macro removes double paragraph marks in the current selection, replacing them with single paragraph marks. Useful for cleaning up
text with extra blank lines.

## RemoveParagraphMarks()

This macro removes paragraph marks from the selected text and replaces them with spaces. 
The function then sets the selected text to the "Normal" style.

## ExportAllModules()

Exports all VBA modules from Normal.dotm into separate .bas files.
If you later want to import them again, simply drag and drop the files into the Modules section of the VBA editor.

## Style_Notex.bas

I used Notex as a marker word in my text to link comments to specific parts of the document. For example, when referring to another text, I inserted the exact wording into a comment and linked it to a Notex placed directly after the period at the end of my own sentence. Since my text often changes during writing, this approach ensured that edits to the main text did not affect the comments.

At the end of the writing process, I could simply search for and replace all occurrences of Notex, and the document would be clean again.

# Importing the bas-files

1) Enable the Developer tab in Word: (Only necessary if you haven't done this before.)
    File → Options → Customize Ribbon → Check "Developer"

2) Go to the Developer tab → Visual Basic.

3) Import bas files:

    - Right-click your VBA project → Import File… → select Code.bas or

    - Drag and drop the files into the Modules section of the VBA editor.

4) In the modules, locate the desired functions and run them.

# License

These scripts have been gathered and modified from various sources over the years and specific origins are not known anymore. While this Word project predates ChatGPT, recent updates and improvements were made with assistance from ChatGPT and Github Copilot.

This project is licensed under the MIT License. See the full license text here: [MIT License](https://opensource.org/licenses/MIT).