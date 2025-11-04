# VBA scripts for MS Word

While working on my PhD, I developed a collection of Microsoft Word macros that helped me write my thesis and manage literature references more efficiently. In this repository, I’m gradually uploading these macros as I revisit or reuse them. They’re quite small and might seem a bit unusual at first, but they’re practical tools that made my writing process smoother — and perhaps they’ll be useful to you, too.

# Modules and macros in this repository

## InsertWithoutBreak()

This macro pastes the current clipboard content into the document, then removes all paragraph breaks (`^13`) from the pasted text and replaces them with spaces. The cleanup is applied only to the newly pasted block, leaving the rest of the document unchanged.

Shortcut: `Alt + Y`

## ExportAllModules()

Exports all VBA modules from `Normal.dotm` into separate `.bas` files.
If you later want to import them again, simply drag and drop the files into the `Modules` section of the VBA editor.

## RemoveExtraParagraphMarks()

This macro replaces all double paragraph marks (`^p^p`) in the current selection with single paragraph marks (`^p`). It helps clean up text by removing unnecessary blank lines.

## ReplaceParagraphMarks()

This macro removes paragraph marks from the selected text and replaces them with spaces. 
The function then sets the selected text to the *Normal* style.

Shortcut: `Alt + W`

## Shortcuts()

TODO Description to be added.

## ShortcutsHeadings()

This macro assigns the shortcuts `Alt + 1` to `Alt + 9` to the corresponding headings.

## SplitSentences()

This macro takes the selected paragraph, applies the *Normal* style, and inserts paragraph breaks after sentence-ending punctuation marks (periods, question marks, and exclamation points). Each sentence is moved to a new line, making the text easier to read or process further.

## StyleNotex.bas

I used “Notex” as a marker word in my text to link comments to specific parts of the document. For example, when referring to another text, I inserted the exact wording into a comment and linked it to a “Notex” placed directly after the period at the end of my own sentence. Since my text often changed during writing, this approach ensured that edits to the main text did not affect the comments.

At the end of the writing process, I could simply search for and replace all occurrences of “Notex,” and the document would be clean again.

This `.bas` file includes:

- `CreateNotexStyle()` — creates the “Notex” style  
- `AssignNotex()` — assigns this style to all “Notex” words in your document

# Cloning this repository

```bash
git clone https://github.com/ZauchnerP/vba_word/ 
```

# Importing the bas-files to Word

1) Enable the Developer tab (only necessary once):

   `File → Options → Customize Ribbon → Check "Developer"`

2) Open the VBA editor:

   `Developer → Visual Basic`

3) Import `bas` files:

    - Right-click your VBA project → `Import File…` → select a `.bas` file or

    - Drag and drop the files into the `Modules` section of the VBA editor.

4) Run the desired macros:

    Locate the module and execute the function you want.

# License

These scripts were inspired by and adapted from various sources over the years, though the original references are no longer identifiable. While this Word project predates ChatGPT, more recent updates and refinements were made with assistance from ChatGPT and GitHub Copilot.

This project is licensed under the MIT License. See the full license text here: [MIT License](https://opensource.org/licenses/MIT).