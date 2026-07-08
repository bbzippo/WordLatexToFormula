In case you pasted a lot of text with Latex into a Word doc and don't want to manually select each equation and click Insert Equation in the ribbon.

**Issue: does not work within tables.**

Save your doc as macro-enabled .docm.
Alt-F11 to open script editor.
Insert - Module.
Paste in the VBA script, close the scripting window.

Alt-F8 to bring up Macros.
Select **ConvertDollarLatexToEquations** and run.
