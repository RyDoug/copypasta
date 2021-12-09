#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

/*
This script build a popup menu which allows you to select specific files that are in the root directory of this script location.
The menu Builder will search recursivley for txt, doc, and docx files located in the root dir of the script.
Submenu's will be built for each directory, for organization of files.
When you select a menu item, the contents of the specified text file will be copied then pasted into your selected text box.
*/

; Specifies the accepted file extensions.
Extensions := "txt,docx,doc"
; Name of the toplevel menu
TopLevel := "topLevel"

; Menu Builder
Loop, Files, *, R
{
    ; Checks if the current file is an accepted type, if not we go to the next iteration.
    if A_LoopFileExt not in %Extensions% ;or A_LoopFileName = ""
        continue

    ; Adds a top level menu for files not in a subfolder
    if (A_LoopFileDir = "")
    {
        Menu, %TopLevel%, Add, %A_LoopFileName%, FileHandler
    }

    ; Adds sub level menu for files in a subfolder
    else
    {
        ; The \ after LoopFileDir is needed so the pathing is correct when we rebuild the full path in the file handler.
        Menu, %A_LoopFileDir%\, Add, %A_LoopFileName%, FileHandler
        Menu, %TopLevel%, Add, %A_LoopFileDir%, :%A_LoopFileDir%\
    }
}
return

; This handler is used to call the paste function which copies and pastes the texts.
FileHandler:
PastaFunc(A_ThisMenu, A_ThisMenuItem)
return

; This func determines if you have a rich text or plain text file, then calls the appropriate function.
PastaFunc(pathName, itemName)
{
    ; Backup clipboard
    ClipSave := ClipboardAll    
    ; Empties clipboard for ClipWait
    Clipboard := ""

    ; Since pathName is going to be equal to A_ThisMenu we know the name of the menu and we set it to blank if the item is in the top level.
    if pathName = TopLevel
        pathName := ""
    
    ; Builds the path to the file
    ; pathName is blank if item is in the top level, and will have the trailing \ if it is in a submenu because we add it to the menu name when creating the submenu.
    itemPath := % A_WorkingDir . "\" . pathName . itemName
    
    ; For doc type files we go to the rich text function
    if A_ThisMenuItem contains .doc
        RichText(itemPath)

    ; For txt type files we go to the plain text function
    else if A_ThisMenuItem contains .txt
        PlainText(itemPath)

    ; This sleep helps make sure the paste finished before setting clipboard back to the backup.
    sleep 500
    ; Restore clipboard
    Clipboard := ClipSave
}

; Using the com object we copy the contents of the word document to clipboard.
; This requires msoffice installed.
RichText(path)
{
    ; Opens the doc/docx file and copies the full text with formatting.
    oDoc := ComObjGet(path)
    oDoc.Range.FormattedText.Copy
    
    ; Waiting 2 seconds to verify that we copied the data from the document.  If the document is empty, we move on.
    ClipWait, 2, 1
    
    ; Paste
    Send ^v

    ; This sleep helps make sure the paste completes before closing the com object.
    ; When testing, this sleep is really nessicary or the paste becomes inconsistent.
    sleep 500
    
    ; Close the com object. We do this after the send because doing it before send is makes the send inconsistent.
    oDoc.Close(0)
    return
}

; Copy contents of txt file to clipboard and paste to current textbox.
PlainText(path)
{
    ; Reads the file, and adds it to the clipboard.
    FileRead, Clipboard, %path%

    ; Ensures the clipboard is not empty
    ClipWait, 1, 1

    ; Paste
    Send ^v
    return
}


^+j::Menu, %TopLevel%, Show ; This pulls up the menu, use arrow keys to navigate (or the mouse...), use enter to select the document to copypasta, use esc to close menu.
^#r::Reload ; Manually reloads this script.  The menu build only happens when the script loads.  If you make a change to your copypasta repository, you must reload or close and reopen this script.