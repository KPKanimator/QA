# Paste clipboard into Excel without formatting

## What does this script do?
It hijacks Ctrl+q keyboard shortcut and when this shortcut is pressed in Excel the script pastes content of clipboard without formatting;

## Why even bother?
Excel does not have option to change default Ctrl+v behaviour and will always paste with formatting if the source data has it (i.e. when copying from Word or websites).

Ways to remove pasted formatting in Excel are cumbersome and cost precious microseconds (you either need to use mouse to change pasted content to plain format or to use few more keyboard shortcuts to paste without formatting).

This becomes annoying when you want quickly or often paste without formatting and this script solves it.

## How to use this script?
You need install AutoHotkey v1 (Windows only). https://www.autohotkey.com/. A software to hijack keystrokes and assign them different functions. Then download the **Excel-paste-without-formatting.ahk** file above, save and then double click it. You can add this file to Windows autostart if needed.
When this script is running in background *(you can verify it by checking in System Tray if you see a green icon with H in it)* then pressing Ctrl+q in excel will paste clipboard content without formatting.

## Google got you here by its mysterious ways
If you're here by accident but you don't want use any scripts just want paste without formatting.

### Option 1: Use mouse or Ctrl
After pasting content from clipboard the small clipboard icon will appear next to the cell with pasted content. Click on it or press **Ctrl**. It will expand and you will be able to select paste method to your liking. Pressing **v** when this selection pop-up is open will change the pasting to one without formatting.

### Option 2: Ctrl + Alt + v
Copy content to clipboard and then in Excel press **Ctrl + Alt + v**. This will open Paste Special window where you can select paste method. When in there instead of mouse you can press **Alt + v** *(in some setups just **v** is enough)* and then **Enter**.

## Need use different shortcut
If **Ctrl+q** is not your thing then in line 7 you can change it to any other shortcut. 
Common special keys to use in there are:

`+` Shift

`!` Alt

`^` Ctrl

So to use **Ctrl + Alt + w** you would need change that line from 

`^q::`

to:

`^!w::`