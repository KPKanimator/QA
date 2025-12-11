#SingleInstance Force

; Press Ctrl+q to paste clipboard into Excel without formatting

; Run this hotkey hijack only if Excel is active
#If WinActive("ahk_exe EXCEL.EXE") ; Check if active window process is run from excel.exe
^q:: ; Ctrl + q - this is the shortcur we are hijacking
    Clipboard := Clipboard ; This line removes any present formatting in clipboard and changes it to plain text
    Send ^v ; This line sends Ctrl+v to the active window
return
#If  ; End condition

