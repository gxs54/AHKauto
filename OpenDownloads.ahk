#Requires AutoHotkey v2.0

; OpenDownloads.ahk — Open Downloads folder
; Hotkey: Shift + Windows key + D

DownloadsPath := "C:\Users\gsteyer\Downloads"

Hotkey("#+d", OpenDownloads)

OpenDownloads(*) {
    if DirExist(DownloadsPath)
        Run(DownloadsPath)
    else
        MsgBox("Downloads folder not found:`n" . DownloadsPath, "OpenDownloads", "Icon!")
}
