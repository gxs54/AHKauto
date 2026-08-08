; MatterJump.ahk — AutoHotkey v2.x  (2025‑07‑02 consolidated & cleaned)

#Requires AutoHotkey v2.0

;───────── USER SETTINGS ─────────
drive          := "Q:"                         ; root share
hotkeyCombo    := "^+j"                        ; Ctrl + Shift + J
enableLogging  := false                        ; true = write MatterJump.log
logFile        := A_ScriptDir "\MatterJump.log"

forceMaximize  := true                         ; true=maximize, false=fixed size
minWidth       := 1400                         ; used only if forceMaximize=false
minHeight      := 900

clientRE       := "^[A-Z]{4}"                  ; first four letters
restRE         := "i)^[PT]\d+|^\d+[PT]"        ; detect P/T anywhere (case‑insensitive)
;─────────────────────────────────


;───────── LOG HELPER (≤100 lines) ─────────
Log(msg) {
    global enableLogging, logFile
    if !enableLogging
        return
    entry := Format("{:s} | {:s}", A_Now, msg)

    arr := []
    if FileExist(logFile) {
        txt := FileRead(logFile)
        if txt != ""
            arr := StrSplit(RTrim(txt,"`n"), "`n", "`r")
    }
    arr.Push(entry)
    while arr.Length > 100
        arr.RemoveAt(1)

    joined := ""
    for _, ln in arr
        joined .= ln "`n"

    if FileExist(logFile)
        FileDelete(logFile)
    if joined != ""
        FileAppend(joined, logFile, "UTF-8")
}

;───────── HOTKEY ─────────
Hotkey(hotkeyCombo, (*) => HandleHotkey())

HandleHotkey() {
    Log("Hot-key pressed")
    sel := GetHighlightedText()
    Log("Clipboard → '" sel "'")

    if sel != "" && TryOpenMatter(sel) {
        Log("Opened via highlight")
        return
    }

    ib := InputBox(
        "Enter full matter number (e.g., WOOS12PUS01)",
        "Open Matter Folder",
        "w300 h120",
        sel
    )
    if ib.Result = "Cancel" {
        Log("Prompt cancelled")
        return
    }
    user := StrUpper(Trim(ib.Value))
    Log("User entered → '" user "'")
    TryOpenMatter(user) ? Log("Opened via prompt")
                        : Log("No folder for '" user "'")
}

;───────── RESOLVE & OPEN ─────────
TryOpenMatter(raw) {
    global drive, clientRE, restRE
    raw := StrUpper(Trim(raw))
    Log("TryOpenMatter('" raw "')")

    ; 1) validate client
    if !RegExMatch(raw, clientRE, &m) {
        Log("Invalid client code")
        return false
    }
    client := m[0]
    rest   := SubStr(raw, 5)

    ; 2) detect P/T
    typeF := ""
    if RegExMatch(rest, restRE, &r)
        typeF := InStr(r[0], "P") ? "P" : "T"

    baseDir := drive "\" SubStr(client,1,1) "\" client

    ; 3) build candidate list
    cand := []

    if typeF {
        stripped := RegExMatch(rest, "^[PT](.*)$", &s) ? s[1] : rest

        cand.Push(baseDir "\" typeF "\" rest)                    ; exact with repeated P/T
        if stripped != rest
            cand.Push(baseDir "\" typeF "\" stripped)            ; exact without repeated P/T

        cand.Push({dir:baseDir "\" typeF, prefix:rest})          ; prefix in P/T
        if stripped != rest
            cand.Push({dir:baseDir "\" typeF, prefix:stripped})
			
		fullName := client rest                      					; STRG90PUS01
		cand.Push(baseDir "\" typeF "\" fullName)                       ; exact
		cand.Push({dir:baseDir "\" typeF, prefix:fullName})             ; prefix	
    }

    cand.Push(baseDir "\" rest)                                  ; exact in root
    cand.Push({dir:baseDir, prefix:rest})                        ; prefix in root
	
	; client+rest in root (CBED2PUS01 style)
	fullRootName := client rest
	cand.Push(baseDir "\" fullRootName)                           ; exact
	cand.Push({dir:baseDir, prefix:fullRootName})                 ; prefix


    ; try candidates
    for , c in cand {
        if IsObject(c) {
            p := FindPrefixDir(c.dir, c.prefix)
            Log("Prefix '" c.prefix "' in '" c.dir "' → " (p?"hit":"none"))
            if p && OpenExplorer(p)
                return true
        } else if DirExist(c) && OpenExplorer(c) {
            return true
        }
    }

    ; 4) MOHN-style two-level (parent ends with 01)
    if typeF && RegExMatch(rest, "^(\d+)" . typeF . "([A-Z]{2})(\d{2})(.*)$", &e1) {
        digits  := e1[1]
        country := e1[2]
        serial  := e1[3]
        suffix  := e1[4]
        parent  := client digits typeF country "01"                 ; e.g. MOHN14PUS01
        child   := digits typeF country serial suffix               ; e.g. 14PUS02CON
        parentDir := baseDir "\" typeF "\" parent
        if DirExist(parentDir) {
            full := parentDir "\" child
            if DirExist(full) && OpenExplorer(full)
                return true
            hit := FindPrefixDir(parentDir, child)
            if hit && OpenExplorer(hit)
                return true
        }
    }

    ; 5) ANCH “same-number” fallback (no repeated P/T in child name)
    if typeF && RegExMatch(rest, "^[PT]0*(\d+)([A-Z]{2})(\d{2})(.*)$", &e2) {
        num        := e2[1]                       ; 101
        country    := e2[2]                       ; US
        serial     := e2[3]                       ; 02
        suffix     := e2[4]                       ; (maybe blank)
        parentPref := num                         ; 101
        childPref  := num country serial suffix   ; 101US02...

        Loop Files baseDir "\" typeF "\" parentPref "*", "D" {
            parentPath := A_LoopFilePath
            hit := FindPrefixDir(parentPath, childPref)
            if hit && OpenExplorer(hit) {
                Log("same-number fallback hit → " hit)
                return true
            }
        }
    }

    ; 6) Deep recursive prefix scan (client folder)
    hit := DeepPrefixScan(baseDir, rest)
    if hit {
        Log("Deep scan hit → " hit)
        return OpenExplorer(hit)
    }

    Log("No matching folder found")
    return false
}

;───────── CLIPBOARD → TEXT ─────────
GetHighlightedText() {
    global A_Clipboard
    saved := ClipboardAll()
    A_Clipboard := ""

    Send "^c"
    Sleep 60
    Send "^c"
    ClipWait(0.8)

    if A_Clipboard = "" {
        SendMessage 0x301, 0, 0,, "A"   ; WM_COPY
        ClipWait(0.8)
        if A_Clipboard = "" {
            Send "^{Insert}"
            ClipWait(0.8)
        }
    }
    txt := Trim(A_Clipboard)
    A_Clipboard := saved
    return txt
}

;───────── OPEN EXPLORER (new hwnd, size/monitor, COM view) ─────────
OpenExplorer(path) {
    global forceMaximize, minWidth, minHeight
    Log("Run Explorer → " path)

    old := GetExplorerHwnds()
    Run("explorer.exe /n," Chr(34) path Chr(34))

    ; find newly-created hwnd
    newHwnd := 0
    start := A_TickCount
    while (A_TickCount - start < 5000) {
        cur := GetExplorerHwnds()
        for _, h in cur {
            isOld := false
            for _, oh in old {
                if (oh = h) {
                    isOld := true
                    break
                }
            }
            if !isOld {
                newHwnd := h
                break
            }
        }
        if newHwnd
            break
        Sleep 50
    }

    if !newHwnd {
        Log("No new Explorer hwnd detected – skipping tweaks")
        return true
    }

    WinActivate newHwnd
    WinRestore  newHwnd

    ; monitor under mouse
    MouseGetPos &mx, &my
    monCount := MonitorGetCount()
    useIdx := 0
    Loop monCount {
        MonitorGetWorkArea(A_Index, &L, &T, &R, &B)
        if (mx >= L && mx < R && my >= T && my < B) {
            useIdx := A_Index
            break
        }
    }
    if (useIdx = 0)
        useIdx := MonitorGetPrimary()
    MonitorGetWorkArea(useIdx, &L, &T, &R, &B)

    if forceMaximize
		WinMaximize newHwnd
	else
		EnsureExplorerSize(newHwnd, minWidth, minHeight)


    ; set view, sort & grouping silently via COM
    ; (retry: the COM doc exists before navigation finishes, and settings
    ;  applied too early are silently dropped — confirm via read-back)
    applied := false
    Loop 25 {
        win := GetShellWindowFromHwnd(newHwnd)
        if win {
            try {
                doc := win.Document
                ; 8 = Content view
                doc.CurrentViewMode := 8
                ; leading "-" = descending (shell canonical sort syntax)
                doc.SortColumns := "prop:-System.DateModified;"
                ; group direction is NOT scriptable via doc.GroupBy (always
                ; ascending) — needs IFolderView2::SetGroupBy with fAscending=0
                ; (read-back is a substring match: Explorer may normalize the
                ;  string, and an exact compare would spin all 25 retries)
                if InStr(doc.SortColumns, "-System.DateModified") && SetGroupByDateModifiedDesc(win) {
                    applied := true
                    break
                }
            }
        }
        Sleep 100
    }
    if !applied
        Log("View/sort/group not confirmed for hwnd " newHwnd)

    Log("Explorer adjusted & positioned (hwnd " newHwnd ")")
    return true
}

;───────── HELPERS ─────────
FindPrefixDir(dir, prefix) {
    dir := RTrim(dir,"\/")
    Loop Files dir "\" prefix "*", "D"
        return A_LoopFilePath
    return ""
}

DeepPrefixScan(rootDir, prefix) {
    rootDir := RTrim(rootDir, "\/")
    Loop Files rootDir "\*", "DR" {
        if InStr(A_LoopFileName, prefix) = 1
            return A_LoopFilePath
    }
    return ""
}

GetExplorerHwnds() {
    return WinGetList("ahk_class CabinetWClass")
}

GetShellWindowFromHwnd(hwnd) {
    for winItem in ComObject("Shell.Application").Windows {
        try if (winItem.HWND = hwnd)
            return winItem            ; IWebBrowser2 (shell window)
    }
    return ""
}

; Group by Date Modified, newest group first, via IFolderView2::SetGroupBy.
SetGroupByDateModifiedDesc(winItem) {
    static SID_STopLevelBrowser := "{4C96BE40-915C-11CF-99D3-00AA004AE837}"
    static IID_IShellBrowser    := "{000214E2-0000-0000-C000-000000000046}"
    static IID_IFolderView2     := "{1AF3A467-214F-4298-908E-06B03E0B39F9}"
    sv := 0
    try {
        sb := ComObjQuery(winItem, SID_STopLevelBrowser, IID_IShellBrowser)
        ComCall(15, sb, "ptr*", &sv)               ; IShellBrowser::QueryActiveShellView
        fv2 := ComObjQuery(sv, IID_IFolderView2)
        ; PROPERTYKEY for System.DateModified = {B725F130-...} pid 14
        pk := Buffer(20)
        DllCall("ole32\CLSIDFromString", "wstr", "{B725F130-47EF-101A-A5F1-02608C9EEBAC}", "ptr", pk)
        NumPut("uint", 14, pk, 16)
        ComCall(17, fv2, "ptr", pk, "int", 0)      ; IFolderView2::SetGroupBy, fAscending=0
        return true
    } catch {
        return false
    } finally {
        if sv
            ObjRelease(sv)
    }
}

EnsureExplorerSize(hwnd, minW, minH) {
    ; First pass
    WinGetPos &x, &y, &w, &h, hwnd
    if (w < minW || h < minH) {
        ; find monitor under mouse
        MouseGetPos &mx, &my
        monCnt := MonitorGetCount()
        use := 0
        Loop monCnt {
            MonitorGetWorkArea(A_Index, &L, &T, &R, &B)
            if (mx >= L && mx < R && my >= T && my < B) {
                use := A_Index
                break
            }
        }
        if (use = 0)
            use := MonitorGetPrimary()
        MonitorGetWorkArea(use, &L, &T, &R, &B)

        w2 := Max(minW,  R - L - 40)
        h2 := Max(minH,  B - T - 80)
        x2 := L + ((R - L - w2) // 2)
        y2 := T + ((B - T - h2) // 2)
        WinMove hwnd, x2, y2, w2, h2
    }

    ; Second pass (Windows sometimes shrinks again right after)
    SetTimer (() => (
        WinGetPos(&xx, &yy, &ww, &hh, hwnd),
        (ww < minW || hh < minH) ? WinMove(hwnd,, , , Max(minW, ww), Max(minH, hh)) : ""
    ), -300)
}
