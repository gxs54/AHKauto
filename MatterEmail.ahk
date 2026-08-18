#Requires AutoHotkey v2.0+
#SingleInstance Force

; Allow a second thread ONLY so a press arriving while the prompt is open — or
; during the up-to-waitSecs wait for the new Outlook window — reaches the
; handler and gets told so. At the default of 1 the press is discarded before
; any script code runs and the hotkey goes silently dead. The busy flag below
; still serializes the real work.
#MaxThreadsPerHotkey 2

; ── CONFIG ───────────────────────────────────────────────────────────────
hotCombo   := "^+m"            ; Ctrl + Shift + M triggers the search
outlookExe := "outlook.exe"    ; Adjust if Outlook lives elsewhere
waitSecs   := 10               ; Max seconds to wait for the *new* window
; ─────────────────────────────────────────────────────────────────────────

Hotkey(hotCombo, HandleHotkey)
return

HandleHotkey(*) {
    static busy := false
    if busy {
        TrayTip("Still working on the last search…", "MatterEmail")
        return
    }
    busy := true
    try
        SearchMatterAllMailboxes()
    finally
        busy := false
}

; AutoHotkey's InputBox has no always-on-top option, so the prompt can end up
; behind Outlook — and while it sits there unnoticed the hotkey does nothing at
; all. Raise it as soon as it exists so that state is unreachable.
TopmostInputBox(prompt, title, options := "", default := "") {
    tries := 0
    Raise() {
        if (hwnd := WinExist(title)) {
            try {
                WinSetAlwaysOnTop(true, hwnd)
                WinActivate(hwnd)
            }
            SetTimer(Raise, 0)
        } else if (++tries > 40)
            SetTimer(Raise, 0)
    }
    SetTimer(Raise, 50)
    ib := InputBox(prompt, title, options, default)
    SetTimer(Raise, 0)
    return ib
}

SearchMatterAllMailboxes() {
    global outlookExe, waitSecs

    ; 1 — Capture selected text or prompt ---------------------------------
    saved := A_Clipboard
    A_Clipboard := ""
    Send("^c")
    ClipWait(0.5)
    matter := Trim(A_Clipboard)
    A_Clipboard := saved

    if (matter = "") {
        ib := TopmostInputBox("Enter the matter number:", "Matter Number Needed")
        if ib.Result = "Cancel"
            return
        matter := Trim(ib.Value)
        if (matter = "")
            return MsgBox("A matter number is required.", "Nothing to search", 48)
    }

    ; 2 — Remember current Outlook windows --------------------------------
    oldWins := WinGetList("ahk_class rctrl_renwnd32")

    ; 3 — Launch a NEW Outlook window focused on Inbox ---------------------
    Run(outlookExe . ' /select "outlook:inbox"')
    deadline := A_TickCount + (waitSecs * 1000)
    newHwnd := 0

    Loop {
        Sleep 200
        curWins := WinGetList("ahk_class rctrl_renwnd32")
        for hwnd in curWins {
            if !isMember(hwnd, oldWins) {
                newHwnd := hwnd
                break
            }
        }
        if (newHwnd)
            break
        if (A_TickCount > deadline)
            return MsgBox("Timed out waiting for the new Outlook window.", "Search aborted", 16)
    }

    WinActivate("ahk_id " newHwnd)

    ; 4 — Focus search and set scope --------------------------------------
    Send("^1")        ; Ensure Mail module
    Send("^e")        ; Focus Search box
    Sleep 150         ; Give UI a moment
    Send("^!a")       ; Ctrl + Alt + A → All Mailboxes
    Sleep 100

    ; 5 — Type the query and search ---------------------------------------
    Send("{Text}" . matter)
    Send("{Enter}")
}

; ---- helper: membership check --------------------------------------------
isMember(val, arr) {
    for v in arr
        if (v = val)
            return true
    return false
}
