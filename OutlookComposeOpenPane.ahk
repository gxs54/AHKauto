#Requires AutoHotkey v2.0+
#SingleInstance Force

; ── CONFIG ───────────────────────────────────────────────────────────────
hotCombo        := "^+l"                ; Ctrl + Shift + L
outlookExe      := "outlook.exe"
waitSecs        := 10                   ; seconds to wait for a new compose window
keytipDelayMs   := 80                   ; pause between each keytip keystroke
modalTitle      := "Office Add-ins"     ; title of the modal opened by Alt,0,7
modalWaitSecs   := 3                    ; seconds to wait for the modal
modalActSeq     := "{Tab}{Enter}"       ; keystrokes inside the modal to activate LLM Edit
modalActPause   := 600                  ; ms to wait after modalActSeq before checking
logPath         := EnvGet("TEMP") . "\OutlookComposeOpenPane.log"
composeClass    := "rctrl_renwnd32"
composeTitleSub := "- Message ("        ; matches HTML / Plain Text inspectors
; ─────────────────────────────────────────────────────────────────────────

Hotkey(hotCombo, OpenLlmEditPane)
Log("=== AHK script loaded ===")
return

Log(msg) {
    global logPath
    stamp := FormatTime(, "HH:mm:ss") "." Format("{:03d}", A_MSec)
    try FileAppend(stamp " [AHK] " msg "`r`n", logPath, "UTF-8")
}

OpenLlmEditPane(*) {
    global logPath, modalTitle, modalWaitSecs, modalActSeq, modalActPause
    global keytipDelayMs

    Log("------ hotkey fired ------")

    hwnd := ResolveComposeTarget()
    if (!hwnd)
        return

    WinActivate("ahk_id " hwnd)
    if !WinWaitActive("ahk_id " hwnd, , 2) {
        Log("FAIL: could not activate hwnd=" hwnd)
        return ShowError("Could not activate the compose window. Log: " logPath)
    }

    ; Open Office Add-ins modal via QAT keytips
    Send("{Alt down}{Alt up}")
    Sleep keytipDelayMs
    Send("0")
    Sleep keytipDelayMs
    Send("7")

    if !WinWait(modalTitle, , modalWaitSecs) {
        Log("FAIL: '" modalTitle "' modal did not open within " modalWaitSecs "s")
        return ShowError("'" modalTitle "' modal did not open. Log: " logPath)
    }
    WinActivate(modalTitle)
    WinWaitActive(modalTitle, , 1)

    Send(modalActSeq)
    Sleep modalActPause
    if WinExist(modalTitle) {
        Log("FAIL: modal still open after '" modalActSeq "'")
        ShowError("Modal did not close after " modalActSeq ". Log: " logPath)
    } else {
        Log("OK")
    }
}

ResolveComposeTarget() {
    global composeClass, composeTitleSub, outlookExe, waitSecs, logPath

    ; 1. Reuse an already-popped-out compose if it's the active window
    if (WinActive("ahk_class " composeClass)) {
        title := WinGetTitle("A")
        if InStr(title, composeTitleSub) {
            h := WinGetID("A")
            Log("reuse active compose hwnd=" h)
            return h
        }
    }

    ; Snapshot existing compose windows so we can identify the new one
    oldWins := WinGetList("ahk_class " composeClass)

    ; 2. Use Outlook COM to pop out an inline draft or start a Reply All
    action := TriggerViaCom()

    ; 3. Fallback: brand new compose
    if (action == "") {
        Log("no inline/selection; launching new compose")
        Run(outlookExe . ' /c ipm.note')
        action := "new"
    } else {
        Log("action via COM: " action)
    }

    ; Wait for a new compose window to appear
    deadline := A_TickCount + (waitSecs * 1000)
    Loop {
        Sleep 200
        curWins := WinGetList("ahk_class " composeClass)
        for h in curWins {
            if !isMember(h, oldWins) && InStr(WinGetTitle(h), composeTitleSub) {
                Log("new compose hwnd=" h)
                return h
            }
        }
        if (A_TickCount > deadline) {
            Log("FAIL: timeout waiting for compose (action=" action ")")
            ShowError("Timed out waiting for the compose window. Log: " logPath)
            return 0
        }
    }
}

; Returns one of: "popout-inline", "replyall", or "" (nothing actionable).
TriggerViaCom() {
    outlook := 0
    try outlook := ComObjActive("Outlook.Application")
    if (!outlook) {
        Log("COM: Outlook not accessible (not running?)")
        return ""
    }

    explorer := 0
    try explorer := outlook.ActiveExplorer
    if (!explorer)
        return ""

    ; 2a. Inline reply / forward in the reading pane — pop it out
    inline := 0
    try inline := explorer.ActiveInlineResponse
    if (inline) {
        try {
            inline.Display()
            return "popout-inline"
        } catch as e {
            Log("COM: inline.Display failed: " e.Message)
        }
    }

    ; 2b. Selected mail item — Reply All (popped out by default for new items)
    try {
        if (explorer.Selection.Count > 0) {
            item := explorer.Selection.Item(1)
            reply := item.ReplyAll()
            reply.Display()
            return "replyall"
        }
    } catch as e {
        Log("COM: ReplyAll failed: " e.Message)
    }

    return ""
}

ShowError(msg) {
    ToolTip(msg)
    SetTimer(() => ToolTip(), -6000)
}

isMember(val, arr) {
    for v in arr
        if (v = val)
            return true
    return false
}
