#Requires AutoHotkey v2.0
#SingleInstance Force

; Allow a second thread ONLY so a press that arrives while the prompt is
; already open reaches the handler and gets told so. At the default of 1 the
; press is discarded before any script code runs and the hotkey goes silently
; dead. The busy flag below still serializes the real work.
#MaxThreadsPerHotkey 2

; ─────────────── USER‑FACING SETTINGS ───────────────
hotCombo := "^!r"       ; ⇧Ctrl+Alt+R
defaultTime := "09:00 AM"
dialogSize  := "w450 h200"
; ────────────────────────────────────────────────────

Hotkey(hotCombo, HandleHotkey)

HandleHotkey(*) {
    static busy := false
    if busy {
        TrayTip("The reminder prompt is already open.", "Add Reminder")
        return
    }
    busy := true
    try
        AddReminder()
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

AddReminder() {
    global defaultTime, dialogSize
    ; 1) Ask the user for a natural‑language reminder
    prompt :=
        "Type your reminder (examples):`n" .
        '  - "remind mike to give an answer in two days"`n' .
        '  - "remind me next Tuesday at 10 am to prepare and file a reply"'
    ib := TopmostInputBox(prompt, "Add Reminder", dialogSize)
    if ib.Result != "OK"
        return
    user := ib.Value

    ; 2) Parse the phrase ----------------------------------------------------
    dateOffset := 0
    weekday    := ""
    timeTxt    := ""
    textReminder := user

    ; ‑‑ “in X days”
    if RegExMatch(user, "\bin\s+(\d+)\s+days?\b", &m)
        dateOffset := m[1]

    ; ‑‑ “tomorrow”
    if RegExMatch(user, "\btomorrow\b", &_)
        dateOffset := 1

    ; ‑‑ Monday … Sunday
    if RegExMatch(user, "\b(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b", &m)
        weekday := StrLower(m[1])

    ; ‑‑ “at 10 am / 9:30 pm” etc.
    if RegExMatch(user, "\bat\s+(\d{1,2})(:(\d{1,2}))?\s*(am|pm)\b", &m) {
        hr := m[1], mn := m[3] != "" ? m[3] : "00", ap := m[4]
        timeTxt := Format("{:02}:{:02} {}", hr, mn, ap)
    }

    ; 3) Calculate the target date ------------------------------------------
    target := A_Now
    if weekday != "" {                ; “next Tuesday”
        Loop 7 {
            target := DateAdd(target, 1, "days")
            if StrLower(FormatTime(target, "dddd")) = weekday
                break
        }
    } else if dateOffset
        target := DateAdd(target, dateOffset, "days")

    finalDate := FormatTime(target, "MM/dd/yyyy")
    if timeTxt = ""
        timeTxt := defaultTime

    ; 4) Strip directive words for cleaner flag text -------------------------
    textReminder := RegExReplace(textReminder, "(?i)\b(remind|reminder)( me)?\b")
    textReminder := RegExReplace(textReminder, "(?i)\bnext\b")
    textReminder := RegExReplace(textReminder, "(?i)\bin\s+\d+\s+days?\b")
    textReminder := RegExReplace(textReminder, "(?i)\btomorrow\b")
    textReminder := RegExReplace(textReminder, "(?i)\b(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b")
    textReminder := RegExReplace(textReminder, "(?i)\bat\s+\d{1,2}(:\d{1,2})?\s*(am|pm)\b")
    textReminder := Trim(textReminder, " .,-")

    ; 5) Send to Outlook’s “Flag for Follow‑Up” dialog -----------------------
    Send("^+g")               ; open flag dialog  (Ctrl + Shift + G)
    Sleep(300)
    SendText(textReminder)
    Send("{Tab}")
    SendText(finalDate)
    Send("{Tab}")
    SendText(timeTxt)
    Send("{Enter}")
}
