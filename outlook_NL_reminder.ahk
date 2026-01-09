﻿#Requires AutoHotkey v2.0
#SingleInstance Force

; ─────────────── USER‑FACING SETTINGS ───────────────
hotkey   := "^!r"       ; ⇧Ctrl+Alt+R
defaultTime := "09:00 AM"
dialogSize  := "w450 h200"
; ────────────────────────────────────────────────────

Hotkey(hotkey, AddReminder)

AddReminder(*) {
    ; 1) Ask the user for a natural‑language reminder
    prompt :=
        "Type your reminder (examples):`n" .
        '  - "remind mike to give an answer in two days"`n' .
        '  - "remind me next Tuesday at 10 am to prepare and file a reply"'
    ib := InputBox(prompt, "Add Reminder", dialogSize)
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
