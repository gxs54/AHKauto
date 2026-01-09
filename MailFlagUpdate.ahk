#Requires AutoHotkey v2.0
#SingleInstance Force
SetTitleMatchMode 2
; MailFlagUpdate.ahk — Outlook flag updater (AHK v2)
; Hotkey: Windows key + Shift + U

; =========================
; Config
; =========================
HK_COMBO := "#+u"      ; Windows key + Shift + U
DEBUG  := false       ; set true to see ToolTips
LOG_FILE := "MailFlagUpdate.log"

; =========================
; Reentrancy guard
; =========================
running := false

; =========================
; Hotkey bind (v2 function style)
; =========================
Hotkey(HK_COMBO, DoFlagUpdate)  ; v2 function-call syntax

; Show startup message
ToolTip("MailFlagUpdate loaded. Hotkey: " . HK_COMBO, 20, 20)
SetTimer(() => ToolTip(), -3000)  ; Hide tooltip after 3 seconds

; =========================
; Logging function
; =========================
LogMessage(msg) {
    global LOG_FILE
    timestamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    logLine := timestamp . " | " . msg . "`n"
    FileAppend(logLine, LOG_FILE)
    if DEBUG
        ToolTip(msg, 20, 20)
}

; =========================
; Main hotkey handler
; =========================
DoFlagUpdate(*) {
    global running, DEBUG
    LogMessage("Hotkey triggered!")
    
    if running
        return
    running := true
    Critical

    try {
        item := GetSelectedOutlookItem()
        if !item {
            LogMessage("No selectable Outlook item.")
            return
        }

        ; --- Extract baseline date source (flag text, then subject) ---
		flagText := ""
		try {
			flagText := item.FlagRequest
		} catch {
			; ignore – property not present on some item types
		}

		baseStr := flagText
		if !baseStr {
			try {
				baseStr := item.Subject
			} catch {
				; ignore
			}
		}

        LogMessage("Parsing date from: '" . baseStr . "'")
        baseDeadline := ParseFirstDate(baseStr)  ; "" if not found
        if (baseDeadline = "") {
            LogMessage("No date found in Flag text/Subject. Only tertiary rule can apply.")
        } else {
            LogMessage("Baseline deadline: " . Human(baseDeadline))
        }

        ; --- Read current due date ---
        curDue := ReadOutlookDate(item, "FlagDueBy")  ; "" if none
        if (curDue != "")
            LogMessage("Current due: " . Human(curDue))

        now := A_Now

        ; --- Parse flag suffix to determine month progression ---
        flagSuffix := ""
        if (baseStr != "") {
            if RegExMatch(baseStr, "i)\b2\s*mo\s*rr\b", &m) {
                flagSuffix := "2mo_rr"
            } else if RegExMatch(baseStr, "i)\b3\s*mo\b", &m) {
                flagSuffix := "3mo"
            }
        }
        LogMessage("Detected flag suffix: " . flagSuffix)

        ; --- Secondary checkpoints based on flag suffix ---
        newDue := ""
        if (baseDeadline != "" && flagSuffix != "") {
            if (flagSuffix = "2mo_rr") {
                ; For 2 mo rr: move to next month, then -14 days
                LogMessage("Before DateAdd: baseDeadline = " . baseDeadline)
                
                ; Extract year, month, day from baseDeadline (YYYYMMDDHHMMSS format)
                y := SubStr(baseDeadline, 1, 4)
                m := SubStr(baseDeadline, 5, 2)
                d := SubStr(baseDeadline, 7, 2)
                LogMessage("Parsed: year=" . y . ", month=" . m . ", day=" . d)
                
                ; Add 1 month
                m := Integer(m) + 1
                if (m > 12) {
                    m := 1
                    y := Integer(y) + 1
                }
                nextMonth := Format("{:04}{:02}{:02}090000", y, m, d)
                LogMessage("After manual calculation: nextMonth = " . nextMonth)
                
                d1 := AdjustToMonday(DateAdd(nextMonth, -14, "Days"))
                LogMessage("2mo_rr logic: " . Human(baseDeadline) . " → " . Human(nextMonth) . " → " . Human(d1))
                
                ; For 2mo_rr: apply if within 14 days of the ORIGINAL baseline deadline
                if IsWithinDays(now, baseDeadline, 14) {
                    newDue := d1
                    LogMessage("Within 14 days of ORIGINAL deadline. Applying 2mo_rr progression: " . Human(newDue))
                } else {
                    LogMessage("NOT within 14 days of ORIGINAL deadline " . Human(baseDeadline) . " (today: " . Human(now) . ")")
                }
            } else if (flagSuffix = "3mo") {
                ; For 3 mo: move to +2mo, then -14 days
                d1 := AdjustToMonday(DateAdd(DateAdd(baseDeadline, 2, "Months"), -14, "Days"))
                LogMessage("3mo logic: " . Human(baseDeadline) . " → " . Human(d1))
                
                if IsWithinDays(now, d1, 14) {
                    newDue := d1
                    LogMessage("Within 14 days of (deadline +2mo −14d for 3mo). Setting Due to: " . Human(newDue))
                }
            }
        }

        ; --- Tertiary rule: if NOT within 14 days of the baseline deadline, bump current due +1 week (Mon adjust) ---
        if (newDue = "") {
            if (baseDeadline = "" || !IsWithinDays(now, baseDeadline, 14)) {
                baseForBump := curDue != "" ? curDue : now
                tmp := DateAdd(baseForBump, 7, "Days")
                newDue := AdjustToMonday(tmp)
                LogMessage("Tertiary bump → +1 week (Mon adjust): " . Human(newDue))
            } else {
                LogMessage("Within 14 days of the baseline deadline; no tertiary bump.")
            }
        }

        if (newDue = "") {
            LogMessage("No change applied.")
            return
        }

        ; --- Apply to Outlook ---
        EnsureFlagged(item)
        WriteOutlookStartDate(item, newDue)
        WriteOutlookDue(item, newDue)
        WriteOutlookReminder(item, false)  ; Disable reminder

        ; Don't auto-save - let user review in dialog
        ; item.Save
        
        ; Open "Flag to" dialog for review (Ctrl+Shift+G)
        WinActivate("ahk_class rctrl_renwnd32")
        Sleep 100
        Send "^+g"
        
        LogMessage("Flag dialog opened with Start: " . Human(newDue) . " and Due: " . Human(newDue))

    } catch as e {
        MsgBox("MailFlagUpdate error:`n" . e.Message)
    } finally {
        running := false
        ToolTip()  ; Clear any lingering tooltips
        KeyWait("LWin")
        KeyWait("Shift")
    }
}

; =========================
; Outlook helpers
; =========================
GetSelectedOutlookItem() {
    ol := ""
    try {
        ol := ComObjActive("Outlook.Application")
    } catch as e1 {
        try {
            ol := ComObject("Outlook.Application")
        } catch as e2 {
            return ""
        }
    }
    exp := ""
    try {
        exp := ol.ActiveExplorer
    } catch as e3 {
        return ""
    }
    if !exp
        return ""
    sel := ""
    try {
        sel := exp.Selection
    } catch as e4 {
        return ""
    }
    if !sel || sel.Count = 0
        return ""
    return sel.Item(1)
}

EnsureFlagged(item) {
    ; FlagStatus: 0=None, 1=Complete, 2=Marked
    try {
        if (item.FlagStatus != 2) {
            item.FlagStatus := 2
        }
    } catch as e {
        ; ignore
    }
    try {
        ; FlagIcon 1 = Red (optional)
        if (!item.FlagIcon) {
            item.FlagIcon := 1
        }
    } catch as e {
        ; ignore
    }
}

ReadOutlookDate(item, propName) {
    ; Returns AHK timestamp (YYYYMMDDHH24MISS) or ""
    try {
        dt := item.%propName%
        if dt {
            return NormalizeToStamp(dt)
        }
    } catch as e {
        ; ignore
    }
    return ""
}

WriteOutlookStartDate(item, stamp) {
    try {
        item.TaskStartDate := FormatTime(stamp, "yyyy-MM-dd HH:mm")
    } catch as e {
        throw Error("Failed to write TaskStartDate: " . e.Message)
    }
}

WriteOutlookDue(item, stamp) {
    try {
        item.FlagDueBy := FormatTime(stamp, "yyyy-MM-dd HH:mm")
    } catch as e {
        throw Error("Failed to write FlagDueBy: " . e.Message)
    }
}

WriteOutlookReminder(item, enable) {
    try {
        if (enable) {
            item.ReminderSet := true
            item.ReminderTime := FormatTime(enable, "yyyy-MM-dd HH:mm")
        } else {
            item.ReminderSet := false
        }
    } catch as e {
        throw Error("Failed to write ReminderTime: " . e.Message)
    }
}

; =========================
; Date utilities (all return/use YYYYMMDDHH24MISS)
; =========================
NormalizeToStamp(dt) {
    try {
        return FormatTime(dt, "yyyyMMddHHmmss")
    } catch as e {
        parsed := ParseFirstDate(dt)
        if (parsed != "")
            return parsed
        return ""
    }
}

IsWithinDays(stampA, stampB, days) {
    if (stampA = "" || stampB = "")
        return false
    diff := Abs(DateDiff(stampA, stampB, "Days"))
    return diff <= days
}

AdjustToMonday(stamp) {
    w := Integer(FormatTime(stamp, "WDay")) ; 1=Sun, 7=Sat
    if (w = 7)
        return DateAdd(stamp, 2, "Days")
    if (w = 1)
        return DateAdd(stamp, 1, "Days")
    return stamp
}

DateAtHour(stamp, hour) {
    y := SubStr(stamp, 1, 4)
    m := SubStr(stamp, 5, 2)
    d := SubStr(stamp, 7, 2)
    hh := Format("{:02}", hour)
    return y m d hh "0000"
}

Human(stamp) {
    return FormatTime(stamp, "ddd, MMM d yyyy HH:mm")
}

Tip(msg, on) {
    if on
        ToolTip(msg, 20, 20)
}

; -------------------------
; ParseFirstDate(text) — returns YYYYMMDDHH24MISS or ""
; Supports: yyyy-mm-dd, mm/dd/yyyy, "Sep 30, 2025", "30 Sep 2025", etc.
; Time defaults to 09:00.
; -------------------------
ParseFirstDate(text) {
    if !IsSet(text) || !text
        return ""

    ; ISO: 2025-09-30 or 2025/09/30
    if (RegExMatch(text, "\b(\d{4})[-/](\d{1,2})[-/](\d{1,2})\b", &m)) {
        y := m[1], mo := m[2], d := m[3]
        result := ValidateStamp(y, mo, d)
        LogMessage("ISO match: " . y . "/" . mo . "/" . d . " → " . result)
        return result
    }

    ; US: 9/30/2025 or 09/30/25
    if (RegExMatch(text, "\b(\d{1,2})/(\d{1,2})/(\d{2,4})\b", &m)) {
        mo := m[1], d := m[2], y := m[3]
        if (StrLen(y) = 2)
            y := (y < 70 ? "20" y : "19" y)
        result := ValidateStamp(y, mo, d)
        LogMessage("US match: " . mo . "/" . d . "/" . y . " → " . result)
        return result
    }

    ; Month name (Sep 30, 2025)
    months := Map("jan",1,"feb",2,"mar",3,"apr",4,"may",5,"jun",6,"jul",7,"aug",8,"sep",9,"oct",10,"nov",11,"dec",12)
    if (RegExMatch(text, "\b([A-Za-z]{3,9})\s+(\d{1,2})(?:st|nd|rd|th)?(?:,)?\s+(\d{4})\b", &m)) {
        moStr := StrLower(SubStr(m[1],1,3)), d := m[2], y := m[3]
        if months.Has(moStr)
            return ValidateStamp(y, months[moStr], d)
    }

    ; Day Month Year (30 Sep 2025)
    if (RegExMatch(text, "\b(\d{1,2})\s+([A-Za-z]{3,9})(?:,)?\s+(\d{4})\b", &m)) {
        d := m[1], moStr := StrLower(SubStr(m[2],1,3)), y := m[3]
        if months.Has(moStr)
            return ValidateStamp(y, months[moStr], d)
    }

    return ""
}


ValidateStamp(y, mo, d) {
    y  := Integer(y)
    mo := Integer(mo)
    d  := Integer(d)
    if (y < 1900 || y > 2100)
        return ""
    if (mo < 1 || mo > 12)
        return ""

    ; last day of month = (first of next month) - 1 day
    firstOfMonth := Format("{:04}{:02}01000000", y, mo)
    lastOfMonth  := DateAdd(DateAdd(firstOfMonth, 1, "Months"), -1, "Days")
    maxDay       := Integer(SubStr(lastOfMonth, 7, 2))

    if (d < 1)
        d := 1
    if (d > maxDay)
        d := maxDay

    return Format("{:04}{:02}{:02}090000", y, mo, d)
}

; Clear any lingering tooltips on script exit
OnExit(CleanupTooltips)

CleanupTooltips(*) {
    ToolTip()
}
