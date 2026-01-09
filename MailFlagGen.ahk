; MailFlagGen.ahk  –  AutoHotkey v2
#Requires AutoHotkey v2
#SingleInstance Force
#Warn
SetTitleMatchMode 2

; ---------- USER SETTINGS ----------
DefaultOffsetDays := 7   ; if no date found → due = today + 7 days
DefaultDueHour    := 9   ; 09:00 local
; ------------------------------------

; Date regex alternatives:
;  1) YYYY-MM-DD
;  2) MM/DD/YYYY
;  3) MM/DD/YY
;  4) Month DD, YYYY    (comma optional)
;  5) Month DD          (no year)
DatePattern := "i)(\d{4}-\d{2}-\d{2})"                              ; ISO 2025-07-08
            . "|(\d{1,2}/\d{1,2}/\d{4})"                            ; 07/08/2025
            . "|(\d{1,2}/\d{1,2}/\d{2})"                             ; 07/08/25
            . "|((January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}(?:,\s*|\s+)\d{4})" ; Month DD, YYYY
            . "|((January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2})"                  ; Month DD

#HotIf WinActive("ahk_class rctrl_renwnd32")   ; Outlook main window
^+f:: AddSmartFlag()
#HotIf


; ------------ helper functions ------------
ParseToTimestamp(str) {
    str := Trim(str)

    ; ---- ISO YYYY-MM-DD ----
    if RegExMatch(str, "^\d{4}-\d{2}-\d{2}$") {
        return StrReplace(str, "-", "") . "000000"
    }

    ; ---- US MM/DD/YYYY ----
    if RegExMatch(str, "^\d{1,2}/\d{1,2}/\d{4}$") {
        p := StrSplit(str, "/")
        y := p[3], m := Format("{:02}", p[1]), d := Format("{:02}", p[2])
        return y . m . d . "000000"
    }

    ; ---- US MM/DD/YY (assume 20YY) ----
    if RegExMatch(str, "^\d{1,2}/\d{1,2}/\d{2}$") {
        p := StrSplit(str, "/")
        y := 2000 + Integer(p[3])
        m := Format("{:02}", p[1]), d := Format("{:02}", p[2])
        return y . m . d . "000000"
    }

    ; ---- Month DD, YYYY (comma optional) ----
    if RegExMatch(str, "i)^(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2})(?:,\s*|\s+)(\d{4})$", &m) {
        monthMap := Map(
            "January","01","February","02","March","03","April","04","May","05","June","06",
            "July","07","August","08","September","09","October","10","November","11","December","12")
        mon := monthMap[m[1]]
        day := Format("{:02}", m[2])
        yr  := m[3]
        return yr . mon . day . "000000"
    }

    ; ---- Month DD (no year) → assume current year ----
    if RegExMatch(str, "i)^(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2})$", &m) {
        monthMap := Map(
            "January","01","February","02","March","03","April","04","May","05","June","06",
            "July","07","August","08","September","09","October","10","November","11","December","12")
        mon := monthMap[m[1]]
        day := Format("{:02}", m[2])
        yr  := FormatTime(A_Now, "yyyy")
        return yr . mon . day . "000000"
    }

    return ""
}

ToOleFmt(ts) => FormatTime(ts, "yyyy-MM-dd HH:mm")

NextWeekday(ts) {                     ; bump Sat/Sun → Monday
    wd := FormatTime(ts, "ddd")       ; "Mon" … "Sun"
    if (wd = "Sat")
        return DateAdd(ts, 2, "Days")
    if (wd = "Sun")
        return DateAdd(ts, 1, "Days")
    return ts
}

; ---------- helper: add N business days (skips Sat/Sun) ----------
AddBusinessDays(ts, n) {                 ; ts = yyyyMMddHHmmss
    while (n > 0) {
        ts := DateAdd(ts, 1, "Days")
        wd := FormatTime(ts, "ddd")      ; "Mon" … "Sun"
        if (wd != "Sat" && wd != "Sun")
            n -= 1
    }
    return ts
}
; -----------------------------------------------------------------

MakeDefaultDue() {
    global DefaultOffsetDays, DefaultDueHour
    base := DateAdd(A_Now, DefaultOffsetDays, "Days")
    due  := FormatTime(base, "yyyyMMdd") . Format("{:02}0000", DefaultDueHour)
    return NextWeekday(due)
}
; -------------------------------------------


; ---------- helper: split text into sentences ----------
SplitSentences(txt) {
    ; Convert sentence-ending punctuation (period, question mark, semicolon)
    ; followed by whitespace OR any line break into a newline, then split.
    tmp := RegExReplace(txt, "(?<=[\.\?;])\s+|\R", "`n")
    ; Trim each element and discard blanks
    arr := []
    for , s in StrSplit(tmp, "`n")
        if (s := Trim(s))
            arr.Push(s)
    return arr
}
; -------------------------------------------------------


; -------------- NEW AddSmartFlag() ---------------------
AddSmartFlag() {
    global DatePattern

    ; Pattern without the leading “i)” for inline concatenation
    dateBare := SubStr(DatePattern, 3)

    ; ---------- Outlook handles ----------
    try outlook := ComObjActive("Outlook.Application")
    catch {
        MsgBox("Outlook is not running.", "MailFlagGen", "Iconx")
        return
    }
    try mail := outlook.ActiveExplorer.Selection.Item(1)
    catch {
        MsgBox("Click a single email first.", "MailFlagGen", "Iconx")
        return
    }
    ; -------------------------------------

    subject := mail.Subject
    body    := mail.Body

    ; ======== 1) SPECIAL “AIR” SUBJECT RULE (2 BUSINESS DAYS) ========
    if RegExMatch(subject, "i)USPTO\s+Automated\s+Interview\s+Request\s+\(AIR\)") {
        flagText := "remind re EXI req"
        dueTS    := AddBusinessDays(FormatTime(A_Now, "yyyyMMddHHmmss"), 2)
        ApplyFlag(mail, flagText, dueTS, dueTS)
        return
    }
    ; =================================================================

    ; ======== 2) SMART DEADLINE DATE PICK (skip “Filing Date” lines) ==
    bodyClean := RegExReplace(body, "(?im)^.*\bfiling\s*date\b.*$", "")
    foundStr  := ""
    posOrig   := 0

    for s in SplitSentences(bodyClean) {
        if !RegExMatch(s, "i)\bdeadline\b")
            continue
        ; (a) date AFTER the word “deadline”
        if RegExMatch(s, "i)\bdeadline\b.*?" . dateBare, &mAfter) {
            foundStr := mAfter[0]
            posOrig  := InStr(body, foundStr)
            break
        }
        ; (b) otherwise first date in that sentence
        if RegExMatch(s, DatePattern, &mAny) {
            foundStr := mAny[0]
            posOrig  := InStr(body, foundStr)
            break
        }
    }

    ; Fallback: first date in bodyClean, then subject
    if (foundStr = "") {
        if RegExMatch(bodyClean, DatePattern, &mFB) {
            foundStr := mFB[0]
            posOrig  := InStr(body, foundStr)
        } else if RegExMatch(subject, DatePattern, &mFB) {
            foundStr := mFB[0]
            posOrig  := 1
        }
    }
    ; =================================================================

    ; ======== 3) PARSE, BUILD FLAG TEXT, APPLY SUFFIX RULES ==========
    if (foundStr != "") {
        foundTS := ParseToTimestamp(foundStr)   ; yyyyMMddHHmmss
        haveDate := (foundTS != "")
    } else {
        haveDate := false
    }

    if haveDate {
        showDate := FormatTime(foundTS, "MM/dd/yy")
        flagText := "remind re " showDate " dl"

        ; Base due/start = 14 days earlier, bumped off weekends
        dueTS   := NextWeekday(DateAdd(foundTS, -14, "Days"))
        startTS := dueTS

        ; Build context around the chosen date for keyword checks
        ctx := SubStr(body, Max(posOrig-200,1), 400)

        ; ---- suffix logic ----
        hasISF := RegExMatch(ctx, "i)\bissue\s+fee\b") || RegExMatch(body, "i)\bissue\s+fee\b")

        ; rr when either “restriction requirement” OR “four/4-month(s)” is present
        hasRR  := RegExMatch(ctx, "i)\brestriction\s+requirement\b")
               || RegExMatch(body, "i)\brestriction\s+requirement\b")
               || RegExMatch(ctx, "i)\b(?:four|4)[-\s]?month(?:s)?\b")
               || RegExMatch(body, "i)\b(?:four|4)[-\s]?month(?:s)?\b")

        hasOA  := RegExMatch(ctx, "i)\boffice\s+action\b")
               || RegExMatch(subject, "i)\boffice\s+action\b")

        if hasRR {
            flagText .= " 2 mo rr"         ; overrides OA’s “3 mo”
        } else if hasOA {
            flagText .= " 3 mo "
        }
        if hasISF
            flagText .= " ISF"
        ; ----------------------
    } else {
        flagText := "Follow-up"
        startTS  := MakeDefaultDue()
        dueTS    := startTS
    }
    ; =================================================================

    ; ======== 4) APPLY FLAG & OPEN DIALOG =============================
    ApplyFlag(mail, flagText, startTS, dueTS)
}
; -------------- END AddSmartFlag() ---------------------

; ---------- flag + dialog ----------
ApplyFlag(mail, txt, startTS, dueTS) {
    mail.FlagRequest   := txt
    mail.FlagIcon      := 2          ; red flag
    mail.TaskStartDate := ToOleFmt(startTS)
    mail.TaskDueDate   := ToOleFmt(dueTS)
    mail.ReminderSet   := false
    mail.FlagStatus    := 2          ; marked
    mail.Save()

    ; Open “Flag to” dialog for review (Ctrl+Shift+G)
    WinActivate("ahk_class rctrl_renwnd32")
    Sleep 100
    Send "^+g"
}
