; MailReminder.ahk  (AutoHotkey v2)
; - Reply-All reminder generator
; - Draft OA reply template (with extensions + “no need to file …” line)
; - OA reminder template with dynamic roll-forward of expired deadlines
; - Greeting from your most recent sent chunk (ends in -Grant or “Best regards,↵Grant”)
; - HTML paragraphs (<p>), simple markup (**bold**, _italic_, __underline__)
; - Deadlines bold+underlined; ack line italic with visible ***

#Requires AutoHotkey v2.0
#SingleInstance Force

DEFAULT_SALUTATION := "Best regards," . "`n-Grant J. STEYER"

; =========================
;        Hotkey
; =========================
#r::{
    olApp := GetOutlook()
    if !IsObject(olApp) {
        MsgBox("Outlook is not available.", "MailReminder")
        return
    }
    sel := olApp.ActiveExplorer.Selection
    if !sel || sel.Count = 0 {
        MsgBox("Please select an email in Outlook first.", "MailReminder")
        return
    }

    mail := sel.Item(1)
    subj := mail.Subject
    body := mail.Body
    if !body
        body := mail.HTMLBody

    ; Normalize
    body := StrReplace(body, "`r`n", "`n")
    body := StrReplace(body, "`r", "`n")
    lines := StrSplit(body, "`n")

    ; ---------- Identify chunks ----------
    top := GetTopChunkRange(lines)               ; current message at top
    grantAnyIdx := FindLastGrantLineIndex(lines) ; nearest "-Grant" anywhere

    ; ---------- Salutation ----------
    salutation := ""
    if (grantAnyIdx > 0) {
        salutation := ExtractSalutationFromGrant(lines, grantAnyIdx, DEFAULT_SALUTATION)
    } else {
        salutation := ExtractSalutationGeneralInRange(lines, top.start, top.end, DEFAULT_SALUTATION)
        if (!salutation)
            salutation := DEFAULT_SALUTATION
    }

    ; ---------- Greeting ----------
    greeting := ""
    if (grantAnyIdx > 0) {
        greeting := ExtractGreetingFromGrantChunk(lines, grantAnyIdx)
    }
    if (!greeting) {
        greeting := ExtractGreetingInRange(lines, top.start, top.end)
        if (!greeting) {
            greeting := ExtractGreetingGeneric(body)
            if !greeting
                greeting := "Hello,"
        }
    }

    ; --- NEW: specific trademark response/extension rule ---
    if HandleTrademarkExtDeadline(mail)
        return

    ; ---------- Build “Grant chunk” text for Draft-OA detection ----------
    topText := GetSliceText(lines, top.start, top.end)
    topLower := StrLower(topText)

    ; ------ Draft OA Reply template detection (STRICT) ------
    isReportIndicators :=
    (
        InStr(topLower, "enclosed is a copy")
     || InStr(topLower, "attached is a copy")
     || InStr(topLower, "we are reporting this office action")
     || InStr(topLower, "we received a non-final office action")
     || InStr(topLower, "we received a first office action")
    )

    isDraftOAReply :=
        RegExMatch(topLower, "\bdraft\b")
     && RegExMatch(topLower, "\b(reply|response)\b")
     && InStr(topLower, "office action")
     && !isReportIndicators

    if (isDraftOAReply) {
        draftDeadline := ExtractDateAfterKeyAny(topText, "extendable")
        if !draftDeadline
            draftDeadline := ExtractDateAfterKeyAny(topText, "deadline")
        if !draftDeadline
            draftDeadline := ExtractGenericDateAny(topText)
        if !draftDeadline
            draftDeadline := ExtractBestDeadline(body)

        remExtDraft := ExtractRemainingExtensionsAlt(topText)
        if (remExtDraft = 0)
            remExtDraft := ExtractRemainingExtensions(topText)
        if (remExtDraft = 0) {
            remExtDraft := ExtractRemainingExtensionsAlt(body)
            if (remExtDraft = 0)
                remExtDraft := ExtractRemainingExtensions(body)
        }
        ; If no extensions found, use default based on email content
        if (remExtDraft = 0) {
            remExtDraft := GetDefaultExtensionCount(body)
        }
        usedExtDraft := ExtractUsedExtensions(body)

        deadlineTok := "__**" . (draftDeadline ? draftDeadline : "the current deadline") . "**__"

        para := "As a kind reminder, please send us your instructions for responding to the Office Action, "
              . "keeping in mind the extendable " . deadlineTok . " deadline."
        if (remExtDraft > 0) {
            numWord := NumberWord(remExtDraft)
            moreTok := (usedExtDraft > 0) ? " more" : ""
            para .= "  It is possible to take " . numWord . moreTok . " extension" . (remExtDraft=1 ? "" : "s")
                 . " (for a total of " . numWord . " more month" . (remExtDraft=1 ? "" : "s") . ") for additional fees."
            para .= "  There is no need to file a request an extension of time if an extension is needed.  "
                 .  "Rather, we can simply pay the fee when responding to the Office Action at a later date."
        }

        paras := []
        paras.Push(greeting)
        paras.Push(para)
        for ln in SplitIntoLines(salutation)
            paras.Push(ln)

        CreateReplyAllPrepend(olApp, mail, subj, JoinWithNewlines(paras))
        return
    }

    ; ------ OA Reminder detection ------
    lowerBody := StrLower(body)
    isOA := InStr(lowerBody, "office action")
        && (
            InStr(lowerBody, "period for responding to the office action")
         || InStr(lowerBody, "deadline for responding to the office action")
         || InStr(lowerBody, "deadline to respond to the office action")
         || InStr(lowerBody, "extension of time expires on")
         || (InStr(lowerBody, "expires on") && InStr(lowerBody, "office action"))
        )

    if isOA {
        ; Use ONLY the top chunk of the client’s latest message, sanitized.
        topTextSan := SanitizeTopForParsing(topText)
		prevGrant := (grantAnyIdx > 0) ? GetPrevGrantText(lines, grantAnyIdx) : ""

        ; 1) Top chunk
		oaDeadlineStr := ExtractDateAfterKeyAny(topTextSan, "extension of time expires on")
		if !oaDeadlineStr
			oaDeadlineStr := ExtractDateAfterKeyAny(topTextSan, "deadline for responding to the office action is")
		if !oaDeadlineStr
			oaDeadlineStr := ExtractDateAfterKeyAny(topTextSan, "deadline for responding to the office action")
		if !oaDeadlineStr
			oaDeadlineStr := ExtractDateAfterKeyAny(topTextSan, "deadline to respond to the office action")
		if !oaDeadlineStr
			oaDeadlineStr := ExtractDateAfterKeyAny(topTextSan, "expires on")
		if !oaDeadlineStr
			oaDeadlineStr := ExtractGenericDateAny(topTextSan)

		; 2) Fallback to latest Grant chunk (your last sent block), NOT whole body
		if !oaDeadlineStr && prevGrant {
			oaDeadlineStr := ExtractDateAfterKeyAny(prevGrant, "extension of time expires on")
			if !oaDeadlineStr
				oaDeadlineStr := ExtractDateAfterKeyAny(prevGrant, "deadline for responding to the office action is")
			if !oaDeadlineStr
				oaDeadlineStr := ExtractDateAfterKeyAny(prevGrant, "deadline for responding to the office action")
			if !oaDeadlineStr
				oaDeadlineStr := ExtractDateAfterKeyAny(prevGrant, "deadline to respond to the office action")
			if !oaDeadlineStr
				oaDeadlineStr := ExtractDateAfterKeyAny(prevGrant, "expires on")
			if !oaDeadlineStr
				oaDeadlineStr := ExtractGenericDateAny(prevGrant)
		}

		; 3) Last resort: whole body AFTER stripping header stamps
		if !oaDeadlineStr {
			bodyClean := RemoveQuotedHeaders(body)
			oaDeadlineStr := ExtractDateAfterKeyAny(bodyClean, "extension of time expires on")
			if !oaDeadlineStr
				oaDeadlineStr := ExtractDateAfterKeyAny(bodyClean, "deadline for responding to the office action is")
			if !oaDeadlineStr
				oaDeadlineStr := ExtractDateAfterKeyAny(bodyClean, "deadline for responding to the office action")
			if !oaDeadlineStr
				oaDeadlineStr := ExtractDateAfterKeyAny(bodyClean, "deadline to respond to the office action")
			if !oaDeadlineStr
				oaDeadlineStr := ExtractDateAfterKeyAny(bodyClean, "expires on")
			if !oaDeadlineStr
				oaDeadlineStr := ExtractGenericDateAny(bodyClean)
		}


        ; Extensions available/used — only parse top chunk.
        usedExt := ExtractUsedExtensions(topTextSan)
        remExt  := ExtractRemainingExtensions(topTextSan)
        if (remExt = 0)
            remExt := ExtractRemainingExtensionsAlt(topTextSan)
			
			if (prevGrant && usedExt = 0 && remExt = 0) {
				usedExt := ExtractUsedExtensions(prevGrant)
				remExt  := ExtractRemainingExtensions(prevGrant)
				if (remExt = 0)
					remExt := ExtractRemainingExtensionsAlt(prevGrant)
			}
			
			; If no extensions found, use default based on email content
			if (remExt = 0) {
				remExt := GetDefaultExtensionCount(body)
			}


        ; ---- Roll the deadline forward if it already passed ----
        deadlineYMD := ParseDateYMD(oaDeadlineStr)
        advancedUsed := 0
        if (deadlineYMD) {
            while (A_Now > deadlineYMD && remExt > 0) {
                deadlineYMD := AddMonthsClamped(deadlineYMD, 1)
                usedExt += 1
                remExt  -= 1
                advancedUsed += 1
            }
        }

        shownDeadline := oaDeadlineStr
        if (deadlineYMD)
            shownDeadline := FormatDatePretty(deadlineYMD)

        deadlineTok := "__**" . (shownDeadline ? shownDeadline : "the current deadline") . "**__"

        paras := []
        paras.Push("_***Please acknowledge receipt of this email and kindly copy IP@RennerOtto.com on all correspondence***_")
        paras.Push(greeting)

        oaLine := ""
        if (usedExt > 0) {
            oaLine := "As a kind reminder, the period for responding to the Office Action within the "
                   .  WordOrdinal(usedExt) . " extension of time expires on " . deadlineTok . "."
        } else {
            oaLine := "As a kind reminder, the period for responding to the Office Action expires on " . deadlineTok . "."
        }

        if (remExt > 0) {
            numWord := NumberWord(remExt)
            moreTok := (usedExt > 0) ? " more" : ""
            oaLine .= "  It is possible to take " . numWord . moreTok . " extension" . (remExt=1 ? "" : "s")
                  .  " (for a total of " . numWord . " more month" . (remExt=1 ? "" : "s") . ") for additional fees."
            paras.Push(oaLine)
            paras.Push("There is no need to file a request an extension of time if an extension is needed.  Rather, we can simply pay the fee when responding to the Office Action at a later date.")
        } else {
            paras.Push(oaLine)
            paras.Push("__There are no further extensions available and the application will be abandoned if a reply is not filed.__")
        }

        for ln in SplitIntoLines(salutation)
            paras.Push(ln)

        CreateReplyAllPrepend(olApp, mail, subj, JoinWithNewlines(paras))
        return
    }

    ; ------ Default (non-OA) flow ------
    deadlineStr := ExtractBestDeadline(body)
    actionRaw   := ExtractRequestedAction(body)
    actionClean := TransformOrCleanAction(actionRaw)

    hasPlease := RegExMatch(StrLower(actionClean), "^\s*please\b")
    reminder := "A kind reminder to " . (hasPlease ? "" : "please ") . actionClean
    if deadlineStr
        reminder .= ", keeping in mind the " . "__**" . deadlineStr . "**__" . "."
    else
        reminder .= "."

    paras := []
    paras.Push(greeting)
    paras.Push(reminder)
    for ln in SplitIntoLines(salutation)
        paras.Push(ln)

    CreateReplyAllPrepend(olApp, mail, subj, JoinWithNewlines(paras))
}

; ===================== NEW RULE: Trademark OA/Extension Deadline Reminder =====================

MatchesTrademarkOAExtDeadline(txt) {
    txtl := StrLower(txt)
    return InStr(txtl, "deadline to file a response or a request for an extension of time")
        && InStr(txtl, "trademark application")
}

ExtractDeadline(txt) {
    months := "January|February|March|April|May|June|July|August|September|October|November|December|Jan\.?|Feb\.?|Mar\.?|Apr\.?|May|Jun\.?|Jul\.?|Aug\.?|Sep\.?|Sept\.?|Oct\.?|Nov\.?|Dec\.?"
    re1 := "(" months ")\s+\d{1,2},?\s+\d{4}"
    if RegExMatch(txt, re1, &m1) {
        return Trim(m1[0], ".,;: ")
    }
    re2 := "\b\d{1,2}/\d{1,2}/\d{2,4}\b"
    if RegExMatch(txt, re2, &m2) {
        return m2[0]
    }
    return ""
}

BuildTrademarkExtDeadlineBody(deadline) {
    if (deadline = "")
        deadline := "[[Deadline]]"
    body := "A kind reminder that the deadline to file a response or a request for an extension of time is " deadline ". If a response or extension is not filed by " deadline ", this application will become abandoned.`r`n"
    body .= "Unless we receive your instructions to the contrary, we will take no further action and allow this application to go abandoned."
    return body
}

_GetPlainText(item) {
    txt := ""
    try {
        txt := item.Body
    } catch as e {
    }
    if (txt != "" && StrLen(Trim(txt)) > 0)
        return txt
    try {
        html := item.HTMLBody
        if (html != "") {
            txt := RegExReplace(html, "<[^>]+>", " ")
            txt := RegExReplace(txt, "\s{2,}", " ")
            return txt
        }
    } catch as e {
    }
    return ""
}

HandleTrademarkExtDeadline(item) {
    bodyText := ""
    try {
        bodyText := _GetPlainText(item)
    } catch as e {
        return false
    }

    if !MatchesTrademarkOAExtDeadline(bodyText)
        return false

    deadline := ExtractDeadline(bodyText)

    try {
        app  := item.Application
        mail := app.CreateItem(0)  ; olMailItem

        ; If you have custom hooks, you can replace the lines below.
        ; These defaults avoid dependency on external helpers.
        mail.To := item.SenderEmailAddress
        mail.CC := ""

        subjBase := "Trademark response/extension deadline"
        mail.Subject := subjBase

        greet := "Hello,"
        sign  := "-Grant"

        bodyCore := BuildTrademarkExtDeadlineBody(deadline)
        mail.Body := greet "`r`n`r`n" bodyCore "`r`n`r`n" sign

        mail.Display()
        return true
    } catch as e {
        try {
            MsgBox "Error creating Trademark deadline reminder:`n" e.Message
        } catch {
        }
        return false
    }
}
; ===================== END RULE =====================

; =========================
;   Helper: Sanitize top chunk
; =========================
SanitizeTopForParsing(s) {
    lines := SplitIntoLines(s)
    out := ""
    for i, ln in lines {
        t := Trim(ln)
        if (t = "")
        {
            out .= (out ? "`n" : "") . ""
            continue
        }
        if RegExMatch(t, "i)^\*\*\*Please acknowledge receipt")
            continue
        if RegExMatch(t, "i)^As a kind reminder,")
            continue
        if RegExMatch(t, "i)^There is no need to file a request an extension of time")
            continue
        out .= (out ? "`n" : "") . ln
    }
    return out
}

; =========================
;   Date extractors with Month DD, YYYY and DD Month YYYY
; =========================
ExtractDateAfterKeyAny(body, key) {
    m := ""
    ; Month DD, YYYY
    pat1 := "i)" . key . ".*?\b(January|February|March|April|May|June|July|August|September|October|November|December)\b\s+\d{1,2},\s+\d{4}"
    if RegExMatch(body, pat1, &m) {
        if RegExMatch(m[0], "i)\b(January|February|March|April|May|June|July|August|September|October|November|December)\b\s+\d{1,2},\s+\d{4}", &d)
            return Trim(d[0])
    }
    ; DD Month YYYY
    pat2 := "i)" . key . ".*?\b\d{1,2}\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}"
    if RegExMatch(body, pat2, &m) {
        if RegExMatch(m[0], "i)\b\d{1,2}\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}", &d)
            return Trim(d[0])
    }
    ; Numeric M/D/(YY)YY
    pat3 := "i)" . key . ".*?\b\d{1,2}/\d{1,2}/\d{2,4}\b"
    if RegExMatch(body, pat3, &m) {
        if RegExMatch(m[0], "\b\d{1,2}/\d{1,2}/\d{2,4}\b", &d)
            return Trim(d[0])
    }
    return ""
}

ExtractGenericDateAny(txt) {
    m := ""
    if RegExMatch(txt, "i)\b(January|February|March|April|May|June|July|August|September|October|November|December)\b\s+\d{1,2},\s+\d{4}", &m)
        return Trim(m[0])
    if RegExMatch(txt, "i)\b\d{1,2}\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}", &m)
        return Trim(m[0])
    if RegExMatch(txt, "i)\b\d{1,2}/\d{1,2}/\d{2,4}\b", &m)
        return Trim(m[0])
    return ""
}

ParseDateYMD(dateStr) {
    if !dateStr
        return ""
    m := ""
    ; Month DD, YYYY
    if RegExMatch(dateStr, "i)(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),\s+(\d{4})", &m) {
        mon := MonthToNum(m[1])
        day := Format("{:02}", m[2] + 0)
        yr  := m[3]
        return yr . mon . day . "000000"
    }
    ; DD Month YYYY
    if RegExMatch(dateStr, "i)(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})", &m) {
        day := Format("{:02}", m[1] + 0)
        mon := MonthToNum(m[2])
        yr  := m[3]
        return yr . mon . day . "000000"
    }
    ; Numeric M/D/(YY)YY
    if RegExMatch(dateStr, "i)\b(\d{1,2})/(\d{1,2})/(\d{2,4})\b", &m) {
        mon := Format("{:02}", m[1] + 0)
        day := Format("{:02}", m[2] + 0)
        yr  := (StrLen(m[3]) = 2) ? ("20" . m[3]) : m[3]
        return yr . mon . day . "000000"
    }
    return ""
}

FormatDatePretty(yyyymmddHHmiss) {
    if (StrLen(yyyymmddHHmiss) < 8)
        return ""
    yyyy := SubStr(yyyymmddHHmiss, 1, 4)
    mm   := SubStr(yyyymmddHHmiss, 5, 2)
    dd   := SubStr(yyyymmddHHmiss, 7, 2)
    ts := yyyy . mm . dd . "000000"
    return FormatTime(ts, "MMMM d, yyyy")
}

; EOM-safe +months (Jan 30 → Feb 28/29)
AddMonthsClamped(yyyymmddHHmiss, months) {
    y  := Integer(SubStr(yyyymmddHHmiss, 1, 4))
    m  := Integer(SubStr(yyyymmddHHmiss, 5, 2))
    d  := Integer(SubStr(yyyymmddHHmiss, 7, 2))
    hh := Integer(SubStr(yyyymmddHHmiss, 9, 2))
    mi := Integer(SubStr(yyyymmddHHmiss,11, 2))
    ss := Integer(SubStr(yyyymmddHHmiss,13, 2))

    totalM := m + months
    newY := y + Floor((totalM - 1) / 12)
    newM := Mod(totalM - 1, 12) + 1
    dim := DaysInMonth(newY, newM)
    newD := (d > dim) ? dim : d

    return Format("{:04}{:02}{:02}{:02}{:02}{:02}", newY, newM, newD, hh, mi, ss)
}

DaysInMonth(y, m) {
    nextY := y, nextM := m + 1
    if (nextM > 12) {
        nextM := 1
        nextY := y + 1
    }
    start := Format("{:04}{:02}01000000", y, m)
    next  := Format("{:04}{:02}01000000", nextY, nextM)
    return DateDiff(next, start, "D")
}

MonthToNum(txt) {
    txt := StrLower(txt)
    abbr := SubStr(txt, 1, 3)
    switch abbr {
        case "jan": return "01"
        case "feb": return "02"
        case "mar": return "03"
        case "apr": return "04"
        case "may": return "05"
        case "jun": return "06"
        case "jul": return "07"
        case "aug": return "08"
        case "sep": return "09"
        case "oct": return "10"
        case "nov": return "11"
        case "dec": return "12"
        default:    return "01"
    }
}

; =========================
;     Chunk & boundary helpers
; =========================
GetTopChunkRange(lines) {
    i := 1
    while (i <= lines.Length) {
        ln := Trim(lines[i])
        if RegExMatch(ln, "i)^(from:|sent:|to:|cc:|subject:|attachments|caution:|-----original message-----)")
            return { start: 1, end: i-1 }
        i++
    }
    return { start: 1, end: lines.Length }
}

FindGrantLineIndexInRange(lines, start, end) {
    i := start
    last := 0
    while (i <= end) {
        ln := Trim(lines[i])
        if RegExMatch(ln, "i)^\s*-\s*grant\b")
            last := i
        i++
    }
    return last
}

; =========================
;   Greeting / Salutation
; =========================
ExtractGreetingInRange(lines, start, end) {
    ; Prefer Dear/Hi/Hello inside the range
    i := start
    while (i <= end) {
        ln := Trim(lines[i])
        if (ln != "" && SubStr(ln,1,3) != "***" && RegExMatch(ln, "i)^(dear|hi|hello)\b"))
            return TrimToGreetingPhrase(ln)
        i++
    }
    ; Fallback: first short line ending with ':' or ','
    i := start
    while (i <= end) {
        ln := Trim(lines[i])
        if RegExMatch(ln, "^[^\r\n]{2,60}[:|,]\s*$")
            return TrimToGreetingPhrase(ln)
        i++
    }
    return ""
}

ExtractGreetingGeneric(body) {
    lines := StrSplit(body, "`n")
    maxScan := (lines.Length < 120) ? lines.Length : 120
    loop maxScan {
        ln := Trim(lines[A_Index])
        if (ln = "" || SubStr(ln,1,3) = "***")
            continue
        if RegExMatch(ln, "i)^(dear|hi|hello)\b")
            return TrimToGreetingPhrase(ln)
        if RegExMatch(ln, "^[^\r\n]{2,60}[:|,]\s*$")
            return TrimToGreetingPhrase(ln)
    }
    return ""
}

; Use the last “-Grant …” chunk: go up to its preceding Subject:, then scan for greeting
ExtractGreetingFromGrantChunk(lines, grantIdx) {
    rg := GetGrantChunkRange(lines, grantIdx)
    start := rg.start, end := rg.end - 1
    if (end < start)
        end := start
    ; 1) Dear/Hi/Hello
    i := start
    while (i <= end) {
        ln := Trim(lines[i])
        if RegExMatch(ln, "i)^(dear|hi|hello)\b")
            return TrimToGreetingPhrase(ln)
        i++
    }
    ; 2) First line that ends with ":" or "," (2–60 chars)
    i := start
    while (i <= end) {
        ln := Trim(lines[i])
        if RegExMatch(ln, "^[^\r\n]{2,60}[:|,]\s*$")
            return TrimToGreetingPhrase(ln)
        i++
    }
    return ""
}

TrimToGreetingPhrase(ln) {
    m := ""
    if RegExMatch(ln, "i)^(.*?[,:])", &m)
        return Trim(m[1])
    return Trim(ln)
}

; Find slice [Subject_next_line .. grantIdx]
GetGrantChunkRange(lines, grantIdx) {
    subjIdx := FindPrevSubjectIndex(lines, grantIdx)
    start := (subjIdx > 0) ? subjIdx + 1 : 1
    return { start: start, end: grantIdx }
}

FindPrevSubjectIndex(lines, fromIdx) {
    i := fromIdx
    while (i > 0) {
        ln := Trim(lines[i])
        if RegExMatch(ln, "i)^subject:")
            return i
        i--
    }
    return 0
}

; Last line matching "-Grant" anywhere
FindLastGrantLineIndex(lines) {
    idx := 0
    loop lines.Length {
        ln := Trim(lines[A_Index])
        if RegExMatch(ln, "i)^\s*-\s*grant\b")
            idx := A_Index
    }
    return idx
}

; Salutation from "-Grant": include previous non-empty line if its length <= 3,
; else force "Best regards," above "-Grant".
ExtractSalutationFromGrant(lines, grantIdx, defaultSal) {
    anchor := Trim(lines[grantIdx])
    i := grantIdx - 1
    prev := ""
    while (i > 0) {
        t := Trim(lines[i])
        if (t != "") {
            prev := t
            break
        }
        i--
    }
    if (prev = "")
        return defaultSal
    if (StrLen(prev) > 3)
        return "Best regards," . "`n" . anchor
    return prev . "`n" . anchor
}

; Salutation when no "-Grant": look for "Best regards," followed by "Grant" (no dash)
ExtractSalutationGeneralInRange(lines, start, end, defaultSal) {
    i := start
    while (i <= end - 1) {
        ln := Trim(lines[i])
        nxt := Trim(lines[i+1])
        if (RegExMatch(ln, "i)^best\s+regards,?$") && RegExMatch(nxt, "i)^grant\b"))
            return ln . "`n" . nxt
        i++
    }
    return ""
}

; =========================
;     Requested Action
; =========================
ExtractRequestedAction(body) {
    m := ""
    if RegExMatch(body, "i)\bplease\s+(have|send|confirm|review|advise|provide|sign|execute|let\s+us\s+know)\b.+?(?:(?<=\.)|(?<=\!)|(?<=\?)|(?=\R))", &m)
        return CleanSentence(m[0])
    if RegExMatch(body, "i)\bwe\s+look\s+forward\s+to\s+receiving\b.+?(?:(?<=\.)|(?<=\!)|(?<=\?)|(?=\R))", &m)
        return CleanSentence(m[0])
    if RegExMatch(body, "i).*\bplease\b.+?(?:(?<=\.)|(?<=\!)|(?<=\?)|(?=\R))", &m)
        return CleanSentence(m[0])
    return "send us your instructions"
}

CleanSentence(s) {
    s := Trim(s)
    s := RegExReplace(s, "\s+", " ")
    return s
}

TransformOrCleanAction(action) {
    actLow := StrLower(action)
    if InStr(actLow, "document") && RegExMatch(actLow, "i)\bexecut") && InStr(actLow, "return")
        return "please send us the executed documents"
    action := RegExReplace(action, "i)^\s*(please|kindly)\b\s*(to\s+)?", "")
    action := Trim(action)
    action := RegExReplace(action, "[\.\!\:\;]+$")
    if (StrLen(action) >= 1)
        action := StrLower(SubStr(action, 1, 1)) . SubStr(action, 2)
    return action
}

; =========================
;        Deadlines (generic)
; =========================
ExtractBestDeadline(body) {
    d := ExtractDateAfterKey(body, "deadline")
    if d
        return d
    d := ExtractDateAfterKey(body, "expires on")
    if d
        return d
    d := ExtractDateAfterKey(body, "due")
    if d
        return d
    d := ExtractGenericDate(body)
    return d ? d : ""
}

ExtractDateAfterKey(body, key) {
    pat := "i)" . key . ".*?\b(January|February|March|April|May|June|July|August|September|October|November|December)\b\s+\d{1,2},\s+\d{4}"
    m := ""
    if RegExMatch(body, pat, &m) {
        if RegExMatch(m[0], "i)\b(January|February|March|April|May|June|July|August|September|October|November|December)\b\s+\d{1,2},\s+\d{4}", &d)
            return Trim(d[0])
    }
    return ""
}

ExtractGenericDate(txt) {
    m := ""
    if RegExMatch(txt, "i)\b(January|February|March|April|May|June|July|August|September|October|November|December)\b\s+\d{1,2},\s+\d{4}", &m)
        return Trim(m[0])
    return ""
}

ParseDateYMD_basic(dateStr) {
    if !dateStr
        return ""
    m := ""
    if RegExMatch(dateStr, "i)(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),\s+(\d{4})", &m) {
        mon := MonthToNum(m[1])
        day := Format("{:02}", m[2] + 0)
        yr  := m[3]
        return yr . mon . day . "000000"
    }
    return ""
}

; =========================
;   Office Action fields
; =========================

; Determine default extension count based on email content
GetDefaultExtensionCount(body) {
    bodyLower := StrLower(body)
    
    ; Check if deadline is explicitly stated as final or non-extendable
    if (InStr(bodyLower, "final") && InStr(bodyLower, "deadline")) {
        return 0
    }
    if (InStr(bodyLower, "non-extendable") || InStr(bodyLower, "non extendable")) {
        return 0
    }
    if (InStr(bodyLower, "no extensions") || InStr(bodyLower, "cannot be extended")) {
        return 0
    }
    
    ; Check for "restriction requirement" - gives 4 extensions
    if (InStr(bodyLower, "restriction requirement")) {
        return 4
    }
    
    ; Default case - 3 extensions available
    return 3
}

; Test function to verify extension logic (can be removed in production)
TestExtensionLogic() {
    ; Test cases
    test1 := "This is a regular office action with a deadline."
    test2 := "This office action has a final deadline that cannot be extended."
    test3 := "This is a restriction requirement office action."
    test4 := "This office action is non-extendable."
    
    MsgBox("Test 1 (regular): " . GetDefaultExtensionCount(test1) . " extensions`n" .
           "Test 2 (final): " . GetDefaultExtensionCount(test2) . " extensions`n" .
           "Test 3 (restriction): " . GetDefaultExtensionCount(test3) . " extensions`n" .
           "Test 4 (non-extendable): " . GetDefaultExtensionCount(test4) . " extensions")
}

ExtractUsedExtensions(body) {
    m := ""
    if RegExMatch(body, "i)within\s+the\s+((?:first|second|third|fourth)|(?:\d+(?:st|nd|rd|th)))\s+extension\s+of\s+time", &m) {
        w := StrLower(m[1])
        num := RegExReplace(w, "\D")
        if (num != "")
            return Integer(num)
        switch w {
            case "first":  return 1
            case "second": return 2
            case "third":  return 3
            case "fourth": return 4
        }
    }
    return 0
}

ExtractRemainingExtensions(body) {
    m := ""
    if RegExMatch(body, "i)It\s+is\s+possible\s+to\s+take\s+(one|two|three|four|five|\d+)\s+(?:additional\s+)?extensions\b", &m) {
        w := StrLower(m[1])
        switch w {
            case "one":   return 1
            case "two":   return 2
            case "three": return 3
            case "four":  return 4
            case "five":  return 5
            default:      return Integer(w)
        }
    }
    return 0
}

ExtractRemainingExtensionsAlt(body) {
    m := ""
    if RegExMatch(body, "i)(?:may\s+be\s+|can\s+be\s+)?extended\s+(?:for\s+up\s+to|up\s+to)\s+(one|two|three|four|five|\d+)(?:\s*\(\s*\d+\s*\))?\s+months?", &m)
        return NumberWordToInt(m[1])
    if RegExMatch(body, "i)increments\s+of\s+one\s+month\s+up\s+to\s+a\s+total\s+of\s+(one|two|three|four|five|\d+)\s+months?", &m)
        return NumberWordToInt(m[1])
    if RegExMatch(body, "i)up\s+to\s+a\s+total\s+of\s+(one|two|three|four|five|\d+)\s+months?", &m)
        return NumberWordToInt(m[1])
    return 0
}

WordOrdinal(n) {
    switch n {
        case 1: return "first"
        case 2: return "second"
        case 3: return "third"
        case 4: return "fourth"
        default: return n . "th"
    }
}

NumberWord(n) {
    switch n {
        case 1: return "one"
        case 2: return "two"
        case 3: return "three"
        case 4: return "four"
        case 5: return "five"
        default: return n
    }
}

NumberWordToInt(w) {
    w := StrLower(w)
    switch w {
        case "one":   return 1
        case "two":   return 2
        case "three": return 3
        case "four":  return 4
        case "five":  return 5
        default:      return Integer(w)
    }
}

RenderHtml(text) {
    q := Chr(34)  ; double-quote character
    lines := SplitIntoLines(text)
    buf := "<div>"
    for _, ln in lines {
        ; detect the "no further extensions..." line (while it still has leading __)
        shouldHighlight := RegExMatch(ln, "i)^__There are no further extensions available")
        htmlLine := MarkupToHtml(ln)  ; convert __, **, _ to <u>, <strong>, <em>

        if (shouldHighlight) {
            htmlLine := "<span style=" . q . "background-color:yellow" . q . ">" . htmlLine . "</span>"
        }

        buf .= "<p style=" . q . "margin:0 0 12px 0" . q . ">" . htmlLine . "</p>"
    }
    buf .= "</div>"
    return buf
}

MarkupToHtml(s) {
    ; Escape HTML first so literal text is safe
    s := HtmlEscape(s)

    ; Convert __underline__ first (so _italics_ below doesn't eat the underscores)
    ; __...__  -> <u>...</u>
    s := RegExReplace(s, "s)__(.+?)__", "<u>$1</u>")

    ; Convert **bold** next, but ignore *** (keep triple-asterisks visible)
    ; **...** (not part of ***) -> <strong>...</strong>
    s := RegExReplace(s, "s)(?<!\*)\*\*(.+?)\*\*(?!\*)", "<strong>$1</strong>")

    ; Convert _italic_ last, but ignore __ (double underscore already handled)
    ; _..._ (not part of __) -> <em>...</em>
    s := RegExReplace(s, "s)(?<!_)_(.+?)_(?!_)", "<em>$1</em>")

    ; Preserve any remaining linebreaks within a “line” (rare)
    s := StrReplace(s, "`n", "<br>")
    return s
}

HtmlEscape(s) {
    s := StrReplace(s, "&", "&amp;")
    s := StrReplace(s, "<", "&lt;")
    s := StrReplace(s, ">", "&gt;")
    return s
}

; =========================
;   NEW HELPERS (initially added to satisfy #Warn and runtime)
; =========================
GetOutlook() {
    ; Try to grab or start Outlook, return COM Application or "" on failure.
    try {
        return ComObject("Outlook.Application")
    } catch {
        try {
            Run "outlook.exe"
            Sleep 1500
            return ComObject("Outlook.Application")
        } catch {
            return ""
        }
    }
}

SplitIntoLines(text) {
    if text = ""
        return []
    ; Normalize CRLF to LF before splitting
    t := StrReplace(StrReplace(text, "`r`n", "`n"), "`r", "`n")
    return StrSplit(t, "`n")
}

JoinWithNewlines(arr) {
    out := ""
    for _, v in arr {
        out .= (out ? "`r`n" : "") . v
    }
    return out
}

GetSliceText(lines, start, end) {
    if (start < 1)
        start := 1
    if (end > lines.Length)
        end := lines.Length
    out := ""
    i := start
    while (i <= end) {
        out .= (out ? "`n" : "") . lines[i]
        i++
    }
    return out
}

CreateReplyAllPrepend(olApp, origMail, subj, prependText) {
    try {
        reply := origMail.ReplyAll
    } catch {
        reply := origMail.Reply  ; fallback if ReplyAll not available
    }
    if subj
        reply.Subject := subj

    ; Build an HTML block from the plain-text+markup and prepend it.
    html := RenderHtml(prependText)
    try {
        reply.HTMLBody := html . reply.HTMLBody
    } catch {
        ; Fallback if HTMLBody throws (should be rare)
        reply.Body := RegExReplace(prependText, "\n", "`r`n") . "`r`n`r`n" . reply.Body
    }

    reply.Display()
}

GetPrevGrantText(lines, grantAnyIdx) {
    if (grantAnyIdx <= 0)
        return ""
    rg := GetGrantChunkRange(lines, grantAnyIdx)
    end := rg.end - 1  ; exclude the "-Grant" line itself
    if (end < rg.start)
        end := rg.start
    return GetSliceText(lines, rg.start, end)
}

RemoveQuotedHeaders(txt) {
    cleaned := []
    for ln in SplitIntoLines(txt) {
        t := Trim(ln)
        ; Drop common reply/forward header lines
        if RegExMatch(t, "i)^(from:|sent:|to:|cc:|subject:|attachments|caution:|-----original message-----)")
            continue
        ; Drop day-of-week timestamp lines: "Tuesday, September 16, 2025 9:36 AM"
        if RegExMatch(t, "i)^(mon|tue|wed|thu|fri|sat|sun)[a-z]*,\s+\w+\s+\d{1,2},\s+\d{4}\b")
            continue
        cleaned.Push(ln)
    }
    return JoinWithNewlines(cleaned)
}