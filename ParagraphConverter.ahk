; ==============================
; Word Template Paragraph Replacer (AutoHotkey v2)
; Hotkey: Ctrl+Shift+T
; Storage: In-script AHK array (no external parsers)
; ==============================
#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn

; Allow a second thread ONLY so a press arriving while the template GUI or a
; placeholder prompt is open reaches the handler and gets told so. At the
; default of 1 the press is discarded before any script code runs and the
; hotkey goes silently dead. The busy flag below still serializes the work.
#MaxThreadsPerHotkey 2

; -------- CONFIG --------
placeholderRx := "{{(.*?)}}"                  ; placeholder pattern
logFile       := A_ScriptDir "\TemplateReplace.log"

; -------- TEMPLATE DB --------
; Add templates here
global templates := [
    Map(
        "id",    "examiner_interview_intro",
        "title", "Report Letter - Examiner Interview - Intro paragraph",
        "body",  "As instructed in your {{ClientLetterDate}} letter, we performed an Examiner Interview to discuss the claim amendments and arguments included in your letter. We again note that the deadline for responding to the Office Action is {{OAResponseDate}}, although this deadline is extendable for up to three months if necessary with a payment of an additional fee.",
        "placeholders", [
            Map("name","ClientLetterDate","prompt","Enter the client's letter date (e.g., July 10, 2025)","type","date"),
            Map("name","OAResponseDate","prompt","Office Action response deadline","type","date")
        ]
    )
]

; -------- HOTKEY --------
^+t:: HandleHotkey()
return

HandleHotkey() {
    static busy := false
    if busy {
        TrayTip("A template dialog is already open.", "Template Replacer")
        return
    }
    busy := true
    try
        HandleTemplateReplace()
    finally
        busy := false
}

; =======================================
;               MAIN
; =======================================
HandleTemplateReplace() {
    global templates, placeholderRx
    try {
        selText := GetWordSelection()
        if !selText {
            MsgBox "No text selected in Word.", "Template Replacer", 48
            return
        }
        if (templates.Length = 0) {
            MsgBox "No templates loaded.", "Template Replacer", 48
            return
        }
        chosen := ShowTemplateSearchGUI(templates)
        if !IsObject(chosen)
            return  ; cancelled

        ; pass selected text so we can auto-detect ONLY whitelisted dates
        filled := FillPlaceholders(chosen["body"], chosen["placeholders"], selText)
        if (filled = "")
            return  ; user aborted

        ReplaceWordSelection(filled)
        LogAction("Replaced selection with template: " chosen["id"])
    } catch Error as e {
        MsgBox "Error: " e.Message "`n" e.What, "Template Replacer", 16
        LogAction("ERROR: " e.Message)
    }
}

; =======================================
;           WORD HELPERS
; =======================================
GetWordSelection() {
    word := ComObjActive("Word.Application")
    return word.Selection.Text
}
ReplaceWordSelection(text) {
    word := ComObjActive("Word.Application")
    word.Selection.Text := text
}

; =======================================
;          TEMPLATE SEARCH GUI
; =======================================
ShowTemplateSearchGUI(templates) {
    local g, edt, lv, btnOK, btnCancel, row, chosenId, t
    local submitted := false, closed := false

    g := Gui("+AlwaysOnTop", "Search Templates")
    g.MarginX := 10, g.MarginY := 10

    g.Add("Text",, "Search:")
    edt := g.Add("Edit", "w420 vSearch")

    lv := g.Add("ListView", "w420 r12 vLV AltSubmit", ["Title", "ID"])
    for t in templates
        lv.Add("", t["title"], t["id"])
    lv.ModifyCol(1, 280), lv.ModifyCol(2, 120)

    btnOK     := g.AddButton("Default w80", "OK")
    btnCancel := g.AddButton("w80", "Cancel")

    edt.OnEvent("Change", (ctrl,*) => FilterLV(ctrl.Text, lv, templates))
    lv.OnEvent("DoubleClick", (*) => btnOK.Click())

    btnOK.OnEvent("Click", (*) => (
        row := lv.GetNext(),
        chosenId := (row ? lv.GetText(row, 2) : ""),
        submitted := (row != 0),
        closed := true,
        g.Destroy()
    ))
    btnCancel.OnEvent("Click", (*) => (submitted := false, closed := true, g.Destroy()))
    ; Must Destroy, and must not return a truthy value: a Close handler whose
    ; result is non-zero cancels the default close, so clicking the X used to
    ; leave this always-on-top window stranded on screen after the script had
    ; already moved on. Verified 2026-08-18 — WinClose would not shift it.
    g.OnEvent("Close", (*) => (submitted := false, closed := true, g.Destroy(), 0))

    g.Show()
    while !closed
        Sleep 50

    if !submitted || chosenId = ""
        return 0

    for t in templates
        if (t["id"] = chosenId)
            return t
    return 0
}

FilterLV(searchText, lv, templates) {
    lv.Delete()
    if !searchText {
        for t in templates
            lv.Add("", t["title"], t["id"])
        return
    }
    st := StrLower(searchText)
    for t in templates
        if InStr(StrLower(t["title"]), st) || InStr(StrLower(t["body"]), st)
            lv.Add("", t["title"], t["id"])
}

; =======================================
;           PLACEHOLDER FILL
; =======================================
FillPlaceholders(body, placeholders, selText := "") {
    global placeholderRx
    ; Only these names will be auto-filled from the selected paragraph
    static autoDateNames := ["OAResponseDate", "OfficeActionDeadline", "ResponseDeadline"]

    local needed := [], answers := Map()
    local pos := 1, match, name, prompt, typ, p, val, out

    ; collect placeholders
    while RegExMatch(body, placeholderRx, &match, pos) {
        name := match[1]
        if !ArrHas(needed, name)
            needed.Push(name)
        pos := match.Pos + match.Len
    }

    ; prompt / auto-fill
    for name in needed {
        prompt := "Enter value for " name
        typ    := "text"

        if IsObject(placeholders) {
            for p in placeholders {
                if (p["name"] = name) {
                    if p.Has("prompt")
                        prompt := p["prompt"]
                    if p.Has("type")
                        typ := p["type"]
                    break
                }
            }
        }

        auto := ""
        if (typ = "date" && selText != "" && ArrHas(autoDateNames, name)) {
            auto := ExtractDateFromText(selText)
        }

        val := (auto != "") ? auto : PromptForValue(prompt, typ)
        if (val = "") {
            if MsgBox("You left a field blank. Cancel replacement?", "", 36) = "Yes"
                return ""
        }
        answers[name] := val
    }

    out := body
    for name, val in answers
        out := StrReplace(out, "{{" name "}}", val)
    return out
}

ArrHas(arr, val) {
    for v in arr
        if (v = val)
            return true
    return false
}

ExtractDateFromText(text) {
    local m
    ; Prefer dates near 'deadline ... is/on/by'
    if RegExMatch(
        text
      , "i)deadline[^`n]{0,200}?\b(?:is|on|by)\b\s*((January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4})"
      , &m)
        return m[1]

    ; Fallback: any Month DD, YYYY
    if RegExMatch(
        text
      , "i)(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}"
      , &m)
        return m[0]

    return ""
}

PromptForValue(prompt, typ := "text") {
    if (typ = "date") {
        res := TopmostInputBox(prompt, "Date Input", "w360 h140")
        if res.Result = "Cancel"
            return ""
        return res.Value
    }
    res := TopmostInputBox(prompt, "Input", "w420 h140")
    if res.Result = "Cancel"
        return ""
    return res.Value
}

; AutoHotkey's InputBox has no always-on-top option, so a placeholder prompt
; can end up behind Word — and while it sits there unnoticed the hotkey does
; nothing at all. A template with several placeholders opens one of these per
; placeholder, so raise each as soon as it exists.
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

; =======================================
;               LOGGING
; =======================================
LogAction(msg) {
    global logFile
    FileAppend(FormatTime(, "yyyy-MM-dd HH:mm:ss") " - " msg "`r`n", logFile, "UTF-8")
}
