; MailFlagCreate.ahk — Create Outlook email flag with custom date and text
; AutoHotkey v2
; Hotkey: Win+Shift+C

#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn

; ---------- USER SETTINGS ----------
DefaultDaysOffset := 7        ; default days in the future
DefaultDueHour    := 9        ; default hour (24-hour format)
; ------------------------------------

; Hotkey: Win+Shift+C
#+c:: CreateFlag()

CreateFlag() {
    ; Check if Outlook is running
    try {
        outlook := ComObjActive("Outlook.Application")
    } catch {
        MsgBox("Outlook is not running. Please start Outlook first.", "MailFlagCreate", "Iconx")
        return
    }

    ; Check if an email is selected
    try {
        explorer := outlook.ActiveExplorer
        if !explorer || !explorer.Selection || explorer.Selection.Count = 0 {
            MsgBox("Please select an email in Outlook first.", "MailFlagCreate", "Iconx")
            return
        }
    } catch {
        MsgBox("Please select an email in Outlook first.", "MailFlagCreate", "Iconx")
        return
    }

    ; Show dialog to get user input
    result := ShowFlagDialog()
    if !result
        return  ; user cancelled

    daysOffset := result.days
    flagText   := result.text

    ; Calculate target date (start with today, add days, set to default hour)
    baseDate := DateAdd(A_Now, daysOffset, "Days")
    
    ; Set to default hour (09:00:00)
    dateStr := FormatTime(baseDate, "yyyyMMdd")
    targetDate := dateStr . Format("{:02}0000", DefaultDueHour)
    
    ; Adjust to next weekday if weekend
    targetDate := NextWeekday(targetDate)

    ; Format date as MM/dd/yyyy for Outlook dialog
    formattedDate := FormatTime(targetDate, "MM/dd/yyyy")
    
    ; Format time as HH:mm AM/PM
    hour := FormatTime(targetDate, "HH")
    minute := FormatTime(targetDate, "mm")
    hour12 := Integer(hour)
    if (hour12 = 0)
        hour12 := 12
    else if (hour12 > 12)
        hour12 -= 12
    ampm := (Integer(hour) < 12) ? "AM" : "PM"
    formattedTime := Format("{:d}:{:02} {}", hour12, minute, ampm)

    ; Activate Outlook window
    WinActivate("ahk_class rctrl_renwnd32")
    Sleep(150)

    ; Open flag dialog
    Send("^+g")
    Sleep(300)

    ; Fill in the flag text
    SendText(flagText)
    Sleep(50)

    ; Tab to date field
    Send("{Tab}")
    Sleep(50)

    ; Clear any existing date and enter new date
    Send("^a")
    Sleep(50)
    SendText(formattedDate)
    Sleep(50)

    ; Tab to time field
    Send("{Tab}")
    Sleep(50)

    ; Clear any existing time and enter new time
    Send("^a")
    Sleep(50)
    SendText(formattedTime)
    Sleep(50)

    ; Tab to move focus away from time field (to OK button or another control)
    ; This allows Enter key to submit the dialog instead of opening date picker
    Send("{Tab}")
    Sleep(50)
    Send("{Tab}")
    Sleep(50)
    Send("{Tab}")
    Sleep(50)
    Send("{Tab}")
    Sleep(50)
    Send("{Tab}")
    Sleep(50)
    Send("{Tab}")
    Sleep(50)

    ; Dialog is now ready for user to review and press Enter
    ; Focus should now be on OK button so Enter will submit the flag
}

; ---------- Dialog GUI ----------
ShowFlagDialog() {
    submitted := false
    daysValue := DefaultDaysOffset
    textValue := ""
    resultDays := 0
    resultText := ""

    g := Gui("+AlwaysOnTop +ToolWindow", "Create Email Flag")
    g.MarginX := 20
    g.MarginY := 15

    g.Add("Text", "w300", "Days in the future:")
    daysEdit := g.Add("Edit", "w300 Number vDaysField", daysValue)

    ; Date display label (above flag text field)
    dateLabel := g.Add("Text", "w300 y+10", "Flag will be created for: " . CalculateFlagDate(daysValue))

    g.Add("Text", "w300 y+15", "Flag text:")
    textEdit := g.Add("Edit", "w300 vTextField", textValue)

    btnOK := g.Add("Button", "w80 Default xp y+20", "OK")
    btnCancel := g.Add("Button", "w80 x+10", "Cancel")

    ; Update date display when days value changes
    UpdateDateDisplay(*) {
        daysStr := daysEdit.Value
        if (daysStr != "") {
            daysInt := Integer(daysStr)
            if (daysInt >= 0) {
                calculatedDate := CalculateFlagDate(daysInt)
                dateLabel.Text := "Flag will be created for: " . calculatedDate
            }
        }
    }
    daysEdit.OnEvent("Change", UpdateDateDisplay)

    btnOK.OnEvent("Click", OnOK)
    btnCancel.OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Close", (*) => g.Destroy())

    OnOK(*) {
        daysStr := daysEdit.Value
        textStr := textEdit.Value

        if (daysStr = "" || textStr = "") {
            MsgBox("Both fields are required.", "MailFlagCreate", "Iconx")
            return  ; Don't close dialog, let user fix it
        }

        daysInt := Integer(daysStr)
        if (daysInt < 0) {
            MsgBox("Days must be 0 or greater.", "MailFlagCreate", "Iconx")
            return  ; Don't close dialog, let user fix it
        }

        ; Store values before destroying dialog
        resultDays := daysInt
        resultText := textStr
        submitted := true
        g.Destroy()
    }

    ; Show dialog
    g.Show()

    ; Focus on days field and select all
    daysEdit.Focus()
    Sleep(100)
    Send("^a")

    ; Wait for dialog to close
    WinWaitClose(g)

    if !submitted
        return 0

    return {days: resultDays, text: resultText}
}

; ---------- Helper: Calculate Flag Date ----------
CalculateFlagDate(daysOffset) {
    ; Calculate target date using the same logic as CreateFlag()
    ; Returns formatted date string for display
    
    ; Calculate target date (start with today, add days, set to default hour)
    baseDate := DateAdd(A_Now, daysOffset, "Days")
    
    ; Set to default hour (09:00:00)
    dateStr := FormatTime(baseDate, "yyyyMMdd")
    targetDate := dateStr . Format("{:02}0000", DefaultDueHour)
    
    ; Adjust to next weekday if weekend
    targetDate := NextWeekday(targetDate)
    
    ; Format date as "Month Day, Year (DayOfWeek)" (e.g., "January 15, 2024 (Monday)")
    formattedDate := FormatTime(targetDate, "MMMM d, yyyy")
    dayOfWeek := FormatTime(targetDate, "dddd")
    
    return formattedDate . " (" . dayOfWeek . ")"
}

; ---------- Helper: Next Weekday ----------
NextWeekday(ts) {
    ; ts = YYYYMMDDHH24MISS format
    ; Returns timestamp adjusted to next weekday if weekend
    
    wd := FormatTime(ts, "ddd")  ; "Mon", "Tue", ... "Sun"
    
    if (wd = "Sat")
        return DateAdd(ts, 2, "Days")  ; Saturday → Monday (+2 days)
    
    if (wd = "Sun")
        return DateAdd(ts, 1, "Days")  ; Sunday → Monday (+1 day)
    
    return ts  ; Already a weekday
}

