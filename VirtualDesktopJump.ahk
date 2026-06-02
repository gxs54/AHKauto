; VirtualDesktopJump.ahk — AutoHotkey v2.x  (2026-06-02)
;
; Jump directly to virtual desktop 1–9 with  Ctrl + Win + <number>.
;
; Why Ctrl+Win and not plain Win+<number>:
;   Win+1..9 is the built-in Windows shortcut to launch/focus pinned taskbar
;   apps. Ctrl+Win is already the OS modifier for virtual desktops
;   (Ctrl+Win+Left/Right), so this layers on cleanly without hijacking the
;   taskbar launchers.
;
; How it works (no external DLL, robust on Win10/Win11):
;   Windows stores the ordered list of virtual-desktop GUIDs and the current
;   desktop GUID in the registry. We read those, work out how many hops left
;   or right are needed, and send Ctrl+Win+Left/Right that many times. This is
;   deterministic — it knows where it's starting from — unlike older blind
;   keystroke-spam scripts.

#Requires AutoHotkey v2.0
#SingleInstance Force

;───────── USER SETTINGS ─────────
modifiers     := "^#"     ; ^=Ctrl  #=Win   (prefix for the 1–9 hotkeys)
switchDelayMs := 10       ; pause between consecutive desktop hops. With the
                          ; slide animation off (below) each hop is instant, so
                          ; this is tiny — just enough for the shell to register
                          ; each keypress. Bump it if a multi-hop ever misfires.

; The keystroke method walks desktop-by-desktop, so to make a multi-desktop
; jump look instant we disable Windows' desktop-switch slide animation. We set
; it globally once at install; this re-asserts it on launch so the jump stays
; clean even if something flips it back. NOTE: this is a system-wide setting
; ("Animation effects" in Settings ▸ Accessibility ▸ Visual effects) — it also
; removes animation from manual Ctrl+Win+←/→ and a few other UI transitions.
; If you ever want those animations back, set this to false AND re-enable
; "Animation effects" in Settings (otherwise this line turns it off again).
enforceNoAnimation := true
;─────────────────────────────────

if (enforceNoAnimation)
    DllCall("SystemParametersInfoW", "UInt", 0x1043, "UInt", 0, "Ptr", 0, "UInt", 0x3)
    ; SPI_SETCLIENTAREAANIMATION, pvParam=FALSE → disable; SPIF_UPDATEINIFILE|SPIF_SENDWININICHANGE

regBase := "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops"

; Bind Ctrl+Win+1 … Ctrl+Win+9
Loop 9
    Hotkey(modifiers . A_Index, JumpHotkey)

JumpHotkey(*) {
    ; A_ThisHotkey looks like "^#5" — the desktop number is the last character.
    SwitchToDesktop(Integer(SubStr(A_ThisHotkey, -1)))
}

SwitchToDesktop(target) {
    global regBase, modifiers, switchDelayMs

    idsHex := ""
    try idsHex := RegRead(regBase, "VirtualDesktopIDs")
    if (idsHex = "")                 ; can't read the desktop list → bail safely
        return

    chunk := 32                      ; one GUID = 16 bytes = 32 hex chars
    count := StrLen(idsHex) // chunk
    if (target < 1 || target > count) ; desktop doesn't exist → do nothing
        return

    curHex := ReadCurrentDesktopGuid()
    if (curHex = "")
        return

    cur := -1
    Loop count {
        if (SubStr(idsHex, (A_Index - 1) * chunk + 1, chunk) = curHex) {
            cur := A_Index - 1       ; 0-based index of the current desktop
            break
        }
    }
    if (cur = -1)                    ; current GUID not in the list → bail safely
        return

    delta := (target - 1) - cur
    if (delta = 0)                   ; already there
        return

    arrow := (delta > 0) ? "{Right}" : "{Left}"
    hops  := Abs(delta)
    Loop hops {
        Send(modifiers . arrow)
        if (A_Index < hops)
            Sleep(switchDelayMs)
    }
}

; The current-desktop GUID normally lives at the top-level VirtualDesktops key.
; On some Windows builds it only appears under a per-session subkey, so fall
; back to scanning SessionInfo if the top-level value is missing.
ReadCurrentDesktopGuid() {
    global regBase

    cur := ""
    try cur := RegRead(regBase, "CurrentVirtualDesktop")
    if (cur != "")
        return cur

    sessRoot := "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\SessionInfo"
    Loop Reg sessRoot, "K" {         ; enumerate session subkeys
        try {
            v := RegRead(sessRoot "\" A_LoopRegName "\VirtualDesktops", "CurrentVirtualDesktop")
            if (v != "")
                return v
        }
    }
    return ""
}
