; <COMPILER: v1.1.37.01>

#SingleInstance force
#MaxThreads 1
SetTitleMatchMode, 2
CoordMode, Mouse, screen
CoordMode, Pixel, screen

if !A_IsAdmin {
    try {
        if A_IsCompiled
            Run *RunAs "%A_ScriptFullPath%" /force
        else
            Run *RunAs "%A_AhkPath%" /force "%A_ScriptFullPath%"
    }
    catch {
        MsgBox, 262208, 오류, 관리자 권한 승인이 거부되었습니다. 프로그램이 종료됩니다.
    }
    ExitApp ; 권한을 요청한 '현재' 프로세스는 반드시 종료해야 합니다. (새 프로세스가 뜰 것이므로)
}

#include Vars.ahk
#Include Gui.ahk
#Include Functions.ahk

LoadSettings()
LoadQuickSlot(chooseSlotNum)

Gui, Show, x0 y0 w345 h264, Amore Transportation Helper
return

#Include Hotkeys.ahk
#Include Labels.ahk