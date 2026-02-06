; <COMPILER: v1.1.37.01>

#SingleInstance force
#MaxThreads 1
SetTitleMatchMode, 2
CoordMode, Mouse, screen
CoordMode, Pixel, screen

#include Vars.ahk
#Include Gui.ahk
#Include Functions.ahk

LoadSettings()
LoadQuickSlot(chooseSlotNum)

Gui, Show, x0 y0 w345 h264, Amore Transportation Helper
return

#Include Hotkeys.ahk
#Include Labels.ahk

;=======================================================================================================================
HandleSlotSelect(slotNum)
{
    global chooseSlotNum

    chooseSlotNum := slotNum
    IniWrite, %slotNum%, assistantTool1, settings, chooseSlotNum
    LoadQuickSlot(slotNum)

    GuiControl, 2:, SettingSlot%slotNum%, 1
    GuiControl, , Status, %slotNum%번 슬롯 선택 완료.

    SetTimer, ResetStatus, 3000
}

HandleSettingSlotSelect(slotNum)
{
    global chooseSlotNum

    chooseSlotNum := slotNum
    IniWrite, %slotNum%, assistantTool1, settings, chooseSlotNum
    LoadQuickSlot(slotNum)

    GuiControl, 1:, Slot%slotNum%, 1
    GuiControl, 1:, Status, %slotNum%번 슬롯 설정 선택 완료

    SetTimer, ResetStatus, 3000
}


HandleCarInput(idx)
{
    global excelName, xl, chooseSlotNum, assistantTool1

    if(!CheckExcel()) {
        return
    }

    ActivateWindow(excelName)
    ExcelOptimizer(true)

    try
    {
        ; 1. INI에서 해당 슬롯의 탭 문자열을 그대로 읽어옴
        IniRead, savedLine, assistantTool1, slot%chooseSlotNum%, %idx%

        if (savedLine = "ERROR" || savedLine = "") {
            MsgBox, 262208, 알림, 해당 슬롯에 데이터가 없습니다.
            return
        }
        row := StrSplit(savedLine, A_Tab)

        if (Trim(row[1]) = "") {
            MsgBox, 262208, 알림, 해당 슬롯에 차량번호가 없습니다.
            return
        }

        finalLine := ReformCarInfo(row, true)

        MoveSheet(1)

        ; 빈 행 찾기
        targetRow := FindLastRow()

        Clipboard := finalLine
        ClipWait, 1

        if (midOffset > 1 && inputscroll = 1) {
            xl.ActiveWindow.ScrollRow := Max(1, targetRow - midOffset)
        }

        xl.Range("C" . targetRow).Select
        WaitExcel()
        ExcelOptimizer(false)
        Send, ^v{Ctrl Up}
        WaitExcel()

        ;xl.Range("O" . targetRow).NumberFormat := "HH:mm;@" ; 서식 지정
        lastVal := xl.Cells(targetRow, 17).Value ; Q열 (카드/전산 정보)

        if (lastVal = "카드/전산" || lastVal = "48/전산" || lastVal = "50/전산") {
            xl.Range("K" . targetRow).Select
        }
        else {
            xl.Range("Q" . targetRow).Select
        }
    }
    catch e
    {
        RecordLog("^" idx " - 실패: " e.message)
        ExcelOptimizer(false)
        return
    }
}


RegisterSlotFromExcel(idx)
{
    global excelName, xl, chooseSlotNum, assistantTool1

    if(!CheckExcel(true, "ALT")) {
        return
    }
    if(!CarExist()){
        return
    }

    try
    {
        Send, ^{Enter}{Ctrl up}
        WaitExcel()

        ; 1. 엑셀 C열부터 Q열까지 한 번에 복사
        selectionRow := xl.Selection.Row
        Clipboard := ""
        xl.Range("C" . selectionRow . ":Q" . selectionRow).Copy
        ClipWait, 1

        newDataLine := ReformCarInfo(Clipboard, false)

        IniDelete, assistantTool1, slot%chooseSlotNum%, %idx%
        IniWrite, %newDataLine%, assistantTool1, slot%chooseSlotNum%, %idx%

        LoadQuickSlot(chooseSlotNum)
        xl.Application.CutCopyMode := False

        GuiControl, 1:, Status, % chooseSlotNum "번 슬롯 " idx "번 교체 완료"
        SetTimer, ResetStatus, 3000
    }
    catch
    {
        RecordLog("!" idx " - 실패")
        ExitApp
    }
}


