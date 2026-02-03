; <COMPILER: v1.1.37.01>

#SingleInstance force
SetTitleMatchMode, 2
CoordMode, Mouse, screen
CoordMode, Pixel, screen
;SetNumberHotkeys("off")
;SetNumpadHotkeys("off")
;Hotkey,Ctrl,off

#include Vars.ahk
#Include Gui.ahk
#Include Functions.ahk


; 설정 값 불러오기
LoadSettings()
; 퀵슬롯 차량 불러오기
LoadQuickSlot(chooseSlotNum)

Gui, Show, x0 y0 w345 h264, Amore Transportation Helper
return

ResetStatus:
	GuiControl, , Status,대기 중..
	return


GuiClose:
	ExitApp

ExplainBtn:
	MsgBox,[F1] - 차량 출입 현황 재필터링, 차량 입출조회 재조회.`n[F3] - 통합차량 조회 -> 차량번호로 찾기.`n[F4] - 현재 선택된 행의 차량정보를 일일차량현황 빈칸에 입력.`n[F6] - 납품 <-> 반출 전환.`n[F7] - 납품 -> 반출로 빈칸에 입력. `n[INS] - 현재 선택된 행의 차량번호를 TMS 등록창에 입력.`n[CTRL]+[B,T] - "BTOS","TRDT" 입력.`n[NUMLK] - "/전산" 입력.`n[CTRL]+[TAB] - 현재 시간 입력.`n[CTRL]+[1~7] - 핫키에 저장된 차량정보 빈칸에 입력.`n[ALT]+[1~7] - 선택한 행의 차량정보 핫키에 등록.`n[F11] - TMS, 차량현황 엑셀을 프로그램에 등록.`n[SCRLK] - 클릭.`n`n`n`n* 처음 실행 후 차량 조회, 등록 TMS와`n오늘 날짜의 일일차량 엑셀 파일을 실행 하고 등록을 진행하여 주세요.`n`n* 사용 중 명령에 이상이 있을땐`n대기시간을 조정해 속도를 느리게 설정해주세요.`n`n* 작동이 아예 멈추거나 문제가 생겼을땐`n껐다가 키고 다시 등록해 사용해주세요.
	return

HotkeySettingBtn:
	Gui, 2: Show, x330 y0 w935 h264, 핫키 설정
	return

OtherSettingBtn:
	LoadSettings()
	Gui, 3: Show, x330 y0 w570 h264, 기타 설정
	return

RegisterBtn:
	RecordLog("RegisterBtn Pressed")
	RegistPrograms()
	return


Slot1:
	HandleSlotSelect(1)
	return
Slot2:
	HandleSlotSelect(2)
	return
Slot3:
	HandleSlotSelect(3)
	return
Slot4:
	HandleSlotSelect(4)
	return
Slot5:
	HandleSlotSelect(5)
	return
Slot6:
	HandleSlotSelect(6)
	return

SettingSlot1:
	HandleSettingSlotSelect(1)
	return
SettingSlot2:
	HandleSettingSlotSelect(2)
	return
SettingSlot3:
	HandleSettingSlotSelect(3)
	return
SettingSlot4:
	HandleSettingSlotSelect(4)
	return
SettingSlot5:
	HandleSettingSlotSelect(5)
	return
SettingSlot6:
	HandleSettingSlotSelect(6)
    return

SettingWriteBtn:
	SaveQuickSlot(chooseSlotNum)
    return

OtherWriteBtn:
    SaveSettings()
    return


#Include Hotkeys.ahk

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



SetNumberHotkeys(state)
{
    Loop, 10
    {
        key := (A_Index = 10) ? 0 : A_Index
        Hotkey, %key%, %state%
    }
}

SetNumpadHotkeys(state)
{
    Loop, 10
    {
        key := "Numpad" . (A_Index - 1)
        Hotkey, %key%, %state%
    }
}

HandleCarInput(idx)
{
    global excelName, xl, chooseSlotNum, assistantTool1

    if(!CheckExcel()) {
        return
    }

    WinActivate, %excelName%
    WinWaitActive, %excelName%, , 1

    ;SetNumberHotkeys("On")
    ;Hotkey, Ctrl, On

    try
    {
        ; 1. INI에서 해당 슬롯의 탭 문자열을 그대로 읽어옴
        IniRead, savedLine, assistantTool1, slot%chooseSlotNum%, %idx%

        if (savedLine = "ERROR" || savedLine = "") {
            MsgBox, 262208, 알림, 해당 슬롯에 데이터가 없습니다.
            goto Cleanup
        }

        finalLine := ReformCarInfo(savedLine, true)

        ; 3. 엑셀 작업 시작
        xl.Sheets(1).Select
        WaitExcel()

        ; 빈 행 찾기
        targetRow := FindLastRow()

        Clipboard := finalLine
        ClipWait, 1

        WinActivate, %excelName%
        WinWaitActive, %excelName%, , 1

        xl.Range("C" . targetRow).Select
        WaitExcel()
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
        goto Cleanup
    }

Cleanup:
    Send, {Ctrl up}
    ;SetNumberHotkeys("Off")
    ;Hotkey, Ctrl, Off
}


RegisterSlotFromExcel(idx)
{
    global excelName, xl, chooseSlotNum, assistantTool1

    if(!CheckExcel(true, "ALT")) {
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

        dataArr := StrSplit(newDataLine, A_Tab)
        if (dataArr[1] = "" || dataArr[1] = "차량번호")
        {
            MsgBox, 262208, 알림, 차량정보가 올바르지 않습니다.
            return
        }

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


