; ==========================================================
; FUNCTIONS
; ==========================================================
;===========================================================
;세팅 정보를 불러오는 함수
;===========================================================
LoadSettings() {
    global ;

    ; 일반 설정 로드
    ReadToVar("settings", "searchStartRow", 7)
    ReadToVar("settings", "chooseSlotNum", 1)

    ReadToVar("settings", "autoslip", 1, "radio")
    ReadToVar("settings", "searchto", 1, "radio")
    ReadToVar("settings", "onlyexcel", 1, "radio")

    ; 슬롯 라디오 버튼 체크
    if (chooseSlotNum != "") {
        GuiControl, 1:, Slot%chooseSlotNum%, 1
        GuiControl, 2:, SettingSlot%chooseSlotNum%, 1
    }
}

SaveSettings() {
    global
	Gui, 3: Submit, NoHide

    issearchOk := (searchStartRow >= 7 && searchStartRow <= 500000)


    if (issearchOk)
    {
        GuiControl, 1:, Status, 기타 설정 저장 중

        IniWrite, %searchStartRow%, %iniPath%, settings, searchStartRow
        IniWrite, %searchStartRow%, %iniPath%, settings, onlyexcel

        Gui, 3: Hide
        MsgBox, 262144, 알림, 저장되었습니다.
        GuiControl, 1:, Status, 기타 설정 저장 완료
        SetTimer, ResetStatus, 3000
    }
    else
    {
        MsgBox, 262208, 알림, 검색 시작 행 설정 값이 범위를 초과하였습니다. (7~500000)
    }

    autoslip := autoslip1 ? 1 : (autoslip2 ? 2 : 1)
    searchto := searchto1 ? 1 : (searchto2 ? 2 : 1)
    onlyexcel := onlyexcel1 ? 1 : (onlyexcel2 ? 2 : 1)

    IniWrite, %autoslip%, %iniPath%, settings, autoslip
    IniWrite, %searchto%, %iniPath%, settings, searchto
    IniWrite, %onlyexcel%, %iniPath%, settings, onlyexcel
	return
}

;===========================================================
;INI파일의 정보를 읽어오는 공용 함수
;===========================================================
ReadToVar(Section, Key, DefaultValue := "", Type := "") {
    global

    IniRead, readTemp, %iniPath%, %Section%, %Key%, %DefaultValue%
    %Key% := readTemp
    if (Type = "radio") {
        GuiControl, 3:, %Key%%readTemp%, 1
    } else {
        GuiControl, 3:, %Key%, %readTemp%
    }
}


;===========================================================
;현재 선택된 슬롯에 차량 목록을 저장하는 함수
;===========================================================
SaveQuickSlot(chooseSlotNum)
{
    if (chooseSlotNum = 0) {
        MsgBox, 262208, 알림, 슬롯이 선택되지 않았습니다.
        return
    }

    GuiControl, 1:, Status, %chooseSlotNum%번 슬롯 저장 중
    Gui, 2: Submit, NoHide

    IniDelete, %iniPath%, slot%chooseSlotNum%

    Loop, 7
    {
        cIdx := A_Index

        ; GUI의 v옵션 이름인 Car1NumEdit 등과 매칭되도록 구성
        dataLine := Car%cIdx%NumEdit     . A_Tab
                 .  Car%cIdx%NameEdit    . A_Tab
                 .  Car%cIdx%CompanyEdit . A_Tab
                 .  Car%cIdx%PhoneEdit   . A_Tab
                 .  Car%cIdx%ContentEdit . A_Tab
                 .  Car%cIdx%CarryEdit   . A_Tab
                 .  Car%cIdx%DropEdit    . A_Tab
                 .  "/"                  . A_Tab ; 8번 고정

        Loop, 6 ; 9~14번 빈칸
            dataLine .= A_Tab

        ; 15번 항목
        dataLine .= Car%cIdx%CardEdit
        safeData := """" . dataLine . """"

        ; INI 파일에는 키 이름을 숫자(1, 2, 3...)로 저장하여 깔끔하게 관리
        IniDelete, %iniPath%, slot%chooseSlotNum%, %cIdx%
        IniWrite, %safeData%, %iniPath%, slot%chooseSlotNum%, %cIdx%
    }

    IniWrite, %chooseSlotNum%, %iniPath%, settings, chooseSlotNum
    LoadQuickSlot(chooseSlotNum)

    Gui, 2: Hide
    MsgBox, 262144, 알림, %chooseSlotNum%번 슬롯 저장 완료
    GuiControl, 1:, Status, %chooseSlotNum%번 슬롯 저장 완료
    SetTimer, ResetStatus, 3000
    return
}

;===========================================================
; 퀵슬롯 차량목록을 불러오는 함수
;===========================================================
LoadQuickSlot(slotNum) {

    global

    fields := ["Num", "Name", "Company", "Phone", "Content", "Carry", "Drop", "Card"]

    Loop, 7 {
        cIdx := A_Index
        IniRead, rawLine, %iniPath%, slot%slotNum%, %cIdx%, %A_Space%
        Car%cIdx%Data := rawLine
        rawLine := Trim(rawLine, """")
        row := StrSplit(rawLine, A_Tab)

        for fIdx, fName in fields {
            val := (fName = "Card") ? row[15] : row[fIdx]
            ctrlName := "Car" . cIdx . fName
            GuiControl, 2:, % ctrlName . "Edit", %val%

            if (fName = "Num" || fName = "Name") {
                GuiControl, 1:, % ctrlName . "View", %val%
            }
        }
    }
}


;===========================================================
;TMS와 엑셀을 등록하는 함수
;===========================================================
RegistPrograms()
{
    global

    if(onlyexcel = 1){
        IfWinNotExist, [WMSDB] AMOREPACIFIC Transportation Management System
        {
            RecordLog("TMS 미실행 등록시도")
            MsgBox, 262208, 알림, TMS 프로그램이 실행되어지지 않았습니다.`n실행 후 세팅을 진행하여 주세요.
            return
        }

        GroupAdd, tmsWindows, [WMSDB] AMOREPACIFIC Transportation Management System
        GroupActivate, tmsWindows
        WinGetText, textCheck, [WMSDB] AMOREPACIFIC Transportation Management System

        if (InStr(textCheck, "차량입출조회") > 0)
        {
            WinGet, tms1Pid, PID, [WMSDB] AMOREPACIFIC Transportation Management System
            GroupActivate, tmsWindows
            WinGet, tms2Pid, PID, [WMSDB] AMOREPACIFIC Transportation Management System
        }
        else
        {
            WinGet, tms2Pid, PID, [WMSDB] AMOREPACIFIC Transportation Management System
            GroupActivate, tmsWindows
            WinGet, tms1Pid, PID, [WMSDB] AMOREPACIFIC Transportation Management System
        }
}



    FormatTime, excelName,, yyyy년 MM월 dd일 일일 차량현황
    if (!CheckExcel())
    {
        return
    }

    try
    {
        try {
            xl := ComObjActive("Excel.Application")
        } catch {
            RecordLog("Excel COM 객체 연결 실패")
            MsgBox, 262208, 오류, 엑셀과 연결할 수 없습니다. 엑셀을 재시작 하거나 관리자권한으로 실행하세요.
            xl := ""
        return false
        }
        WaitExcel()

        GuiControl, , Status, 프로그램 등록 중..

        ; 창 위치 및 레이아웃 조정
        ActivateWindow(excelName)
        WinMove, ahk_pid %tms1Pid%, , 0, 291
        WinMove, ahk_pid %tms2Pid%, , 895, 291
        WinMove, %excelName%, , 1912, -8

        Gui, Show, x0 y0 w345 h264, Amore Transportation Helper
        WinMaximize, %excelName%
        ActivateWindow(excelName)

    }
    catch
    {
        RecordLog("등록 실패")
        return
    }
    GuiControl, , Status, 프로그램 등록 완료.
    SetTimer, ResetStatus, 3000
    return
}

;===========================================================
; 로그 생성
;===========================================================
RecordLog(sentence)
{
	FileAppend,[%A_MM%-%A_DD% / %A_Hour%:%A_Min%:%A_Sec%] - [%sentence%]`n, ATH_Log.txt
}


;===========================================================
; TMS 체크
;===========================================================
CheckTMS(alert := false, callerName := "", checkActive := false) {
    global

    ; 엑셀만 사용상태
    if(onlyexcel != 1){
        if(alert) { ;알림띄우고 중단해야 할 경우
            MsgBox, 262208, 알림, 엑셀만 사용하도록 설정되어있습니다.`n기타 설정에서 TMS 사용으로 변경하시고`n[F11]키를 눌러 재등록 후 다시 시도해 주세요.
            return false
        }
        return true ; 아니면 통과
    } else { ; TMS사용설정이라 체크 함

        if (tms1Pid = "" || tms2Pid = "") {
            RecordLog("TMS 미등록 시도")
            MsgBox, 262208, 알림, TMS 조회창의 PID가 적용되지 않았습니다.
            return false ; 실패 신호
        }

        if (!WinExist("ahk_pid " . tms1Pid) || !WinExist("ahk_pid " . tms2Pid)) {
            RecordLog("TMS pid 에러")
            MsgBox, 262208, 알림, TMS 조회창을 찾을 수 없습니다.
            return false
        }

        ;활성화 체크루트
        if(checkActive && !WinActive("ahk_pid " . tms2Pid)){
            msg := ""
            switch callerName
            {
                case "Insert":
                    msg := "TMS 등록창이 활성화 되지 않은 채 [Insert]키를 눌렀습니다.`nTMS 등록창을 클릭해 활성화 한 후 다시 시도해주세요."
                case "ScrollLock":
                    msg := "TMS 등록창이 활성화 되지 않은 채 [ScrollLock]키를 눌렀습니다.`nTMS 등록창을 클릭해 활성화 한 후 다시 시도해주세요."
                case "Pause":
                    msg := "TMS 등록창이 활성화 되지 않은 채 [Pause]키를 눌렀습니다.`nTMS 등록창을 클릭해 활성화 한 후 다시 시도해주세요."
                default:
                    msg := "TMS 등록창이 활성화 되지 않았습니다."
            }
            MsgBox, 262208, 알림, % msg
            return false
        }
    }
    return true
}

;===========================================================
; 엑셀 체크
;===========================================================
CheckExcel(checkActive := false, callerName := "") {  ; 기본값은 false (체크 안 함)
    global

    Process, Exist, Excel.exe
    if (!ErrorLevel)
    {
        RecordLog(callerName . " - 엑셀 프로세스 없음")
        MsgBox, 262208, 알림, 엑셀 프로그램이 실행되어 있지 않습니다.`n엑셀을 먼저 실행해 주세요.
        return false
    }

    ; 1. 엑셀 창 존재 여부 및 파일명 등록 확인
    if !WinExist(excelName)
    {
        if (excelName = "")
        {
            RecordLog("엑셀 미등록 시도")
            MsgBox, 262208, 알림, 일일 차량현황 엑셀 파일이 적용되지 않았습니다.`nF11을 눌러 세팅을 진행하여 주세요.
        }
        else
        {
            RecordLog("엑셀 상이 시도")
            MsgBox, 262208, 알림, 일일 차량현황 엑셀 파일의 제목이`n현재날짜와 다르거나 실행 되지 않았습니다.`n현재 날짜의 일일 차량현황 엑셀파일을 실행하여 주세요.
        }
        return false
    }

    ; 2. 엑셀 창이 활성화(맨 위) 되어 있는지 확인, false 시 체크하지 않음
    if (checkActive && !WinActive(excelName))
    {
        msg := ""
        switch callerName
        {
            case "F6":
                msg := "엑셀이 활성화 되지 않은 채 [F6] 버튼을 눌렀습니다.`n상하차 전환을 할 차량정보의 행을 선택 후 눌러주세요."
            case "F7":
                msg := "엑셀이 활성화 되지 않은 채 [F7] 버튼을 눌렀습니다.`n납품 차량정보의 행을 선택 후 눌러주세요."
            case "ALT":
                msg := "엑셀이 활성화 되지 않은 채 [퀵슬롯 등록] 버튼을 눌렀습니다.`n등록할 차량의 행을 선택 후 다시 시도해 주세요."
            case "T":
                msg := "엑셀이 활성화 되지 않은 채 [TRDT] 버튼을 눌렀습니다.`n해당 문구를 입력할 셀을 선택하시고 다시 시도해 주세요."
            case "B":
            msg := "엑셀이 활성화 되지 않은 채 [BTOS] 버튼을 눌렀습니다.`n해당 문구를 입력할 셀을 선택하시고 다시 시도해 주세요."
            default:
                msg := "엑셀이 활성화 되어지지 않은 상태입니다."
        }
        MsgBox, 262208, 알림, % msg
        return false
    }

    return true
}

WaitExcel() {
    global
    maxRetries := 100
    sleepTime := 10

    Loop, %maxRetries%
    {
        try {
            if(xl.Ready && IsObject(xl))
                return true
        } catch {
            return false
        }
        Sleep, %sleepTime%
    }
    return false
}


ReformCarInfo(inputData, time := false, isf7 := false)
{
    if !IsObject(inputData){
        dataArr := StrSplit(inputData, A_Tab)
    }
    else {
        dataArr := inputData
    }

    isOut := (dataArr[5] = "반출")

    finalLine := ""
    Loop, 15
    {
        val := dataArr[A_Index]

        if (A_Index = 8) {
            val := "/"
        } else if (isOut && A_Index >= 9 && A_Index <= 11) {
            val := ""
        }
        else if (A_Index = 12) {
            if(isOut) {
                val := ""
            }
            else if(val != "/"){
                val := ""
            }
        }

        else if (A_Index = 13) {
            if (time) {
                if (isf7) {
                    val := "/"
                }
                else {
                    FormatTime, nowTime,, HH:mm
                    val := nowTime
                }
            }
            else {
                val := ""
            }
        }
        else if (A_Index = 14) {
            if (time) {
                if (isf7) {
                    FormatTime, nowTime,, HH:mm
                    val := nowTime
                } else {
                    val := ""
                }
            }
            else {
                val := ""
            }
        }

        finalLine .= (A_Index = 1 ? "" : A_Tab) . val
    }
    return finalLine
}

;===========================================================
; 가장 아래에 있는 행을 찾는 함수
; @ return Row
;===========================================================
FindLastRow()
{
    global xl
    try
    {
        startRow := 7

        ; A열 마지막 행
        lastRowA := xl.Cells(xl.Rows.Count, 1).End(-4162).Row
        if (lastRowA < startRow) {
            dataCount := 0
        } else {
            dataCount := lastRowA - (startRow - 1)
        }

        if (dataCount = 0) {
            return startRow
        }
        ; C열을 7행부터 한번에 읽기

        lastCheckRow := (startRow - 1) + dataCount
        dataArray := xl.Range("C" startRow ":C" lastCheckRow + 1).Value

        currentRow := 7

        Loop % dataArray.MaxIndex()
        {
            if (Trim(dataArray[A_Index, 1]) = "")
                return startRow + A_Index - 1
        }

        ; C열이 전부 차있으면 다음 입력 위치
        return lastCheckRow + 1
    }
    catch e
    {
        RecordLog("FindLastRow - 에러 발생")
        return 7
    }
}

;===========================================================
; 현재 활성화 되어있는 행에 차량정보가 존재하는지 확인
; @ return boolean
;===========================================================
CarExist() {
    global xl

    Send, {Ctrl Down}{Enter}{Ctrl Up}

    selectionRow := xl.ActiveCell.Row
    rawCarNum := xl.Cells(selectionRow, 3).Value
    checkCarNum := Trim(rawCarNum)

    if (checkCarNum = "" || checkCarNum = "차량번호")
    {
        MsgBox, 262208, 알림, 차량 정보가 없는 행입니다.
        return false
    }
    return true
}


;===========================================================
; 엑셀에 보이고 있는 /전산 차량 갯수를 GUI에 반영
;===========================================================
GetTMSCountFromExcel() {
    global xl
    try {
        lastRow := xl.ActiveSheet.Cells(xl.Rows.Count, "Q").End(-4162).Row

        if (lastRow < 7) {
            count := 0
        } else {
            formula = SUMPRODUCT(SUBTOTAL(3, OFFSET(Q7, ROW(Q7:Q%lastRow%)-7, 0)), --(ISNUMBER(SEARCH("전산", Q7:Q%lastRow%))))
count := Round(xl.Evaluate(formula))
        }

        GuiControl, 1:, Cars, % Format("{:02}", count)
        GuiControl, 1:MoveDraw, Cars
    } catch {
        GuiControl, 1:, Cars, 00
        GuiControl, 1:MoveDraw, Cars
    }
return
}

AutoSlipInput(){
    global xl

    selectionRow := xl.ActiveCell.Row
	rawCarType := xl.Cells(selectionRow, 8).Value
	checkValue := Trim(rawCarType)

	FormatTime, datePart,, yyMMdd
	finaldate := ""

	switch checkValue
	{
		case "1뷰티":
			finaldate := "11110" . datePart . "00"
		case "2뷰티":
			finaldate := "1BTOS" . datePart . "00"
		case "3뷰티":
			finaldate := "1TRDT" . datePart . "00"
		case "대전":
			finaldate := "11111" . datePart . "00"
		case "김천":
			finaldate := "11120" . datePart . "00"
		default:
			finaldate := ""
	}

	if (finaldate != "") {
        Clipboard := ""
        Clipboard := finaldate
        ClipWait, 1
        Send, {F2}^v{Ctrl Up}{F2}
        GuiControl, 1:, Status, 전표번호 입력
    }
	else {
        GuiControl, 1:, Status, 일치하는 지역 없음
    }
    SetTimer, ResetStatus, 3000
    return
}

;===========================================================
; WinActive를 이미 활성화 되어있으면 skip 하는 함수
; @param winTitle
;===========================================================
ActivateWindow(winTitle)
{
    if (WinActive(winTitle))
    {
        return
    }
    WinActivate, %winTitle%
    WinWaitActive, %winTitle%, , 1

    Sleep, 50
}

;===========================================================
; 엑셀의 최적화 기능을 켜고 끔
; @param boolean
;===========================================================
ExcelOptimizer(on)
{
    global xl
    try
    {
        if (on)
        {
            xl.ScreenUpdating := False
            xl.EnableEvents := False
            xl.Calculation := -4135
        }
        else
        {
            xl.Calculation := -4105
            xl.EnableEvents := True
            xl.ScreenUpdating := True
        }
    }
}

;===========================================================
; 엑셀의 시트 이동 함수 이미 선택된 시트이면 아무 동작 하지 않음
; @param sheetRef
;==========================================================
MoveSheet(sheetRef)
{
    global xl
    try
    {
        if (RegExMatch(sheetRef, "^\d+$"))
        {
            if (xl.ActiveSheet.Index != sheetRef)
                xl.Sheets(sheetRef).Select
        }
        else
        {
            if (xl.ActiveSheet.Name != sheetRef)
                xl.Sheets(sheetRef).Select
        }
        WaitExcel()
    }
}

SafeClick(control, pid) {
    ControlClick, %control%, ahk_pid %pid%,,,, D NA
    Sleep, 25
    ControlClick, %control%, ahk_pid %pid%,,,, U NA
}


HandleSlotSelect(slotNum)
{
    global chooseSlotNum

    chooseSlotNum := slotNum
    IniWrite, %slotNum%, %iniPath%, settings, chooseSlotNum
    LoadQuickSlot(slotNum)

    GuiControl, 2:, SettingSlot%slotNum%, 1
    GuiControl, , Status, %slotNum%번 슬롯 선택 완료.

    SetTimer, ResetStatus, 3000
}

HandleSettingSlotSelect(slotNum)
{
    global chooseSlotNum

    chooseSlotNum := slotNum
    IniWrite, %slotNum%, %iniPath%, settings, chooseSlotNum
    LoadQuickSlot(slotNum)

    GuiControl, 1:, Slot%slotNum%, 1
    GuiControl, 1:, Status, %slotNum%번 슬롯 설정 선택 완료

    SetTimer, ResetStatus, 3000
}


HandleCarInput(idx)
{
    global

    if(!CheckExcel()) {
        return
    }

    ActivateWindow(excelName)

    try
    {
        Send, ^{Enter}{Ctrl up}

        ; 1. INI에서 해당 슬롯의 탭 문자열을 그대로 읽어옴
        IniRead, savedLine, %iniPath%, slot%chooseSlotNum%, %idx%
        savedLine := Trim(savedLine, """")

        if (savedLine = "ERROR" || savedLine = "") {
            MsgBox, 262208, 알림, 해당 슬롯에 데이터가 없습니다.
            return
        }
        row := StrSplit(savedLine, A_Tab)

        if (Trim(row[1]) = "") {
            MsgBox, 262208, 알림, 해당 슬롯에 차량번호가 없습니다.
            return
        }

        ExcelOptimizer(true)

        finalLine := ReformCarInfo(row, true)

        MoveSheet(1)

        ; 빈 행 찾기
        targetRow := FindLastRow()

        Clipboard := finalLine
        ClipWait, 1

		ExcelOptimizer(false)

        InputCarInfo(targetRow)

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
    global

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

        IniDelete, %iniPath%, slot%chooseSlotNum%, %idx%
        IniWrite, %newDataLine%, %iniPath%, slot%chooseSlotNum%, %idx%

        LoadQuickSlot(chooseSlotNum)
        xl.Application.CutCopyMode := False

        GuiControl, 1:, Status, % chooseSlotNum "번 슬롯 " idx "번 교체 완료"
        SetTimer, ResetStatus, 3000
    }
    catch
    {
        RecordLog("!" idx " - 실패")
    }
}

InputCarInfo(targetRow){
	global xl
	try {
        targetAddr := "C" . targetRow

        xl.Range(targetAddr).Select

        Loop, 20 {
            if (xl.ActiveCell.Address(0, 0) = targetAddr)
                break
            Sleep, 10
        }

        WaitExcel()

        Send, ^v{Ctrl Up}

        WaitExcel()
    }
    catch {
        ; 객체 참조 오류 발생 시 무시 또는 예외 처리
    }
}


