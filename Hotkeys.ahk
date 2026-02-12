;=============================================================
;핫키
;=============================================================

$F11::
{
	RecordLog("F11 Pressed")
	RegistPrograms()
	return
}

$F1::
{
	if(!CheckTMS())	{
        return
    }
	if(!CheckExcel()) {
		return
	}

	try
    {
		GuiControl, 1:, Status,엑셀, TMS 새로고침중

		if(onlyexcel = 1) {
            ActivateWindow("ahk_pid " . tms1Pid)
            SafeClick("Button22", tms1Pid)
        }

        ActivateWindow(excelName)
        SendInput, {Ctrl Down}{Enter}{Ctrl Up}

        ExcelOptimizer(true)

        MoveSheet(1)

		xl.Range("P6").AutoFilter(16, "=")
        WaitExcel()
		GetTMSCountFromExcel()
        ExcelOptimizer(false)
	}
    catch
    {
        RecordLog("F1 - 실패 (엑셀 객체 오류)")
        ExcelOptimizer(false)
        return
    }

    ; 5. 마무리 및 상태 초기화
    GuiControl, 1:, Status, 엑셀, TMS 새로고침 완료
    SetTimer, ResetStatus, 3000
    return
}

$F3::
{
    if(!CheckExcel()) {
		return
	}
	GuiControl, 1:, Status, 차량번호 조회로 이동
	ActivateWindow(excelName)


	try {
        SendInput, {Ctrl Down}{Enter}{Ctrl up}
		MoveSheet(2)
        ExcelOptimizer(true)

		if (searchto = 1) {
            targetRow := RegExReplace(fromTopRow, "\D")
        } else if (searchto = 2) {
            baseRow := xl.Cells(xl.Rows.Count, 3).End(-4162).Row - RegExReplace(fromBottomRow, "\D")
            targetRow := xl.Cells(baseRow, 3).End(-4162).Row
        }
		if (targetRow < 6) {
			targetRow := 7
		}

        xl.Columns("C").Select
        WaitExcel()
        ExcelOptimizer(false)
        xl.Cells(targetRow, 3).Activate
        WaitExcel()

		SendInput, ^f{End}+{Home}
	} catch {
        ExcelOptimizer(false)
		RecordLog("F3 동작중 실패")
	}

	GuiControl, 1:, Status, 대기 중
	return
}


$F4::
{
    if(!CheckExcel()) {
		return
	}

    ; 찾기 창이나 엑셀 자체가 활성화되어 있는지 추가 체크
    if !WinActive("찾기 및 바꾸기") && !WinActive(excelName)	{
        RecordLog("F4 - 엑셀 미활성화 시도")
        MsgBox, 262208, 알림, [F4]는 찾기 도중 혹은 엑셀 활성화 상태에서만 가능합니다.`n찾거나 복사할 차량정보를 선택 후 눌러주세요.
        return
    }

    GuiControl, 1:, Status, 차량정보 복사 중

    try	{
        ; 2. 차량번호 유효성 검사 (키보드 전송 대신 객체로 직접 확인)
        selectionRow := xl.ActiveCell.Row
        rawCarNum := xl.Cells(selectionRow, 3).Value ; C열(3) 값 확인
		checkCarNum := Trim(rawCarNum)

        if (checkCarNum = "" || checkCarNum = "차량번호")
        {
            RecordLog("F4 - 차량번호 빈칸")
            MsgBox, 262208, 알림, 차량번호 셀이 빈칸입니다.`n복사할 차량정보 행을 선택 후 눌러주세요.
            SetTimer, ResetStatus, 3000
            return
        }
		lastVal := xl.Cells(selectionRow, 17).Value ; Q열 (카드/전산 정보)

        ; 3. 데이터 복사 (C열~Q열)
        rowLine := "C" . selectionRow . ":Q" . selectionRow
        xl.Range(rowLine).Copy
        ClipWait, 1

        ; 4. 탭(A_Tab) 기준으로 분리하여 배열 생성
        newRowData := ReformCarInfo(Clipboard, true)

		xl.Application.CutCopyMode := False
		Clipboard := ""
		Clipboard := newRowData
		ClipWait, 1

        ActivateWindow(excelName)
        ExcelOptimizer(true)
		MoveSheet(1)

		targetRow := FindLastRow()
        ExcelOptimizer(false)
        WaitExcel()

		InputCarInfo(targetRow)

        if (lastVal = "카드/전산" || lastVal = "48/전산" || lastVal = "50/전산")	{
            xl.Range("K" . targetRow).Select
		}
        else {
            xl.Range("Q" . targetRow).Select
		}

        ; 성공 마무리
        GuiControl, 1:, Status, 차량정보 복사 완료
        SetTimer, ResetStatus, 3000
    }
    catch
    {
        ExcelOptimizer(false)
        RecordLog("F4 - 실패")
        MsgBox, 262208, 에러, 엑셀 작업 중 오류가 발생했습니다.
    }
    return
}

$F6::
{
	if(!CheckExcel(true, "F6")) {
		return
	}
	if(!CarExist()){
		return
	}

    try
    {
        ExcelOptimizer(true)
        ; 1. 현재 선택된 행 객체 생성
        selectionRow := xl.Selection.Row

        ; 한 번에 읽기 (3번 호출 → 1번 호출)
        rangeAddr := "G" . selectionRow . ":I" . selectionRow
        values := xl.Range(rangeAddr).Value

        statusVal := values[1][1]  ; G열
        hVal      := values[1][2]  ; H열
        iVal      := values[1][3]  ; I열

        ; 새 값 준비
        if (statusVal = "반출")
        {
            newValues := [["납품", iVal, "1뷰티"]]
            GuiControl, 1:, Status, 반출 -> 납품 전환 완료
        }
        else if (statusVal = "납품")
        {
            newValues := [["반출", "1뷰티", hVal]]
            GuiControl, 1:, Status, 납품 -> 반출 전환 완료
        }
        else
        {
            GuiControl, 1:, Status, 반출/납품 업무가 아님
            SetTimer, ResetStatus, -3000
            return
        }

        ; 한 번에 쓰기
        xl.Range(rangeAddr).Value := newValues
        ExcelOptimizer(false)
    }
    catch
    {
        RecordLog("F6 - 실패")
        ExcelOptimizer(false)
    }

    SetTimer, ResetStatus, 3000
    return

}

$F7::
{
	if(!CheckExcel(true, "F7")) {
		return
	}
	if(!CarExist()){
		return
	}

    GuiControl, 1:, Status, 납품 -> 반출 재입력 중

    try
    {
        currRow := xl.Selection.Row

		sourceType := xl.Range("G" . currRow).Value ; 5번째 열(G열)
        if (sourceType != "납품")
        {
            MsgBox, 262208, 알림, 납품 데이터가 아닙니다.
            xl.Application.CutCopyMode := False
            return
        }

        Clipboard := ""
        xl.Range("C" . currRow . ":R" . currRow).Copy
        ClipWait, 1

        dataArr := StrSplit(Clipboard, A_Tab)

        dataArr[5] := "반출"     ; G열
        placeTemp := dataArr[6]  ; 기존 H열(장소1) 백업
        dataArr[6] := "1뷰티"    ; H열에 새 값 주입
        dataArr[7] := placeTemp  ; I열에 기존 장소1 주입

        newRowData := ReformCarInfo(dataArr, true, true)

        ; 4. 기존 행 P열 수정 (이건 개별 수정 필수)
        ExcelOptimizer(true)
        xl.Range("P" . currRow).Value := "/"

        ; 5. 타겟 행 찾기 및 붙여넣기
        targetRow := FindLastRow()

        xl.Application.CutCopyMode := False
        Clipboard := ""
        Clipboard := newRowData
        ClipWait, 1

        ExcelOptimizer(false)

		InputCarInfo(targetRow)

		xl.Range("K" . targetRow).Select

		GuiControl, 1:, Status, 납품 -> 반출 재입력 완료

    }
    catch
    {
        ExcelOptimizer(false)
        RecordLog("F7 - 실패")
		GuiControl, 1:, Status, 납품 -> 반출 재입력 실패
        return
    }

	SetTimer, ResetStatus, 3000
    return
}


$NumLock::
{
	if(!CheckExcel(true)) {
		return
	}
    Send, /전산
    GuiControl, 1:, Status, "/전산" 입력
    SetTimer, ResetStatus, 3000
    return
}

$Insert::
{
    if(!CheckTMS(true, "Insert")) {
        return
    }
    if(!CheckExcel(true)) {
		return
	}

    ; 편집모드 해제
    SendInput, {Ctrl Down}{Enter}{Ctrl up}
    ; 현재 선택된 셀의 값 가져오기
    try {

        inputCarNum := xl.Cells(xl.ActiveCell.Row, 3).Value

        if(inputCarNum = "" || inputCarNum = "차량번호") {
            RecordLog("Insert - 빈칸 시도")
            MsgBox, 262208, 알림, 차량번호 셀이 빈칸입니다.`n입력할 차량정보 행을 선택 후 눌러주세요.
            return
        }
    } catch {
        RecordLog("Insert - 엑셀 객체 연결 실패")
        return ; ExitApp 보다는 return으로 핫키만 종료하는 것이 안전합니다.
    }

    ; 3. TMS 입력 작업 시작
    GuiControl, 1:, Status, 차량 정보 입력 중

    ; TMS 창 활성화 및 대기

    ActivateWindow("ahk_pid " . tms2Pid)


    ; 텍스트 입력 (ControlSetText는 신뢰도가 높지만 입력 후 대기가 필요할 수 있음)
    ControlSetText, PBEDIT1052, %inputCarNum%, ahk_pid %tms2Pid%
    sleep 100
    ; 버튼 클릭
    SafeClick("Button9", tms2Pid)
    sleep 50

    MouseMove, 1395, 407, 0

    GuiControl, 1:, Status, 차량 정보 입력 완료
    SetTimer, ResetStatus, 3000

    return
}

$^Tab::
{
	if(!CheckExcel(true)) {
		return
	}

    FormatTime, nowTime,, HH:mm
    Clipboard := nowTime
    ClipWait, 1
    Send, {Ctrl Down}v{Ctrl Up}
    GuiControl, 1:, Status, 현재시간 입력
    SetTimer, ResetStatus, 3000
    return
}

$^t::
{
	if(!CheckExcel(true, "T")) {
		return
	}
	if (!CarExist()) {
		return
	}
	if (autoslip = 1) {
		AutoSlipInput()
	} else {
		Clipboard := "TRDT"
		ClipWait, 1
		Send, {F2}^v{Ctrl Up}
		GuiControl, 1:, Status, "TRDT" 입력
		SetTimer, ResetStatus, 3000
	}
    return
}

$^b::
{
	if(!CheckExcel(true, "B")) {
		return
	}
	if (!CarExist()) {
		return
	}
	if (autoslip = 1) {
		AutoSlipInput()
	} else {
		Clipboard := "BTOS"
		ClipWait, 1
		Send, {F2}^v{Ctrl Up}
		GuiControl, 1:, Status, "BTOS" 입력
		SetTimer, ResetStatus, 3000
	}
    return
}

$ScrollLock::
{
    if(!CheckTMS(true, "ScrollLock", true)) {
        return
    }
    ActivateWindow("ahk_pid " . tms2Pid)
    SafeClick("Button12", tms2Pid)
	sleep 200
    ActivateWindow("ahk_pid " . tms1Pid)
    SafeClick("Button22", tms1Pid)
    return
}

$Pause::
{
    if(!CheckTMS(true, "Pause", true)) {
        return
    }
    ActivateWindow("ahk_pid " . tms2Pid)
    SafeClick("Button15", tms2Pid)
	return
}

$^1::HandleCarInput(1)
$^2::HandleCarInput(2)
$^3::HandleCarInput(3)
$^4::HandleCarInput(4)
$^5::HandleCarInput(5)
$^6::HandleCarInput(6)
$^7::HandleCarInput(7)

$!1::RegisterSlotFromExcel(1)
$!2::RegisterSlotFromExcel(2)
$!3::RegisterSlotFromExcel(3)
$!4::RegisterSlotFromExcel(4)
$!5::RegisterSlotFromExcel(5)
$!6::RegisterSlotFromExcel(6)
$!7::RegisterSlotFromExcel(7)

/*
#If (WinActive("찾기 및 바꾸기") || WinActive("ahk_class #32770")) && (searchto = 1)

$NumpadEnter::
$Enter::
    Send, {Shift Down}{Enter}{Shift Up}
    return

#If
*/