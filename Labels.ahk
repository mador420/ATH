ResetStatus:
	GuiControl, , Status, 대기중
	return

GuiClose:
	ExitApp

ExplainBtn:
	MsgBox,[F1] - 차량 출입 현황 재필터링, 차량 입출조회 재조회.`n[F3] - 통합차량 조회 -> 차량번호로 찾기.`n[F4] - 현재 선택된 행의 차량정보를 일일차량현황 빈칸에 입력.`n[F6] - 납품 <-> 반출 전환.`n[F7] - 납품 -> 반출로 빈칸에 입력. `n[INS] - 현재 선택된 행의 차량번호를 TMS 등록창에 입력.`n[CTRL]+[B,T] - "BTOS","TRDT" 입력.`n[NUMLK] - "/전산" 입력.`n[CTRL]+[TAB] - 현재 시간 입력.`n[CTRL]+[1~7] - 핫키에 저장된 차량정보 빈칸에 입력.`n[ALT]+[1~7] - 선택한 행의 차량정보 핫키에 등록.`n[F11] - TMS, 차량현황 엑셀을 프로그램에 등록.`n[SCRLK] - 클릭.`n`n`n`n* 처음 실행 후 차량 조회, 등록 TMS와`n오늘 날짜의 일일차량 엑셀 파일을 실행 하고 등록을 진행해 주세요.`n`n* 사용 중 명령에 이상이 있을땐`n대기시간을 조정해 속도를 느리게 설정해주세요.`n`n* 작동이 아예 멈추거나 문제가 생겼을땐`n껐다가 키고 다시 등록해 사용해주세요.
	return

HotkeySettingBtn:
	Gui, 2: Show, x330 y0 w935 h264, 퀵슬롯 설정
	return

OtherSettingBtn:
	LoadSettings()
	Gui, 3: Show, x330 y0 w370 h264, 기타 설정
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