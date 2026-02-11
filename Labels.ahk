;=================================================================
; LABELS
;=================================================================
ResetStatus:
	GuiControl, , Status, 대기중
	return

GuiClose:
	ExitApp

ExplainBtn:
	MsgBox, 262208, 도움말,
	(
[F1] - 차량 입출 TMS 새로고침, 엑셀 출문 필터링
[F3] - 통합차량 조회 -> 차량번호로 찾기. 설정 검색방향 선택시
       '위로' = 전날의 첫행부터, '아래로' = 지정한 행부터 시작
[F4] - 현재 선택된 행의 차량정보를 일일차량현황 빈칸에 입력
[F6] - 납품 <-> 반출 전환
[F7] - 납품 -> 반출로 빈칸에 입력
[F11] - TMS, 차량현황 엑셀을 프로그램에 등록

[CTRL]+[1~7] - 퀵슬롯에 저장된 차량정보 빈칸에 입력
[ALT]+[1~7] - 선택한 행의 차량정보 퀵슬롯에 등록

[NUMLK] - "/전산" 입력
[CTRL]+[TAB] - 현재 시간 입력
[CTRL]+[B,T] - 설정에 따라 출고지별 전표번호 입력
               또는 "BTOS","TRDT" 입력

[INS] - 현재 선택된 행의 차량번호를 TMS 등록창에 입력
[SCRLK] - TMS 차량입출등록 창 저장버튼 클릭
[PAUSE] - TMS 차량입출등록 창 취소버튼 클릭

* 처음 실행 후 차량 조회, 등록 TMS를 실행하고
  오늘 날짜의 일일차량 엑셀 파일을 열고 등록을 진행해 주세요

* 작동이 아예 멈추거나 문제가 생겼을땐 껐다키고
  다시 등록해 사용해주세요

* TMS이상으로 TMS가 켜지지 않을땐
  기타설정에서 '엑셀만 사용'을 선택하고 사용해주세요

* 구버전에 없던 문제가 생기면 임시방편으로
  구버전 'AssistantTool'을 사용해주세요
)
	return

HotkeySettingBtn:
	Gui, 2: Show, x345 y0 w935 h264, 퀵슬롯 설정
	return

OtherSettingBtn:
	LoadSettings()
	Gui, 3: Show, x345 y0 w450 h264, 기타 설정
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