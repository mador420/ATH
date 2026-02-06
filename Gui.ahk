; --- GUI 1: 메인 화면 ---
Gui, Add, GroupBox, x5 y25 w210 h230
Gui, Add, GroupBox, x220 y25 w115 h93
Gui, Font, S15, 맑은 고딕
Gui, Add, Text, x5 y0 w290 h25 vStatus
Gui, Add, GroupBox, x300 y-10 w35 h40
Gui, Add, Text, x300 y0 w35 h25 vCars Center +BackgroundTrans
Gui, Font, S13, 맑은 고딕

; 1~7번 텍스트 및 에디트(View) 생성
Loop, 7 {
    yPos := 10 + (A_Index * 30) ; y좌표 자동 계산 (40, 70, 100...)
    Gui, Add, Text, x10 y%yPos% w20 h20, %A_Index%
    Gui, Add, Edit, x25 y%yPos% w115 h25 ReadOnly vCar%A_Index%NumView
    Gui, Add, Edit, x147 y%yPos% w60 h25 ReadOnly vCar%A_Index%NameView
}

; 라디오 버튼 1~6번
Gui, Add, Radio, x230 y35 w50 h25 vSlot1 gSlot1, 1번
Gui, Add, Radio, x285 y35 w50 h25 vSlot2 gSlot2, 2번
Gui, Add, Radio, x230 y60 w50 h25 vSlot3 gSlot3, 3번
Gui, Add, Radio, x285 y60 w50 h25 vSlot4 gSlot4, 4번
Gui, Add, Radio, x230 y85 w50 h25 vSlot5 gSlot5, 5번
Gui, Add, Radio, x285 y85 w50 h25 vSlot6 gSlot6, 6번

Gui, Add, Button, x225 y127 w110 h25 Center gExplainBtn, 설명
Gui, Add, Button, x225 y160 w110 h25 Center gHotkeySettingBtn, 퀵슬롯 설정
Gui, Add, Button, x225 y193 w110 h25 Center gOtherSettingBtn, 기타 설정
Gui, Add, Button, x225 y226 w110 h25 Center vRegisterBtn gRegisterBtn, 등록 (F11)

; --- GUI 2: 상세 설정 ---
Gui, 2: Add, GroupBox, x5 y25 w790 h230
Gui, 2: Add, GroupBox, x800 y25 w115 h93
Gui, 2: Font, S13, 맑은 고딕

; 헤더 텍스트
Gui, 2: Add, Text, x25 y15 w100 h25 Center, 차량 번호
Gui, 2: Add, Text, x147 y15 w60 h25 Center, 성명
Gui, 2: Add, Text, x214 y15 w110 h25 Center, 업체명
Gui, 2: Add, Text, x331 y15 w125 h25 Center, 연락처
Gui, 2: Add, Text, x463 y15 w45 h25 Center, 업무
Gui, 2: Add, Text, x567 y15 w60 h25 Center, 상차지
Gui, 2: Add, Text, x634 y15 w60 h25 Center, 하차지
Gui, 2: Add, Text, x701 y15 w85 h25 Center, 카드/전산

; 1~7번 상세 입력칸(Edit) 생성
Loop, 7 {
    yPos := 10 + (A_Index * 30)
    Gui, 2: Add, Text, x10 y%yPos% w20 h20, %A_Index%
    Gui, 2: Add, Edit, x25 y%yPos% w115 h25 vCar%A_Index%NumEdit
    Gui, 2: Add, Edit, x147 y%yPos% w60 h25 vCar%A_Index%NameEdit
    Gui, 2: Add, Edit, x214 y%yPos% w110 h25 vCar%A_Index%CompanyEdit
    Gui, 2: Add, Edit, x331 y%yPos% w125 h25 vCar%A_Index%PhoneEdit
    Gui, 2: Add, Edit, x463 y%yPos% w45 h25 vCar%A_Index%ContentEdit
    Gui, 2: Add, Edit, x567 y%yPos% w60 h25 vCar%A_Index%CarryEdit
    Gui, 2: Add, Edit, x634 y%yPos% w60 h25 vCar%A_Index%DropEdit
    Gui, 2: Add, Edit, x701 y%yPos% w85 h25 vCar%A_Index%CardEdit
}

; GUI 2 사이드 메뉴
Gui, 2: Add, Radio, x811 y35 w50 h25 vSettingSlot1 gSettingSlot1, 1번
Gui, 2: Add, Radio, x866 y35 w50 h25 vSettingSlot2 gSettingSlot2, 2번
Gui, 2: Add, Radio, x811 y60 w50 h25 vSettingSlot3 gSettingSlot3, 3번
Gui, 2: Add, Radio, x866 y60 w50 h25 vSettingSlot4 gSettingSlot4, 4번
Gui, 2: Add, Radio, x811 y85 w50 h25 vSettingSlot5 gSettingSlot5, 5번
Gui, 2: Add, Radio, x866 y85 w50 h25 vSettingSlot6 gSettingSlot6, 6번
Gui, 2: Add, Button, x811 y220 w110 h25 Center gSettingWriteBtn, 저장

; --- GUI 3: 기타 설정 ---
Gui, 3: Font, S13, 맑은 고딕
Gui, 3: Add, Text, x25 y5 w300 h25, 입력되는 행 스크롤
Gui, 3: Font, S12, 맑은 고딕
Gui, 3: Add, Radio, x25 y30 w100 h25 vinputscroll1, 중앙
Gui, 3: Add, Radio, x130 y30 w100 h25 vinputscroll2, 엑셀 기본

Gui, 3: Font, S13, 맑은 고딕
Gui, 3: Add, Text, x25 y65 w300 h25, 전표번호 단축키
Gui, 3: Font, S12, 맑은 고딕
Gui, 3: Add, Radio, x25 y90 w150 h25 vautoslip1, 지역별 자동입력
Gui, 3: Add, Radio, x190 y90 w150 h25 vautoslip2, BTOS / TRDT

Gui, 3: Font, S13, 맑은 고딕
Gui, 3: Add, Text, x25 y125 w200 h25, 차량 검색 방향
Gui, 3: Font, S12, 맑은 고딕
Gui, 3: Add, Radio, x25 y150 w100 h25 vsearchto1, ↑위로
Gui, 3: Add, Radio, x130 y150 w100 h25 vsearchto2, ↓아래로


Gui, 3: Font, S13, 맑은 고딕
Gui, 3: Add, Text, x25 y185 w300 h25, 검색 시작 행(아래로 조회시)
Gui, 3: Add, Edit, x250 y185 w100 h25 vsearchStartRow

Gui, 3: Font, S13, 맑은 고딕

Gui, 3: Add, Text, x25 y220 w300 h25, TMS 대기 (50~150)
Gui, 3: Add, Edit, x180 y220 w50 h25 vtmsIdleTime

Gui, 3: Add, Button, x240 y220 w110 h25 Center gOtherWriteBtn, 저장

GuiControl, , Status,TMS, 엑셀 미등록 상태입니다
GuiControl, , Cars, 00