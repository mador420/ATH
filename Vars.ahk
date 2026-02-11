; ==========================================================
; Vars.ahk - 전역 변수 및 초기 설정 관리
; ==========================================================

global iniPath := A_ScriptDir "\athsettings.ini"

; 1. 기본 설정 및 자동화 관련
global  searchStartRow, chooseSlotNum, searchto, autoslip, onlyexcel

; 2. 엑셀 TMS 객체 관리
global xl, excelName, tms1Pid, tms2Pid

; 3. 슬롯별 전체 데이터
global Car1Data, Car2Data, Car3Data, Car4Data, Car5Data, Car6Data, Car7Data

; 4. 슬롯별 개별 상세 정보
global car1Num, car1Name, car1Company, car1Phone, car1Content, car1Carry, car1Drop, car1Card
global car2Num, car2Name, car2Company, car2Phone, car2Content, car2Carry, car2Drop, car2Card
global car3Num, car3Name, car3Company, car3Phone, car3Content, car3Carry, car3Drop, car3Card
global car4Num, car4Name, car4Company, car4Phone, car4Content, car4Carry, car4Drop, car4Card
global car5Num, car5Name, car5Company, car5Phone, car5Content, car5Carry, car5Drop, car5Card
global car6Num, car6Name, car6Company, car6Phone, car6Content, car6Carry, car6Drop, car6Card
global car7Num, car7Name, car7Company, car7Phone, car7Content, car7Carry, car7Drop, car7Card