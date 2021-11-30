#include <Misc.au3>       ; _Singleton
;#include <MsgBoxConstants.au3>
#include <WinAPI.au3>

Opt("TrayMenuMode", 3)

If _Singleton("winkey", 1) = 0 Then
   PopUp("An occurrence of winkey is already running",3)
   Exit
EndIf

Global $hDLL = DllOpen("user32.dll")
Global $hHook, $keystrokes = "", $mainloop = 1, $kCapLock= False, $NumLock = False, $NumpadClear = False, $PrtScKey = False
Global $kShift = False, $kCtrl = False, $kAlt = False, $kWin = False, $AppsKey = False, $StopKey = False, $EndLoop = False

Global $idHotkeys = TrayCreateItem("About WinKey")
TrayCreateItem("")
;Global $idREC = TrayCreateItem("Record Scripts")
;TrayCreateItem("")
;Global $idLoadScript = TrayCreateItem("Load Script From Text File")
;Global $idLoadScriptXls = TrayCreateItem("Load Macros From Excel File")
;TrayCreateItem("")
;Global $idPlay = TrayCreateItem("Play Scripts")
;Global $idMacros = TrayCreateItem("Run from macros")
;Global $idTasker = TrayCreateItem("Activate timer")
;TrayCreateItem("")
;Global $idLoadCfg = TrayCreateItem("Load Config from file")
;Global $idEditCfg = TrayCreateItem("Edit Config")
;TrayCreateItem("")
Global $idExit = TrayCreateItem("Exit")
Global $keycodes, $keytables, $skeycodes, $skeytables, $sskeytables, $tkeycodes, $tkeytables, $ttkeytables

;------------------------------------------------------------------------------------------------
MAIN()
;------------------------------------------------------------------------------------------------

Func MAIN()
   Local $arr_items, $hmod, $hStub_KeyProc, $i, $n, $s
   $arr_items = "3,8,9,13,16,17,18,19,20,27,32,33,34,35,36,37,38,39,40,44,45,46,95,112,113,114,115,116,117,118,119,120,121,122,123,144,145,255"
   $keycodes = StringSplit($arr_items , ",", $STR_ENTIRESPLIT)
   $arr_items = "{BREAK},{BS},{TAB},{ENTER},{LSHIFT},{LCTRL},{LALT},{PAUSE},{CAPSLOCK},{ESC}, ,{PGUP},{PGDN},{END},{HOME},{LEFT},{UP},{RIGHT},{DOWN},{PRINTSCREEN},{INS},{DEL},{SLEEP},{F1},{F2},{F3},{F4},{F5},{F6},{F7},{F8},{F9},{F10},{F11},{F12},{NUMLOCK},{SCROLLLOCK},{BREAK}"
   $keytables = StringSplit($arr_items , ",", $STR_ENTIRESPLIT)
   $arr_items = "96,97,98,99,100,101,102,103,104,105,106,107,109,110,111"
   $skeycodes = StringSplit($arr_items , ",", $STR_ENTIRESPLIT)
   $arr_items = "0,1,2,3,4,5,6,7,8,9,{*},{+},{-},{.},{/}"
   $skeytables = StringSplit($arr_items , ",", $STR_ENTIRESPLIT)
   $arr_items = "{INS},{END},{DOWN},{PGDN},{LEFT},{},{RIGHT},{HOME},{UP},{PGUP},{*},{+},{-},{.},{/}"
   $sskeytables = StringSplit($arr_items , ",", $STR_ENTIRESPLIT)
   $arr_items = "48,49,50,51,52,53,54,55,56,57,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,186,187,188,189,190,191,192,219,220,221"
   $tkeycodes = StringSplit($arr_items , ",", $STR_ENTIRESPLIT)
   $arr_items = "0,1,2,3,4,5,6,7,8,9,a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z,;,=,,-,.,/,`,[,\,]"
   $tkeytables = StringSplit($arr_items , ",", $STR_ENTIRESPLIT)
   $arr_items = "),{!},@,#,$,{%},{^},&,*,(,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,:,{+},<,_,>,?,{~},{,|,}"
   $ttkeytables = StringSplit($arr_items , ",", $STR_ENTIRESPLIT)

    $hStub_KeyProc = DllCallbackRegister("_KeyProc", "long", "int;wparam;lparam")
   $hmod = _WinAPI_GetModuleHandle(0)
   $hHook = _WinAPI_SetWindowsHookEx($WH_KEYBOARD_LL, DllCallbackGetPtr($hStub_KeyProc), $hmod)

   AboutMe()
   While $mainloop == 1
	  If $PrtScKey Then
		 If $keystrokes=="" Then
			Send( "{LWINDOWN}{LWINUP}" )
		 Else
			$s = StringRight( $keystrokes, 1 )
			Send( "{LWINDOWN}" & $s & "{LWINUP}" )
		 EndIf
		 $keystrokes = ""
		 $PrtScKey = False
	  EndIf
	  Sleep(20)
	  CheckTray()
   WEnd
   DllClose($hDLL)
   PopUp("Thanks for using WinKey app !")
EndFunc   ;==>MAIN

Func AboutMe()
   PopUp("This app uses PRINTSCREEN button as the Windows key")
EndFunc

Func CheckTray()
    Switch TrayGetMsg()
        Case $idExit
            $mainloop = 0
		 Case $idHotkeys
			AboutMe()
    EndSwitch
EndFunc   ;==>CheckTray

Func mapkey($keycode2)
   Local $m, $n, $p, $w, $buf
   $AppsKey = False
   $StopKey = False
   $NumpadClear  = False
   $PrtScKey = False
   $buf = ""
   Switch  $keycode2
	  Case 12 ; Numpad 5 with Numlock=On
		 $NumpadClear = True
		 $StopKey = True
	  Case 20
		 $kCapLock = not $kCapLock
	  Case 44
		 $PrtScKey = True
	  Case 91 ; Left Windows Key
		 $buf = ""
		 $kWin = True
	  Case 92 ; Right Windows Key
		 $buf = ""
		 $kWin = True
	  Case 93
		 $buf = ""
		 $AppsKey = True
	  Case 160 to 161
		 $kShift = True
		 $buf = ""
	  Case 162 to 163
		 $kCtrl= True
		 $buf = ""
	  Case 164 to 165
		 $kAlt = True
		 $buf = ""
	  Case Else
		 For $n = 1 to Number($tkeycodes[0])
			$p = Number($tkeycodes[$n])
			If ($p ==  $keycode2) Then
			   If $kShift Then
				  $buf = $ttkeytables[$n]
				  $kShift = False
			   Else
				  $buf = $tkeytables[$n]
			   EndIf
			   ExitLoop
			EndIf
		 Next
		 If $buf <> "" Then
			$buf = $kCtrl  ? ("^" & $buf) : $buf
			$kCtrl = False
			$buf = $kAlt  ? ("!" & $buf) : $buf
			$kAlt = False
			$buf = $kWin  ? ("#" & $buf) : $buf
			$kWin = False
			Return $buf
		 EndIf
		 If $buf == "" Then
			$m = $skeycodes[0]
			for $n = 1 to $m
			   $p = Number($skeycodes[$n])
			   If ($p ==  $keycode2) Then
				  If $NumLock Then
					 $buf = $skeytables[$n]
				  Else
					 $buf = $sskeytables[$n]
				  EndIf
				  ExitLoop
			   EndIf
			Next
		 EndIf
		 If $buf == "" Then
			$m = $keycodes[0]
			for $n = 1 to $m
			   $p = Number($keycodes[$n])
			   If ($p ==  $keycode2) Then
				  $buf = $keytables[$n]
				  ExitLoop
			   EndIf
			Next
		 EndIf
		 If $buf <> "" Then
			$buf = $kShift ? ("+" & $buf) : $buf
			$kShift = False
			$buf = $kCtrl  ? ("^" & $buf) : $buf
			$kCtrl = False
			$buf = $kAlt  ? ("!" & $buf) : $buf
			$kAlt = False
			$buf = $kWin  ? ("#" & $buf) : $buf
			$kWin = False
			Return $buf
		 EndIf
   EndSwitch
   If $kWin And $kCtrl Then
	  $StopKey = True
   EndIf
   Return $buf
EndFunc

Func _KeyProc($nCode, $wParam, $lParam) ; $wParam = 256257 $lParam = 1034950010349500
	Local $tKEYHOOKS
	$tKEYHOOKS = DllStructCreate($tagKBDLLHOOKSTRUCT, $lParam)
	If $nCode < 0 Then
		Return _WinAPI_CallNextHookEx($hHook, $nCode, $wParam, $lParam)
	EndIf
	If $wParam = 256 Then
		EvaluateKey(DllStructGetData($tKEYHOOKS, "vkCode"))
	Else
		Local $flags = DllStructGetData($tKEYHOOKS, "flags")
		Switch $flags
			Case $LLKHF_ALTDOWN
				EvaluateKey(DllStructGetData($tKEYHOOKS, "vkCode"))
		EndSwitch
	EndIf
	Return _WinAPI_CallNextHookEx($hHook, $nCode, $wParam, $lParam)
EndFunc

Func EvaluateKey($keycode)
   Local $buffer
   $buffer = mapkey($keycode)
   $keystrokes = $keystrokes  & $buffer
 EndFunc

Func Popup($msg, $tm = Default)
   if $tm == Default Then
	  MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Windows Key App by KH", $msg)
   Else
	  MsgBox($MB_SYSTEMMODAL + $MB_ICONINFORMATION, "Windows Key App by KH", $msg, $tm)
   EndIf
EndFunc   ;==>Flash

