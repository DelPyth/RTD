#NoEnv
#SingleInstance, Force
#NoTrayIcon
#KeyHistory 0
#MaxThreadsPerHotkey, 1
ListLines, Off
AutoTrim, Off
SendMode, Input
DetectHiddenWindows,On
SetWorkingDir, %A_ScriptDir%
SetTitleMatchMode, 3
SetBatchLines -1
SetWinDelay, -1
SetKeyDelay, -1, -1
SetMouseDelay, -1

Menu, Edit, Add, &Roll`tAlt+R, Roll
Menu, Edit, Add, R&eset`tCtrl+R, Reset

Menu, Load, Add, % "File`tCtrl+O", Open
Menu, Load, Add, % "Clipboard`tAlt+P", Open
Menu, Action, Add, Load From..., :Load
Menu, Action, Add
Menu, Action, Add, Close`tCtrl+W, Close

Menu, Main, Add, &Edit, :Edit
Menu, Main, Add, &Action, :Action

Loop, 10
	Info .= "Option " A_Index "|"
Coin := ["Heads", "Tails"]

Try
	Menu, Tray, Icon, % A_ScriptDir "\Icon.ico"

Gui, 1:New, -MinimizeBox -Theme, Roll the Dice
Gui, 1:Menu, Main
Gui, 1:Add, Tab2, x0 y0 w380 h290 vTabs, Multi-Sided Dice|Option Chooser
Gui, 1:Tab, 1, 1
Gui, 1:Add, Edit, x60 y25 w100 h21 vMin +Range1-1000 +Number, 1
Gui, 1:Add, UpDown, x135 y25 w25 h20 vMin_, 1
Gui, 1:Add, Edit, x60 y55 w100 h21 vMax +Range1-1000 +Number, 6
Gui, 1:Add, UpDown, x135 y55 w25 h20 vMax_, 6
Gui, 1:Add, Text, x10 y25 w50 h20 0x200, Min
Gui, 1:Add, Text, x10 y55 w50 h20 0x200, Max
Gui, 1:Add, Text, x10 y95 w50 h15, Presets
Gui, 1:Add, DropDownList, x60 y95 w100 h200 vPreset gPreset, 6||8|10|16|20|22|24|32|Heads or Tails
Gui, 1:Tab, 2, 1
Gui, 1:Add, Edit, x200 y25 w175 h20 vLine
Gui, 1:Add, ListBox, x5 y30 w190 h260 vOptions, % Info
Gui, 1:Add, Button, x200 y50 w75 h23 gAddList Default, Add
Gui, 1:Add, Button, x200 y75 w75 h23 gRemoveList, Remove
Gui, 1:Add, Button, x200 y100 w75 h23 gResetList, Reset
Gui, 1:Show, w380 h290
Return

GuiClose:
Close:
	ExitApp

Preset:
	Gui, Submit, NoHide
	If (Preset = "Heads or Tails") {
		For Each, Item in StrSplit("Min|Max|Min_|Max_", "|")
			GuiControl, Disable, % Item
	} Else {
		For Each, Item in StrSplit("Min|Max|Min_|Max_", "|")
			GuiControl, Enable, % Item
		GuiControl,, Min, 1
		GuiControl,, Max, % Preset
	}
	Return

Roll:
	Gui, Submit, NoHide
	Gui, 1:+OwnDialogs
	If (Tabs = "Option Chooser") {
		ControlGet, All_CT, List, Count, ListBox1, A
		If (All_CT = "")
			Return
		Temp := []
		All_CT__ := StrReplace(All_CT, "`n", "|", MaxIndex)
		Loop, Parse, All_CT__, |
			Temp.Push(A_LoopField)
		MsgBoxEx("Choosing...", "RTD",, 5,,, 2, 1)
		MsgBoxEx("I picked '" StrReplace(Temp[Random(MaxIndex - 1, 1)], "`r`n") "!'", "RTD",,,,,, 1)
	} Else {
		If (Preset = "Heads or Tails") {
			MsgBoxEx("Flipping...", "RTD",, 5,,, 2, 1)
			MsgBoxEx("I got... " Coin[Random(2, 1)] "!", "RTD",,,,,, 1)
		} Else {
			MsgBoxEx("Rolling...", "RTD",, 5,,, 2, 1)
			MsgBoxEx("I got... a " Random(Max, Min) "!", "RTD",,,,,, 1)
		}
	}
	Return

AddList:
	Gui, Submit, NoHide
	If (Line = "") || (Line = A_Space) || (Line = A_Space)
		Return
	GuiControl,, Options, % Line
	GuiControl,, Line, % ""
	Return

RemoveList:
	Gui, Submit, NoHide
	GuiControlGet, Out__,, Options
	ControlGet, All_CT, List, Count, ListBox1, A
	All_CT := StrReplace(All_CT, "`n", "|")
	If (Out__ ~= "|.+$")
		Out := "|" StrReplace(All_CT, Out__,,, 1)
	Else
		Out := "|" StrReplace(All_CT, Out__ "|")
	Out := StrReplace(Out, "||", "|")
	GuiControl,, Options, % Out
	;MsgBoxEx(Out)
	Return

ResetList:
	GuiControl,, Options, % "|"
	Info := ""
	Return

Reset:
	Reload

Open:
	If (InStr(A_ThisMenuItem, "File")) {
		Out := SelectFileDLG(["Document Files (*.txt;*.db;*.doc)", "All Files (*.*)"])
		If ((Data := FileOpen(Out, "r").Read()) = "") || (Out = False)
			Return
		GoSub, ResetList
		For Each, Line in StrSplit(Data, "`n")
			Info .= Line "|"
		GuiControl,, Options, % "|" Info
	} Else If (InStr(A_ThisMenuItem, "Clipboard")) {
		GoSub, ResetList
		For Each, Line in StrSplit(Clipboard, "`n")
			Info .= Line "|"
		GuiControl,, Options, % "|" Info
	}
	Return

; A 'true' randomizer function (Not Really)
Random(Min, Max) {
	; Get the new seed to use
	Random, __, % Min, % Max
	Random,, % __
	; Then, use the new seed to create the effect of random
	Random, Out, % Min, % Max
	Return, Out
}

SelectFileDLG(filters, initialDir := "", DefaultExt := "", HWND := "", Title := "") {
	VarSetCapacity(OPENFILENAMEW, (cbOFN := A_PtrSize == 8 ? 152 : 88), 0)
	NumPut(cbOFN, OPENFILENAMEW,, "UInt") ; lStructSize
	NumPut(HWND := !HWND ? WinExist("A") : HWND, OPENFILENAMEW, A_PtrSize, "Ptr") ; hwndOwner

	FinalFilterString := ""
	For _, Filter in Filters {
		Filter__ := RegExReplace(Filter, "(.|\s)+\(")
		Filter__ := StrReplace(Filter__, ")")
		FinalFilterString .= Filter "|" Filter__ "|"
	}

	while ((char := DllCall("ntdll\wcsrchr", "Ptr", &finalFilterString, "UShort", Asc("|"), "CDecl Ptr")))
		NumPut(0, char+0,, "UShort")

	NumPut(&finalFilterString, OPENFILENAMEW, A_PtrSize*3, "Ptr") ; lpstrCustomFilter
	NumPut(1, OPENFILENAMEW, A_PtrSize*(5 + (A_PtrSize == 4)), "UInt") ; nFilterIndex

	max_path := 260 ; if keeping the option to select multiple files, consider raising the size
	vPath := !DefaultExt ? "*.*" : DefaultExt
	VarSetCapacity(vPath, (max_path+2)*2, 0)
	NumPut(&vPath, OPENFILENAMEW, A_PtrSize*(6 + (A_PtrSize == 4)), "Ptr") ; lpstrFile
	NumPut(max_path, OPENFILENAMEW, A_PtrSize*(7 + (A_PtrSize == 4)), "UInt") ; nMaxFile

	NumPut(&(initialDir), OPENFILENAMEW, A_PtrSize*(10 + (A_PtrSize == 4)), "Ptr") ; lpstrInitialDir
	NumPut(&(Title := !Title ? "Select a File" : Title), OPENFILENAMEW, A_PtrSize*(11 + (A_PtrSize == 4)), "Ptr") ; lpstrTitle
	;OFN_ENABLESIZING := 0x00800000
	;OFN_EXPLORER := 0x00080000
	;OFN_CREATEPROMPT := 0x00002000
	;OFN_ALLOWMULTISELECT := 0x00000200
	;OFN_ENABLEHOOK := 0x00000020
	;OFN_HIDEREADONLY := 0x00000004
	vOFNFlags := 0x00880024
	NumPut(vOFNFlags, OPENFILENAMEW, A_PtrSize*(12 + (A_PtrSize == 4)), "UInt") ; Flags

	if (DllCall("comdlg32\GetOpenFileNameW", "Ptr", &OPENFILENAMEW)) {
		dirOrFile := StrGet(&vPath,, "UTF-16")
		if (!NumGet(vPath, (StrLen(dirOrFile) + 1) * 2, "UShort")) {
			info := dirOrFile
		} else {
			; Multiple files selected
			fileNames := &vPath + (NumGet(OPENFILENAMEW, A_PtrSize == 8 ? 100 : 56, "UShort") * 2)
			while (*fileNames) {
				info .= dirOrFile . "\" . StrGet(fileNames,, "UTF-16")
				fileNames += (DllCall("ntdll\wcslen", "Ptr", fileNames, "CDecl Ptr") * 2) + 2
			}
		}
	}
	DllCall("GlobalFree", "Ptr", cb, "Ptr")
	Return info ? info : False
}

MsgBoxEx(Text
		, Title := ""
		, Buttons := ""
		, Icon := ""
		, ByRef CheckText := ""
		, Styles := ""
		, Timeout := ""
		, Owner := ""
		, FontOptions := ""
		, FontName := ""
		, BGColor := ""
		, Callback := "") {
	Static hWnd, y2, p, px, pw, c, cw, cy, ch, f, o, gL, hBtn, lb, DHW, ww, Off, k, v, RetVal
	Static Sound := {2: "*48", 4: "*16", 5: "*64"}

	Gui N_:New, hWndhWnd LabelMsgBoxEx -0xA0000 -DPIScale
	Gui % (Owner) ? "+Owner" . Owner : ""
	Gui N_:Font
	Gui N_:Font, % (FontOptions) ? FontOptions : "s9", % (FontName) ? FontName : "Segoe UI"
	Gui N_:Color, % (BGColor) ? BGColor : "White"
	Gui N_:Margin, 10, 12

	If (IsObject(Icon)) {
		Gui N_:Add, Picture, % "x20 y24 w32 h32 Icon" . Icon[1], % (Icon[2] != "") ? Icon[2] : "shell32.dll"
	} Else If (Icon + 0) {
		Gui N_:Add, Picture, x20 y24 Icon%Icon% w32 h32, user32.dll
		SoundPlay % Sound[Icon]
	}

	Gui N_:Add, Link, % "x" . (Icon ? 65 : 20) . " y" . (InStr(Text, "`n") ? 24 : 32) . " vc", %Text%
	GuicontrolGet c, Pos
	GuiControl Move, c, % "w" . (cw + 30)
	y2 := (cy + ch < 52) ? 90 : cy + ch + 34

	Gui N_:Add, Text, vf -Background ; Footer

	Gui N_:Font
	Gui N_:Font, s9, Segoe UI
	px := 42
	If (CheckText != "") {
		CheckText := StrReplace(CheckText, "*",, ErrorLevel)
		Gui N_:Add, CheckBox, vCheckText x12 y%y2% h26 -Wrap -Background AltSubmit Checked%ErrorLevel%, %CheckText%
		GuicontrolGet p, Pos, CheckText
		px := px + pw + 10
	}

	o := {}
	Loop Parse, Buttons, |, *
	{
		gL := (Callback != "" && InStr(A_LoopField, "...")) ? Callback : "MsgBoxExBUTTON"
		Gui Add, Button, hWndhBtn g%gL% x%px% w90 y%y2% h26 -Wrap, %A_Loopfield%
		lb := A_LoopField
		o[hBtn] := px
		px += 98
	}
	GuiControl +Default, % (RegExMatch(Buttons, "([^\*\|]*)\*", Match)) ? Match1 : StrSplit(Buttons, "|")[1]

	Gui N_:Show, Autosize Center Hide, %Title%
	DHW := A_DetectHiddenWindows
	DetectHiddenWindows On
	WinGetPos,,, ww,, ahk_id %hWnd%
	GuiControlGet p, Pos, %lb% ; Last button
	Off := ww - (px + pw)
	For k, v in o {
		GuiControl Move, %k%, % "x" . (v + Off - 14)
	}
	Guicontrol MoveDraw, f, % "x-1 y" . (y2 - 10) . " w" . ww . " h" . 48

	Gui N_:Show
	Gui N_:+SysMenu %Styles%
	DetectHiddenWindows %DHW%

	If (Timeout) {
		SetTimer MsgBoxExTIMEOUT, % Round(Timeout) * 1000
	}

	If (Owner) {
		WinSet Disable,, ahk_id %Owner%
	}

	GuiControl Focus, f
	Gui N_:Font
	WinwaitClose ahk_id %hWnd%
	Return RetVal

	MsgBoxExESCAPE:
	MsgBoxExCLOSE:
	MsgBoxExTIMEOUT:
	MsgBoxExBUTTON:
		SetTimer MsgBoxExTIMEOUT, Delete

		If (A_ThisLabel == "MsgBoxExBUTTON") {
			RetVal := StrReplace(A_GuiControl, "&")
		} Else {
			RetVal := (A_ThisLabel == "MsgBoxExTIMEOUT") ? "Timeout" : "Cancel"
		}

		If (Owner) {
			WinSet Enable,, ahk_id %Owner%
		}

		Gui N_:Submit
		Gui N_:Destroy
	Return
}
