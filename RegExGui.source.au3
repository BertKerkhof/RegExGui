; ++ RegExGui ++ Regular expressions test utility ++

; Original idea by w0uter
; Modified by steve8tch

; 2008 update:
; newer PCRE implementation
; removed message boxes and splash screens
; added status bar
; added timer

; 2011 update:
; Help button / compiled scripts repaired
; x64 OS aware.
; Different separator char in GUICtrlSetData (thanks FichteFoll)
; Display formating fixed for number of results = 9, or 99, or 999 etc (thanks FichteFoll)

; 2019 remake by Bert Kerkhof
; (1) ComboBox contents save error repaired. Thanks LeaB
; (2) Added menu with six pattern examples. Thanks FritsW
; (3) Adjustable window width facilitates long patterns. Thanks Melba23
; (4) Ruler points to faulty pattern if any
; (5) Added tooltips
; (6) Two-dim presentation of result data
; (7) Added textual explanation of results
; (8) Lime status contrast color. Thanks EmileR
; (9) New alpha channel transparant 256 bits icon
; (10) Detail vs match shift fault in result array is diagnosed
; (11) Works carelessly both with unicode and 8bit input. Thanks DominiqueH
; (12) This source is shorter than the 2011 version

#include <ColorConstants.au3>
#include <FontConstants.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <FileConstants.au3>
#include <AutoItConstants.au3>
#include <GuiConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <Array.au3>
#include <GuiTab.au3>
#include <GuiComboBox.au3>

Global $FileBox, $EditBox, $PatBox, $OffsetBox, $R0, $R1, $R2, $R3, $R4, $FileDisp
Global $ErrDisp, $ExtDisp, $OutDisp, $TimDisp, $StatDisp, $idTab
Global Const $minW = 550, $minH = 620
Global $0Title = "Logical match", $1Title = "Match", $2Title = "Match + detail"
Global $3Title = "All matches", $4Title = "All matches + details"
Global $0Returns = "Returns a logical value", $1Returns = "Returns single array"
Global $3Returns = "Returns array of array"
Global $0Contains = "indicating whether there is a match"
Global $1Contains = "that in first element should contain the full match"
Global $2Contains = "that in first element contains the full match"
Global $2Parts = "and in subsequent elements the requested string parts (Perl/PHP style)"

Func WM_GETMINMAXINFO($hwnd, $Msg, $wParam, $lParam) ; Ensure min gui size
	Local $tMax = DllStructCreate("int;int;int;int;int;int;int;int;int;int", $lParam)
	DllStructSetData($tMax, 7, $minW + 14)
	DllStructSetData($tMax, 8, $minH + 14)
	DllStructSetData($tMax, 9, @DesktopWidth + 16)
	DllStructSetData($tMax, 10, $minH + 14)
	Return 0
	BitAND($hwnd, $Msg, $wParam, $lParam) ; Satisfies AutoIT3 check -w 5
EndFunc   ;==>WM_GETMINMAXINFO

Func WordWrap($S, $nLimit = 88)
	Local $iFound, $sR = ""
	While StringLen($S) > $nLimit
		$iFound = StringInStr(StringLeft($S, $nLimit), " ", 0, -1)
		If $iFound Then
			$sR &= StringLeft($S, $iFound - 1) & @CRLF
			$S = StringMid($S, $iFound + 1)
			ContinueLoop
		EndIf
		$iFound = StringInStr($S & " ", " ")
		$sR &= StringMid($S, $iFound - 1) & @CRLF
		$S = StringMid($S, $iFound + 1)
	WEnd
	Return StringLen($S) ? $sR & $S : StringTrimRight($sR, 2) ; Strip last @CRLF
EndFunc   ;==>WordWrap

Func RegExTest($InBox)
	; Read text:
	Local $sInText = GUICtrlRead($InBox)
	Local $sUTF = ""
	For $I = 1 To StringLen($sInText)
		If AscW(StringMid($sInText, $I, 1)) > 255 Then
			$sUTF = "(*UTF)"
			ExitLoop
		EndIf
	Next

	; Read options:
	Local $nOffset = GUICtrlRead($OffsetBox)
	$nOffset = @error ? 1 : Int($nOffset)
	Local $aMode = [GUICtrlRead($R0), GUICtrlRead($R1), GUICtrlRead($R2), GUICtrlRead($R3), GUICtrlRead($R4)]
	Local $Mode = _ArraySearch($aMode, 1)

	Local $aaA, $hTimer = TimerInit()
	$aaA = StringRegExp($sUTF & $sInText, GUICtrlRead($PatBox), $Mode, $nOffset)
	Local $Err = @error, $Ext = @extended
	GUICtrlSetData($TimDisp, Round(TimerDiff($hTimer), 2) & " ms")

	; Output:
	GUICtrlSetData($ErrDisp, $Err)
	GUICtrlSetData($ExtDisp, $Ext)
	Select ; StatusDisplay:
		Case $Err = 1
			GUICtrlSetData($StatDisp, "Array is invalid - No matches")
			Return GUICtrlSetBkColor($StatDisp, $COLOR_RED)
		Case $Err = 2
			GUICtrlSetData($StatDisp, "Bad pattern @extended = offset of error in pattern")
			Return GUICtrlSetBkColor($StatDisp, $COLOR_RED)
	EndSelect
	GUICtrlSetData($StatDisp, "Complete")
	GUICtrlSetBkColor($StatDisp, $COLOR_LIME)
	Local $sResult, $sTitle, $sExplain, $Alinea = @CRLF & @CRLF, $S = " "
	Switch $Mode
		Case 0 ; Logical match
			$sTitle = $0Title
			$sExplain = $0Returns & $S & $0Contains
		Case 1 ; Match
			$sTitle = $1Title
			$sExplain = $1Returns & $S & $1Contains & ". However"
		Case 2 ; Match + detail
			$sTitle = $2Title
			$sExplain = $1Returns & $S & $2Contains & $S & $2Parts
		Case 3 ; All matches
			$sTitle = $3Title
			$sExplain = $1Returns & $S & $1Contains & ". However"
		Case 4 ; All matches + details
			$sTitle = $4Title
			$sExplain = $3Returns & $S & $2Contains & $S & $2Parts
	EndSwitch
	$sResult = $sTitle & $Alinea & WordWrap($sExplain, 64) & ":" & $Alinea
	If Not IsArray($aaA) Then ; Result -> Single value
		$sResult &= $aaA ? "True = match found" : "False = no match found"
	ElseIf IsArray($aaA[0]) Then ; Result -> array of arrays
		For $I = 0 To UBound($aaA) - 1
			$sResult &= $I & " => ["
			Local $J, $aA = $aaA[$I]
			For $J = 0 To UBound($aA) - 1
				$sResult &= '"' & $aA[$J] & '", '
			Next
			$sResult = StringTrimRight($sResult, 2) & "]" & @CRLF
		Next
	Else ; Result -> array
		For $I = 0 To UBound($aaA) - 1
			$sResult &= $I & ' => "' & $aaA[$I] & '"' & @CRLF
		Next
	EndIf
	GUICtrlSetData($OutDisp, $sResult)
EndFunc   ;==>RegExTest

Func Example($aExample)
	Local $sName = $aExample[0], $aM[6] ; Merge
	$aM[0] = "Data to take in account such as the " & StringLower($sName)
	$aM[1] = "may come in different manners. There is a need to put"
	$aM[2] = "and"
	$aM[3] = "in an orderly fashion: a fact sheet. Formats used for data such as"
	$aM[4] = "may vary per country. Edit the template as needed and add"
	$aM[5] = "to your " & $sName & " list."
	Local $I, $sMerge = "", $aTest = $aExample[2]
	For $I = 0 To 5
		$sMerge &= $aM[$I] & " " & $aTest[$I] & " "
	Next
	Local $sResult = $sName & "s example" & @CRLF & @CRLF & WordWrap($sMerge) & @CRLF
	GUICtrlSetData($EditBox, $sResult)
	_GUICtrlTab_ActivateTab($idTab, 0) ; Ensure editbox is visible
	_GUICtrlComboBox_SetEditText($PatBox, $aExample[1])
	GUICtrlSetState($R4, $GUI_CHECKED) ; All matches + details
	GUICtrlSetData($OffsetBox, 1)
	RegExTest($EditBox)
EndFunc   ;==>Example

Func Main()
	Opt("GUIDataSeparatorChar", Chr(11)) ; VT sep in combo

	; Example patterns:
	Local $P0 = "([123]?\d*) (jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec) (\d\d(?:\d\d)*)" ; Date
	Local $P1 = "(\d{3,4}(?:\s+|-|.)\d{6,7})" ; Phone#
	Local $P2 = "(\p{Lu}\p{Ll}+)\s+(\p{Lu}\p{Ll}+)" ; Name
	Local $P3 = "([012]\d)(?:\:|h)([012345]\d)?(?:\:|m)?([012345]\d)?s?" ; Time
	Local $P4 = "((?:[\w_\-\.])+)@([\w-]+.(?:com|nl|org))" ; Mail address
	Local $P5 = "(?:https?://)?+(?:[\w\-]+\.)*?([\w\-]+\.(info|org|com))" ; Website

	; Data:
	Local $aD0[6] = ["4 jan 2020", "17 apr 1997", "14 feb 2019", "25 dec 2021", "15 jul 2019"] ; Date
	Local $aD1[6] = ["030-5621898", "0592 873567", "0912.639673", "044 6423672", "0592-338080"] ; Phone#
	Local $aD2[6] = ["Piet Oppers", "Ed Jansz", "Zoë Alhan", "Joep Dorwerth", "Bert Kerkhof"] ; Name
	Local $aD3[6] = ["12:43:16", "18h35", "07:30", "10h", "15h30m17s"] ; Time
	Local $aD4[6] = ["v.siebel@mail.com", "a_marie@live.com", "julia@hot.com", "sahim@alhan.nl", "kerkhof.bert@gmail.com"] ; Mail address
	Local $aD5[6] = ["www.pcre.org", "fairfood.org", "https://www.danstools.com/", "regular-expressions.info", "childinstress.com"] ; Website

    ; Combine:
    Local $aX0 = ["Date", $P0, $aD0], $aX1 = ["Phone number", $P1, $aD1]
	Local $aX2 = ["Name", $P2, $aD2], $aX3 = ["Time", $P3, $aD3]
	Local $aX4 = ["Mail adress", $P4, $aD4], $aX5 = ["Website", $P5, $aD5]
	Local $aaExample[6] = [$aX0, $aX1, $aX2, $aX3, $aX4, $aX5]

	; Gui:
	Local Const $Style = $ES_WANTRETURN + $WS_VSCROLL + $ES_AUTOVSCROLL + $ES_AUTOHSCROLL
	Local Const $SOFTYELLOW = 0xFBFFC6, $LIGHTGRAY = 0xD0D0D0
	Local Const $sTitle = "RegEx ++ Education Edition by BertK ++"
	Local $W = $minW, $H = $minH
	Local $hGui = GUICreate($sTitle, $W, $H, -1, -1, $WS_SIZEBOX)
	GUISetFont(10, 500, 0, "Calibri", $hGui)

	; Menu:
	Local $I, $aXid[6], $mExample = GUICtrlCreateMenu("Examples")

	For $I = 0 To UBound($aaExample) - 1
		$aXid[$I] = GUICtrlCreateMenuItem(($aaExample[$I])[0], $mExample)
	Next
	Local $mHelp = GUICtrlCreateMenu("Help")
	Local $iHelp = GUICtrlCreateMenuItem("AutoIT Help", $mHelp)

	; Input panel:
	$idTab = GUICtrlCreateTab(10, 10, 530, 203)
	GUICtrlSetResizing($idTab, $GUI_DOCKAUTO)
	GUICtrlCreateTabItem("Copy and Paste the text to check")
	$EditBox = GUICtrlCreateEdit("", 15, 36, 518, 170, $Style)
	GUICtrlSetBkColor(-1, $SOFTYELLOW)
	GUICtrlCreateTabItem("Load text from File")
	Local $idBrowse = GUICtrlCreateButton("Browse for file", 15, 36, 97, 20)
	$FileDisp = GUICtrlCreateEdit("", 118, 36, 415, 20)
	$FileBox = GUICtrlCreateEdit("", 15, 61, 518, 145, $Style)
	GUICtrlSetBkColor(-1, $SOFTYELLOW)
	GUICtrlCreateTabItem("") ;

	; Pattern panel:
	GUICtrlCreateGroup("The Pattern", 10, 217, 530, 115)
	Local $rU = "12345678901234567890", $rT = "", $rS = ""
	For $I = 1 To 20
		$rT &= StringLeft("         ", 10 - StringLen(String($I))) & $I
		$rS &= $rU
	Next
	Local $idRuler = GUICtrlCreateLabel($rT & @LF & $rS, 20, 235, 494, 35, $SS_LEFTNOWORDWRAP)
	GUICtrlSetFont(-1, 10, $FW_SEMIBOLD, 0, "Courier New")
	GUICtrlSetColor(-1, $LIGHTGRAY)

	$PatBox = GUICtrlCreateCombo("", 15, 267, 520, 15)
	GUICtrlSetFont(-1, 10, $FW_SEMIBOLD, 0, "Courier New")
	GUICtrlSetColor(-1, $COLOR_BLUE)
	GUICtrlSetBkColor(-1, $COLOR_YELLOW)
	Local Const $iniFile = @ScriptDir & "\RegExGui.ini"
	Local $sEdit = IniRead($iniFile, @UserName, "Patterns", "")
	GUICtrlSetData($PatBox, $sEdit) ; Initiate

	Local $idTest = GUICtrlCreateButton("Test", 14, 300, 60, 25)
	Local $idAddPat = GUICtrlCreateButton("Add", 412, 300, 60, 25)
	Local $sTip = "in the inifile together with the RegEx program"
	GUICtrlSetTip(-1, $sTip, "Save your success patterns", $TIP_INFOICON, $TIP_BALLOON)
	Local $idDelPat = GUICtrlCreateButton("Del", 476, 300, 60, 25)

	; Options panel:
	GUICtrlCreateGroup("Options", 10, 342, 141, 146)
	$R0 = GUICtrlCreateRadio($0Title, 15, 361, 132, 18)
	GUICtrlSetState($R0, $GUI_CHECKED)
	GUICtrlSetTip(-1, $0Contains, $0Returns, $TIP_INFOICON, $TIP_BALLOON)
	$R1 = GUICtrlCreateRadio($1Title, 15, 379, 132, 18)
	GUICtrlSetTip(-1, $1Contains, $1Returns, $TIP_INFOICON, $TIP_BALLOON)
	$R2 = GUICtrlCreateRadio($2Title, 15, 397, 132, 18)
	Local $sRap = WordWrap($2Contains & " " & $2Parts, 64)
	GUICtrlSetTip(-1, $sRap, $1Returns, $TIP_INFOICON, $TIP_BALLOON)
	$R3 = GUICtrlCreateRadio($3Title, 15, 415, 132, 18)
	GUICtrlSetTip(-1, $1Contains, $1Returns, $TIP_INFOICON, $TIP_BALLOON)
	$R4 = GUICtrlCreateRadio($4Title, 15, 433, 132, 18)
	GUICtrlSetTip(-1, $sRap, $3Returns, $TIP_INFOICON, $TIP_BALLOON)

	$OffsetBox = GUICtrlCreateInput("1", 15, 457, 40, 20)
	$sTip = "Position to start the search"
	GUICtrlSetTip(-1, $sTip & @LF & "Default is 1", "Offset option", $TIP_INFOICON, $TIP_BALLOON)
	GUICtrlCreateLabel("Offset", 58, 459, 40, 15)

	GUICtrlCreateLabel("@error", 15, 496, 70, 15)
	$ErrDisp = GUICtrlCreateInput("", 15, 512, 40, 20, $ES_READONLY)
	GUICtrlCreateLabel("@extended", 65, 496, 70, 15)
	$ExtDisp = GUICtrlCreateInput("", 65, 512, 62, 20, $ES_READONLY)
	$sTip = "this field points to the location in the pattern that caused it"
	GUICtrlSetTip(-1, $sTip, "On any error", $TIP_INFOICON, $TIP_BALLOON)

	; Output panel:
	GUICtrlCreateLabel("Output", 160, 342, 70, 15)
	$OutDisp = GUICtrlCreateEdit("", 158, 358, 382, 174, $Style)
	GUICtrlSetBkColor(-1, $SOFTYELLOW)

	; Statbar:
	$TimDisp = GUICtrlCreateLabel("Time (ms)", 10, 551, 70, 20, $SS_SUNKEN)
	$StatDisp = GUICtrlCreateLabel("Ready..", 88, 551, 452, 20, $SS_SUNKEN)

	; Main process:
	Local $iDropDown, $InBox = $EditBox ; Holds the selected input box
	GUICtrlSetState($EditBox, $GUI_FOCUS)
	GUISetState(@SW_SHOW)
	GUIRegisterMsg($WM_GETMINMAXINFO, "WM_GETMINMAXINFO") ; Ensure min gui size
	While True
		Local $Msg = GUIGetMsg()
		Local $iFound = _ArraySearch($aXid, $Msg)
		If $iFound > -1 Then
			Example($aaExample[$iFound])
			ContinueLoop
		EndIf
		Switch $Msg
			Case $idTab
				$InBox = GUICtrlRead($idTab) ? $FileBox : $EditBox
			Case $idBrowse
				Local $sFolder = @ScriptDir ; Default file browse folder
				Local $sFilePath = FileOpenDialog("Select text file", $sFolder, "Text files (*.*)", $FD_FILEMUSTEXIST)
				$sFolder = StringTrimRight($sFilePath, StringInStr($sFilePath, "\"))
				GUICtrlSetData($FileDisp, $sFilePath)
				GUICtrlSetData($FileBox, FileRead($sFilePath))
			Case $idTest
				RegExTest($InBox)
			Case $idAddPat
				$sEdit = _GUICtrlComboBox_GetEditText($PatBox)
				$iDropDown = _GUICtrlComboBox_FindStringExact($PatBox, $sEdit)
				If $sEdit = "" Or $iDropDown > -1 Then ContinueLoop ; Item already exists
				_GUICtrlComboBox_AddString($PatBox, $sEdit)
				IniWrite($iniFile, @UserName, "Patterns", _GUICtrlComboBox_GetList($PatBox))
				GUICtrlSetState($PatBox, $GUI_FOCUS)
			Case $idDelPat
				$sEdit = _GUICtrlComboBox_GetEditText($PatBox)
				If $sEdit = "" Then ContinueLoop
				$iDropDown = _GUICtrlComboBox_FindStringExact($PatBox, $sEdit)
				If $iDropDown > -1 Then _GUICtrlComboBox_DeleteString($PatBox, $iDropDown)
				IniWrite($iniFile, @UserName, "Patterns", _GUICtrlComboBox_GetList($PatBox))
				GUICtrlSetState($PatBox, $GUI_FOCUS)
			Case $iHelp ; Supports compilation in x32/x64 default mode:
				Run(@ProgramFilesDir & "\AutoIt3\AutoIt3Help.exe StringRegExp")
				If @error Then MsgBox($MB_SYSTEMMODAL, "error", "Help file not found")
			Case $GUI_EVENT_RESIZED ; Sync ruler:
				GUICtrlSetPos($idRuler, 20 + ((WinGetClientSize($hGui))[0] - $minW) / 36)
			Case $GUI_EVENT_CLOSE
				Exit
		EndSwitch
	WEnd
EndFunc   ;==>Main

Main()
