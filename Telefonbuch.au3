#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Telefon.ico
#AutoIt3Wrapper_Outfile=c:\temp\Telefonbuch_v3.4.Exe
#AutoIt3Wrapper_Compression=0
#AutoIt3Wrapper_Res_Comment=Zur Funktion ist eine Verbindung ins Active Directory notwendig.
#AutoIt3Wrapper_Res_Description=Telefonbuch
#AutoIt3Wrapper_Res_Fileversion=3.4.0.0
#AutoIt3Wrapper_Res_Language=1031
#AutoIt3Wrapper_Run_Tidy=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; *** Start added by AutoIt3Wrapper ***
#include <StructureConstants.au3>
; *** End added by AutoIt3Wrapper ***

#include <AD.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ListViewConstants.au3>
#include <GuiListView.au3>
#include <Misc.au3>
#include <MsgBoxConstants.au3>

Local $hDLL = DllOpen("user32.dll")

$Defaultstatustext = "Name/Firma/Telefonnummer/Handynummer eingeben und mit Eingabetaste suchen." & @CRLF & "Doppelklick auf eine Zeile wählt die Telefonnummer." & @CRLF & "Doppelklick auf die Handynummer wählt diese, sofern vorhanden." & @CRLF & "Shift+Klick lässt ein Ändern der Nummer vor der Wahl zu"
#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Telefonbuch", 751, 401, 192, 124)
$Input1 = GUICtrlCreateInput("Name/Firma", 32, 16, 233, 21)
$ListView1 = GUICtrlCreateListView("Anzeigename|Firma|Telefon|Handy", 32, 80, 693, 305, $LVS_SINGLESEL, $LVS_EX_FULLROWSELECT)
$Label1 = GUICtrlCreateLabel($Defaultstatustext, 300, 16, 400, 57)
$Icon1 = GUICtrlCreateButton("Über", 40, 40, 32, 32)
$Icon2 = GUICtrlCreateButton("Hilfe", 80, 40, 32, 32)

#EndRegion ### END Koda GUI section ###



$click_id = GUICtrlCreateDummy()
$dblclick_id = GUICtrlCreateDummy()
GUIRegisterMsg($WM_NOTIFY, "WM_Notify_Events")
$Resultcount = 0

GUISetState(@SW_SHOW)
GUICtrlSetState($Input1, $GUI_FOCUS)


While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Input1
			GUICtrlSetData($Label1, "Suche läuft...")
			_GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($ListView1))
			$Searchstring = GUICtrlRead($Input1)
			_AD_Open()
			$aData = _AD_GetObjectsInOU("", "(&(|(telephonenumber=*)(mobile=*))(|(displayname=*" & $Searchstring & "*)(company=*" & $Searchstring & "*)(telephonenumber=*" & $Searchstring & "*)(Mobile=*" & $Searchstring & "*)))", 2, "Displayname,Company,telephonenumber,Mobile")
			_AD_Close()
			GUICtrlSetState($Input1, $GUI_FOCUS)
			GUICtrlSetData($Label1, "Suche läuft...")
			$Resultcount = UBound($aData) - 1
			For $i = 1 To UBound($aData) - 1
				GUICtrlCreateListViewItem($aData[$i][0] & "|" & $aData[$i][1] & "|" & $aData[$i][2] & "|" & $aData[$i][3], $ListView1)
				GUICtrlSetData($Label1, "Es werden " & $i & " von " & $Resultcount & " Ergebnisse(n) angezeigt.")
			Next
			_GUICtrlListView_SetColumnWidth($ListView1, 0, $LVSCW_AUTOSIZE)
			_GUICtrlListView_SetColumnWidth($ListView1, 1, $LVSCW_AUTOSIZE)
			_GUICtrlListView_SetColumnWidth($ListView1, 2, $LVSCW_AUTOSIZE)
			_GUICtrlListView_SetColumnWidth($ListView1, 3, $LVSCW_AUTOSIZE)


			If $Resultcount < 1 Then
				GUICtrlSetData($Label1, "Es wurden keine Einträge gefunden.")
			EndIf
		Case $Icon1

			MsgBox(0, "Info", "Zur Funktion ist eine Verbindung ins Active Directory notwendig.")
		Case $Icon2

			MsgBox(0, "Hilfe", $Defaultstatustext)
		Case $click_id


			OnClick(GUICtrlRead($click_id))

			ConsoleWrite("_IsPreadfsadfsssed - Shift Key was not pressed." & @CRLF)
		Case $dblclick_id
			OnDoubleClick(GUICtrlRead($dblclick_id))
			GUICtrlSetData($Label1, $Defaultstatustext)

	EndSwitch


WEnd
DllClose($hDLL)

Func OnClick($subitem)
	$item = _GUICtrlListView_GetNextItem($ListView1) ; current selected
	$value = _GUICtrlListView_GetItemText($ListView1, $item, $subitem)
	If $subitem < 2 Then
		$value = _GUICtrlListView_GetItemText($ListView1, $item, 2)
	EndIf
	If $subitem = 3 And $value = "" Then
		$value = _GUICtrlListView_GetItemText($ListView1, $item, 2)
	EndIf
	If $Resultcount > 0 Then
		If _IsPressed("10", $hDLL) Then
			ConsoleWrite("_IsPressed - Shift Key was pressed." & @CRLF)
			Local $modifiednumber = InputBox("Rufnummer wählen", "Rufnummer kann hier vor dem wählen angepasst werden." & @CRLF & "Nach einem Klick auf OK wird gewählt.", $value)
			ConsoleWrite($modifiednumber & @CRLF)
			If $modifiednumber = "" Then
				ConsoleWrite("nichts zu wählen" & @CRLF)
			Else
				$iPID = ShellExecute("tel:" & StringRegExpReplace($modifiednumber, "[^0-9+]", ""), "", @SW_SHOWMAXIMIZED)

			EndIf
		Else
			ConsoleWrite("_IsPressed - Shift Key was not pressed." & @CRLF)
		EndIf
	EndIf
	ConsoleWrite('Click: index:' & $item & ' subitem:' & $subitem & ' value:' & $value & @CRLF)
EndFunc   ;==>OnClick

Func OnDoubleClick($subitem)
	$item = _GUICtrlListView_GetNextItem($ListView1) ; current selected
	$value = _GUICtrlListView_GetItemText($ListView1, $item, $subitem)
	If $subitem < 2 Then
		$value = _GUICtrlListView_GetItemText($ListView1, $item, 2)
	EndIf
	If $subitem = 3 And $value = "" Then
		$value = _GUICtrlListView_GetItemText($ListView1, $item, 2)
	EndIf

	ConsoleWrite('DoubleClick: index:' & $item & ' subitem:' & $subitem & ' value:' & $value & @CRLF)
	$iPID = ShellExecute("tel:" & StringRegExpReplace($value, "[^0-9+]", ""), "", @SW_SHOWMAXIMIZED)
EndFunc   ;==>OnDoubleClick

Func WM_Notify_Events($hWndGUI, $MsgID, $wParam, $lParam)
	#forceref $hWndGUI, $MsgID, $wParam

	Local $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
	Local $hWndFrom = DllStructGetData($tNMHDR, "hWndFrom")
	Local $nNotifyCode = DllStructGetData($tNMHDR, "Code")

	If $wParam = $ListView1 Then
		If $nNotifyCode = $NM_CLICK Then
			$tInfo = DllStructCreate($tagNMITEMACTIVATE, $lParam)
;~             $item = DllStructGetData($tInfo, "Index")
			$subitem = DllStructGetData($tInfo, "SubItem")
;~             OnClick($item,$subitem)
			GUICtrlSendToDummy($click_id, $subitem)
		EndIf

		If $nNotifyCode = $NM_DBLCLK Then
			$tInfo = DllStructCreate($tagNMITEMACTIVATE, $lParam)
;~             $item = DllStructGetData($tInfo, "Index")
			$subitem = DllStructGetData($tInfo, "SubItem")
;~             OnDoubleClick($item,$subitem)
			GUICtrlSendToDummy($dblclick_id, $subitem)
		EndIf
	EndIf
EndFunc   ;==>WM_Notify_Events
