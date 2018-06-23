#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ObjMarker.ico
#AutoIt3Wrapper_Outfile_x64=ObjMark.exe
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Fileversion=2.1.3
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;version history
;0.5 (2008-03-13)
;	* created an idea with no changing parameters for marking Helios-IT objects
;	* just marks objects one-by-one
;1.0 (2008-03-14)
;	* first "public" release
;1.2 (2008-03-17)
;	+ delay option
;	+ timer option
;	* changed some messages
;2.0 (2008-03-25)
;	* almost all is new functionality
;2.1 (2009-06-26)
;	+ pages
;2.1.1 (2010-09-07)
;	* fixed focus losing when versionlist length exceeds 80
;2.1.2 (2013-12-06)
;   + queries
;2.1.3 (2018-06-23)
;   * small fixes
;   * added option for specifying target NAV version, used for max length of Version List
;~ TODOs
;~ + add localization
;~     visual
;~     processing with NAV through hotkeys (with localized lang)
;~ * check columns order before proceed
;~ * perform correct modifying of Modified due to localizations and changed columns order
;~ * clear marked objects before proceed
;~ ? Excel file for importing list (which format?)
;~ ? documentation trigger modification implementation
;~   > performing changed date-time setting back
;~      might be through creating codeunit and running it
;~      or simple automated "manual" changing object-by-object
;~ ? redesign
;~ ? wizard mode (might be better idea for checking purpose)
;~ ? adding codeunit which perform modifying Version List and Modified using our tool
;~ ? also, date and time saving and restoring, for this purpose, table should be created through our tool


#include <GUIConstants.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>
#include <UpdownConstants.au3>
#include <WindowsConstants.au3>
#include <Array.au3>

HotKeySet("{PAUSE}","_zzz")
$title = "Object Marker"
$version = "2.1.3"
$inifile =  @ScriptDir & "\ObjMark.ini"
$NAVOldtitle = ""
$NAVtitle = ""
$linescount = 0
Dim $listlines[20000]
Dim $objlist[9][3]; [][0] - object name, [][1] - objects quantity, [][2] - object alt-symbol
Dim $obj[9][20000]
Dim $objects[9][2]
	$objects[0][0] = "COD"
	$objects[0][1] = "c"
	$objects[1][0] = "DAT"
	$objects[1][1] = "o"
	$objects[2][0] = "FOR"
	$objects[2][1] = "m"
	$objects[3][0] = "MEN"
	$objects[3][1] = "s"
	$objects[4][0] = "PAG"
	$objects[4][1] = "g"
	$objects[5][0] = "REP"
	$objects[5][1] = "p"
	$objects[6][0] = "QUE"
	$objects[6][1] = "q"
	$objects[7][0] = "TAB"
	$objects[7][1] = "b"
	$objects[8][0] = "XML"
	$objects[8][1] = "x"
$objtypes = 0

$ListFile = IniRead($inifile,"Settings","ListFile","")
$DelayMin = IniRead($inifile,"Settings","DelayMin",10)
$DelayMax = IniRead($inifile,"Settings","DelayMax",2000)
$Delay = IniRead($inifile,"Settings","Delay",200)
$TextToAddToVL = IniRead($inifile,"Settings","TextToAddToVL","")
$SetModified = IniRead($inifile,"Settings","SetModified",False)
$SetModifiedYes = IniRead($inifile,"Settings","SetModifiedYes",False)
$AddTextToVL = IniRead($inifile,"Settings","AddTextToVL",False)
$NAVVersion2013R2Plus = IniRead($inifile,"Settings","NAVVersion2013r2Plus",False)
$ClearFilter = IniRead($inifile,"Settings","ClearFilter",True)
$SetMarkedOnly = IniRead($inifile,"Settings","SetMarkedOnly",True)
$SetOpenExport = IniRead($inifile,"Settings","SetOpenExport",True)

If $SetModified = "False" Then
	$SetModified = False
Else
	$SetModified = True
EndIf
If $SetModifiedYes = "False" Then
	$SetModifiedYes = False
Else
	$SetModifiedYes = True
EndIf
If $AddTextToVL = "False" Then
	$AddTextToVL = False
Else
	$AddTextToVL = True
EndIf
If $ClearFilter = "False" Then
	$ClearFilter = False
Else
	$ClearFilter = True
EndIf
If $SetMarkedOnly = "False" Then
	$SetMarkedOnly = False
Else
	$SetMarkedOnly = True
EndIf
If $SetOpenExport = "False" Then
	$SetOpenExport = False
Else
	$SetOpenExport = True
EndIf
If $NAVVersion2013R2Plus = "False" Then
	$NAVVersion2013R2Plus = False
Else
	$NAVVersion2013R2Plus = True
EndIf

$FormTop = IniRead($inifile,"Form","Top",100)
$FormLeft = IniRead($inifile,"Form","Left",100)

$Form_objectmarker = GUICreate($title & " v." & $version, 432, 244, $FormLeft, $FormTop)
$Button_about = GUICtrlCreateButton("About", 8, 150, 42, 87, 0)
GUICtrlSetTip($Button_about,"About tool")

$Group_Objects = GUICtrlCreateGroup("Object List", 8, 6, 417, 60)
	$Label_file = GUICtrlCreateLabel("File", 16, 26, 20, 17)
	GUICtrlSetTip($Label_file,"File with list of objects")
	$Input_file = GUICtrlCreateInput($ListFile, 40, 22, 311, 21, BitOR($ES_AUTOHSCROLL,$ES_READONLY))
	GUICtrlSetTip($Input_file,"File with list of objects")
	$i = 0
	If $ListFile = "" Then
		$i = $BS_DEFPUSHBUTTON
	EndIf
	$Button_file = GUICtrlCreateButton("Choose...", 358, 22, 60, 21, $i)
	GUICtrlSetTip($Button_file,"Click to choose file with list of objects")
	If $ListFile = "" Then
		GUICtrlSetState($Button_file, $GUI_FOCUS)
	EndIf
	$Label_objects = GUICtrlCreateLabel("The total number of objects to process: <unknown>", 16, 46, 247, 17)
	GUICtrlSetTip($Label_objects,"Quantity of recognized objects to mark")
GUICtrlCreateGroup("", -99, -99, 1, 1)

$Group_Preparation = GUICtrlCreateGroup("Microsoft Dynamics NAV", 288, 72, 137, 65)
	$Checkbox_clearfilter = GUICtrlCreateCheckbox("Clear All Filters", 296, 88, 118, 17)
	GUICtrlSetTip($Checkbox_clearfilter,"Clear filters that were probably set in Object Designer before proceeding" & @CRLF _
									  & "If set, marking will be applied with bulk method" & @CRLF _
									  & "If not set, marking will be applied slowly object by object")
	$i = 0
	If $ListFile <> "" Then
		$i = $BS_DEFPUSHBUTTON
	EndIf
	$Button_choosewindow = GUICtrlCreateButton("Select NAV Window", 296, 108, 120, 21, 0)
	GUICtrlSetTip($Button_choosewindow,"Need to specify which Microsoft Dynamics NAV Development Environment window you are going to work")
	If $ListFile <> "" Then
		GUICtrlSetState($Button_choosewindow, $GUI_FOCUS)
	EndIf
GUICtrlCreateGroup("", -99, -99, 1, 1)

$Group_Finishing = GUICtrlCreateGroup("Upon Finishing", 7, 72, 185, 65)
	$Checkbox_setmarkedonly = GUICtrlCreateCheckbox("Select Marked Objects Only", 15, 88, 173, 17, BitOR($BS_CHECKBOX,$BS_AUTOCHECKBOX,$BS_TOP,$WS_TABSTOP))
	GUICtrlSetTip($Checkbox_setmarkedonly,"Automatically set filter on Marked Only after finishing mark process")
	$Checkbox_openexport = GUICtrlCreateCheckbox("Open Export Objects Dialog Box", 15, 112, 174, 17, BitOR($BS_CHECKBOX,$BS_AUTOCHECKBOX,$BS_TOP,$WS_TABSTOP))
	GUICtrlSetTip($Checkbox_openexport,"Automatically open Export Objects dialog after finishing mark process" & @CRLF & "Cannot be set without Select Marked Objects Only.")
GUICtrlCreateGroup("", -99, -99, 1, 1)

$Group_Delay = GUICtrlCreateGroup("Delay", 199, 72, 82, 65)
	$Input_delay = GUICtrlCreateInput($Delay, 207, 88, 48, 21, BitOR($ES_RIGHT,$ES_AUTOHSCROLL,$ES_NUMBER))
	GUICtrlSetTip($Input_delay,"Keystrokes emulating delay in milliseconds [10..2000]")
	$Updown_delay = GUICtrlCreateUpdown($Input_delay, BitOR($UDS_ALIGNRIGHT,$UDS_ARROWKEYS,$UDS_NOTHOUSANDS))
	GUICtrlSetLimit($Updown_delay, $DelayMax, $DelayMin)
	$Label_ms = GUICtrlCreateLabel("ms", 258, 91, 17, 17)
	GUICtrlSetTip($Label_ms,"Keystrokes emulating delay in milliseconds")
GUICtrlCreateGroup("", -99, -99, 1, 1)

$Group_Functional = GUICtrlCreateGroup("Additional Functions", 56, 144, 209, 93)
	$Checkbox_addtoVL = GUICtrlCreateCheckbox("Add Text to Version List", 64, 163, 134, 17)
	GUICtrlSetTip($Checkbox_addtoVL,"Add specific text to Version List for marked objects")
	$Input_addtoVL = GUICtrlCreateInput($TextToAddToVL, 200, 160, 59, 21)
	GUICtrlSetTip($Input_addtoVL,"Text that will be added to Version List for marked objects" & @CRLF & "Comma added automatically")
	$Checkbox_NAVVersion248 = GUICtrlCreateCheckbox("NAV Version >= 2013R2 (7.1)", 64, 186, 190, 17)
	GUICtrlSetTip($Checkbox_NAVVersion248,"Choose to define if maximum length of Version List 80 characters (NAV 2013 and older) or 248 characters (NAV 2013R2 and newer)" & @CRLF & "This cannot be determined automatically")
	$Checkbox_modified = GUICtrlCreateCheckbox("Modified", 64, 209, 60, 17)
	GUICtrlSetTip($Checkbox_modified,"Set Modified flag to specific value")
	$Radio_modifiedyes = GUICtrlCreateRadio("Yes", 127, 209, 41, 17)
	GUICtrlSetTip($Radio_modifiedyes,"Set Modified flag to Yes (TRUE)")
	$Radio_modifiedno = GUICtrlCreateRadio("No", 171, 209, 41, 17)
	GUICtrlSetTip($Radio_modifiedno,"Set Modified flag to No (FALSE)")
GUICtrlCreateGroup("", -99, -99, 1, 1)

$Group_Ready = GUICtrlCreateGroup("Readiness Status", 271, 144, 154, 93)
	$Checkbox_objects = GUICtrlCreateCheckbox("Object List File Selected", 278, 160, 139, 19, BitOR($BS_CHECKBOX,$BS_AUTOCHECKBOX))
	GUICtrlSetTip($Checkbox_objects,"Is file with list of objects imported")
	$Checkbox_window = GUICtrlCreateCheckbox("NAV Window Selected", 278, 181, 139, 19, BitOR($BS_CHECKBOX,$BS_AUTOCHECKBOX))
	GUICtrlSetTip($Checkbox_window,"Is Microsoft Dynamics NAV Development Environment window specified")
	$Button_start = GUICtrlCreateButton("Start Marking >>", 278, 205, 139, 24, 0)
	GUICtrlSetTip($Button_start,"Start process of marking objects in selected NAV window")
	GUICtrlSetState($Checkbox_objects, $GUI_DISABLE)
	GUICtrlSetState($Checkbox_window, $GUI_DISABLE)
	GUICtrlSetState($Button_start, $GUI_DISABLE)
GUICtrlCreateGroup("", -99, -99, 1, 1)

	If $ClearFilter = True Then
		GUICtrlSetState($Checkbox_clearfilter, $GUI_CHECKED)
	Else
		GUICtrlSetState($Checkbox_clearfilter, $GUI_UNCHECKED)
	EndIf

	If $SetMarkedOnly = True Then
		GUICtrlSetState($Checkbox_setmarkedonly, $GUI_CHECKED)
	Else
		GUICtrlSetState($Checkbox_setmarkedonly, $GUI_UNCHECKED)
	EndIf

	If $SetOpenExport = True Then
		GUICtrlSetState($Checkbox_openexport, $GUI_CHECKED)
	Else
		GUICtrlSetState($Checkbox_openexport, $GUI_UNCHECKED)
	EndIf

	If $AddTextToVL = True Then
		GUICtrlSetState($Checkbox_addtoVL, $GUI_CHECKED)
		GUICtrlSetState($Input_addtoVL, $GUI_ENABLE)
		GUICtrlSetState($Checkbox_NAVVersion248, $GUI_ENABLE)
	Else
		GUICtrlSetState($Checkbox_addtoVL, $GUI_UNCHECKED)
		GUICtrlSetState($Input_addtoVL, $GUI_DISABLE)
		GUICtrlSetState($Checkbox_NAVVersion248, $GUI_DISABLE)
	EndIf
	If $NAVVersion2013R2Plus = True Then
		GUICtrlSetState($Checkbox_NAVVersion248, $GUI_CHECKED)
	Else
		GUICtrlSetState($Checkbox_NAVVersion248, $GUI_UNCHECKED)
	EndIf

	If $SetModified = True Then
		GUICtrlSetState($Checkbox_modified, $GUI_CHECKED)
		GUICtrlSetState($Radio_modifiedyes, $GUI_ENABLE)
		GUICtrlSetState($Radio_modifiedno, $GUI_ENABLE)
		If $SetModifiedYes = True Then
			GUICtrlSetState($Radio_modifiedyes, $GUI_CHECKED)
			GUICtrlSetState($Radio_modifiedno, $GUI_UNCHECKED)
		Else
			GUICtrlSetState($Radio_modifiedyes, $GUI_UNCHECKED)
			GUICtrlSetState($Radio_modifiedno, $GUI_CHECKED)
		EndIf
	Else
		GUICtrlSetState($Radio_modifiedyes, $GUI_DISABLE)
		GUICtrlSetState($Radio_modifiedno, $GUI_DISABLE)
	EndIf
GUISetState(@SW_SHOW)

If $ListFile <> "" Then
	_LoadFile()
EndIf
While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			If $NAVOldtitle <> "" Then
				WinSetTitle($NAVtitle,"Status Bar",$NAVOldtitle)
			EndIf
			_SaveIni()
			Exit

		Case $Button_about
			$message  = $title & " tool was created by Paul Furlet," & @CRLF
			$message &= @TAB & "using AutoIt scripting lagnuage and editor SciTE." & @CRLF & @CRLF
			$message &= "Thanks to:" & @CRLF
			$message &= @TAB & "Ivan Cherkashin" & @TAB & "for help and testing" & @CRLF
			$message &= @TAB & "Denis Fedorov" & @TAB & "for proofreading" & @CRLF & @CRLF
			$message &= "Communicate with me:" & @CRLF
			$message &= @TAB & "E-Mail" & @TAB & "paul.furlet@gmail.com" & @CRLF
			$message &= @TAB & "Skype" & @TAB & "paul.furlet"
			MsgBox(8192+64, "About " & $title & " v." & $version, $message)

		Case $Button_file
			$message = $title & ". Select file with list of objects."
			$ListFile = FileOpenDialog($message, "", "Text files (*.txt)|All files (*.*)", 1)
			If @error Then
				ContinueLoop
			EndIf
			_LoadFile()
			GUICtrlSetStyle($Button_choosewindow,$BS_DEFPUSHBUTTON)
			GUICtrlSetState($Button_choosewindow,$GUI_FOCUS)
			IniWrite($inifile,"Settings","ListFile",$ListFile)
			_ReadyCheck()

		Case $Button_choosewindow
			If $NAVOldtitle <> "" Then
				WinSetTitle($NAVtitle,"Status Bar",$NAVOldtitle)
			EndIf
			WinActivate("[CLASS:Shell_TrayWnd]")
			GUISetState(@SW_HIDE)
			$message = "Now activate the necessary Microsoft Dynamics NAV window."
			TrayTip($title, $message, 10, 2)
			WinWaitActive("[CLASS:C/SIDE Application]")
			$NAVOldtitle = WinGetTitle("[CLASS:C/SIDE Application]")
			$NAVtitle = $title & " is working with this window. Do not touch the mouse or keyboard during the marking process!"
			WinSetTitle($NAVOldtitle, "Status Bar", $NAVtitle)
			GUISetState(@SW_SHOW)
			$message  = "The Microsoft Dynamics NAV window title has been changed for the correct work of the tool." & @CRLF
			$message &= "It will be restored after the marking process is completed."
			TrayTip($title, $message, 10, 1)
			GUICtrlSetState($Checkbox_window,$GUI_CHECKED)
			_ReadyCheck()

		Case $Checkbox_clearfilter
			If GUICtrlRead($Checkbox_clearfilter) = $GUI_CHECKED Then
				$ClearFilter = True
			Else
				$ClearFilter = False
			EndIf
			IniWrite($inifile,"Settings","ClearFilter",$ClearFilter)

		Case $Checkbox_modified
			If GUICtrlRead($Checkbox_modified) = $GUI_CHECKED Then
				GUICtrlSetState($Radio_modifiedyes,BitOR($GUI_ENABLE,$GUI_CHECKED,$GUI_FOCUS))
				GUICtrlSetState($Radio_modifiedno,BitOR($GUI_ENABLE,$GUI_UNCHECKED))
				$SetModified = True
			Else
				GUICtrlSetState($Radio_modifiedyes,BitOR($GUI_DISABLE,$GUI_UNCHECKED))
				GUICtrlSetState($Radio_modifiedno,BitOR($GUI_DISABLE,$GUI_UNCHECKED))
				$SetModified = False
			EndIf
			IniWrite($inifile,"Settings","SetModified",$SetModified)

		Case $Radio_modifiedyes
			$SetModifiedYes = True
			IniWrite($inifile,"Settings","SetModifiedYes",$SetModifiedYes)

		Case $Radio_modifiedno
			$SetModifiedYes = False
			IniWrite($inifile,"Settings","SetModifiedYes",$SetModifiedYes)

		Case $Checkbox_addtoVL
			If GUICtrlRead($Checkbox_addtoVL) = $GUI_CHECKED Then
				GUICtrlSetState($Input_addtoVL, BitOR($GUI_ENABLE,$GUI_FOCUS))
				GUICtrlSetState($Checkbox_NAVVersion248, $GUI_ENABLE)
				$AddTextToVL = True
			Else
				GUICtrlSetState($Input_addtoVL, $GUI_DISABLE)
				GUICtrlSetState($Checkbox_NAVVersion248, $GUI_DISABLE)
				$AddTextToVL = False
			EndIf
			IniWrite($inifile,"Settings","AddTextToVL",$AddTextToVL)

		Case $Input_addtoVL
			$TextToAddToVL = GUICtrlRead($Input_addtoVL)
			If $TextToAddToVL = "" Then
				GUICtrlSetState($Input_addtoVL, $GUI_DISABLE)
				GUICtrlSetState($Checkbox_addtoVL, $GUI_UNCHECKED)
				$AddTextToVL = False
				IniWrite($inifile,"Settings","AddTextToVL",$AddTextToVL)
			EndIf
			IniWrite($inifile,"Settings","TextToAddToVL",$TextToAddToVL)

		Case $Input_delay
			If Number(GUICtrlRead($Input_delay)) > $DelayMax Then
				GUICtrlSetData($Input_delay, $DelayMax)
			EndIf
			If Number(GUICtrlRead($Input_delay)) < $DelayMin Then
				GUICtrlSetData($Input_delay, $DelayMin)
			EndIf
			$Delay = GUICtrlRead($Input_delay)
			IniWrite($inifile,"Settings","Delay",$Delay)

		Case $Updown_delay
			$Delay = GUICtrlRead($Input_delay)
			IniWrite($inifile,"Settings","Delay",$Delay)

		Case $Checkbox_setmarkedonly
			If GUICtrlRead($Checkbox_setmarkedonly) = $GUI_CHECKED Then
				$SetMarkedOnly = True
			Else
				$SetMarkedOnly = False
				GUICtrlSetState($Checkbox_openexport, $GUI_UNCHECKED)
				$SetOpenExport = False
			EndIf
			IniWrite($inifile,"Settings","SetMarkedOnly",$SetMarkedOnly)
			IniWrite($inifile,"Settings","SetOpenExport",$SetOpenExport)

		Case $Checkbox_openexport
			If GUICtrlRead($Checkbox_openexport) = $GUI_CHECKED Then
				$SetOpenExport = True
				GUICtrlSetState($Checkbox_setmarkedonly, $GUI_CHECKED)
				$SetMarkedOnly = True
			Else
				$SetOpenExport = False
			EndIf
			IniWrite($inifile,"Settings","SetOpenExport",$SetOpenExport)
			IniWrite($inifile,"Settings","SetMarkedOnly",$SetMarkedOnly)

		Case $Checkbox_NAVVersion248
			If GUICtrlRead($Checkbox_NAVVersion248) = $GUI_CHECKED Then
				 $NAVVersion2013R2Plus = True
			Else
				 $NAVVersion2013R2Plus = False
			EndIf
			IniWrite($inifile,"Settings","NAVVersion2013r2Plus",$NAVVersion2013R2Plus)

		Case $Button_start
			Opt("SendKeyDelay",$Delay)
			$message  = "Marking will be processed with the following parameters:" & @CRLF & @CRLF
			$message &= "File: " & @TAB & @TAB & @TAB & $ListFile & @CRLF
			$message &= "Objects to process: " & @TAB & $linescount & @CRLF
			$message &= "Clear all filters: " & @TAB & @TAB & $ClearFilter & @CRLF
			$message &= "Delay: " & @TAB & @TAB & @TAB & $Delay & @CRLF
			If $AddTextToVL = True Then
				If $NAVVersion2013R2Plus = True Then
					$message &= "Target NAV version" & @TAB & @TAB & "2013R2 or newer" & @CRLF
					$message &= @TAB & "Version List len" & @TAB & @TAB & "248" & @CRLF
				Else
					$message &= "Target NAV version" & @TAB & @TAB & "2013 or older" & @CRLF
					$message &= @TAB & "Version List len" & @TAB & @TAB & "80" & @CRLF
				EndIf
			EndIf
			$message &= "Change Version List: " & @TAB & $AddTextToVL & @CRLF
			If $AddTextToVL = True Then
				$message &= @TAB & "Text to add: " & @TAB & @TAB & $TextToAddToVL & @CRLF
			EndIf
			$message &= "Change Modified field: " & @TAB & $SetModified & @CRLF
			If $SetModified = True Then
				If $SetModifiedYes = True Then
					$message &= @TAB & "Set Modified to:" & @TAB & @TAB & "Yes" & @CRLF
				Else
					$message &= @TAB & "Set Modified to:" & @TAB & @TAB & "No" & @CRLF
				EndIf
			EndIf
			$message &= "Select Marked Only: " & @TAB & $SetMarkedOnly & @CRLF
			$message &= "Open Export Objects: " & @TAB & $SetOpenExport & @CRLF & @CRLF
			$message &= "To cancel marking process and exit the " & $title & ", press PAUSE"
			$ret = MsgBox(1 + 32 + 8192, $title, $message)
			If $ret = 1 Then
				_Proceed()
			Else
				WinSetTitle($NAVtitle,"Status Bar",$NAVOldtitle)
				$NAVOldtitle = ""
				$NAVtitle = ""
				GUICtrlSetState($Checkbox_window, $GUI_UNCHECKED)
				GUICtrlSetState($Button_start, $GUI_DISABLE)
			EndIf
	EndSwitch
WEnd

Func _SaveIni()
	IniWrite($inifile,"Settings","ListFile",$ListFile)
	IniWrite($inifile,"Settings","SetModified",$SetModified)
	IniWrite($inifile,"Settings","SetModifiedYes",$SetModifiedYes)
	IniWrite($inifile,"Settings","Delay",$Delay)
	IniWrite($inifile,"Settings","AddTextToVL",$AddTextToVL)
	IniWrite($inifile,"Settings","TextToAddToVL",$TextToAddToVL)
	IniWrite($inifile,"Settings","ClearFilter",$ClearFilter)
	IniWrite($inifile,"Settings","SetMarkedOnly",$SetMarkedOnly)
	IniWrite($inifile,"Settings","SetOpenExport",$SetOpenExport)
	IniWrite($inifile,"Settings","DelayMin",$DelayMin)
	IniWrite($inifile,"Settings","DelayMax",$DelayMax)
	IniWrite($inifile,"Settings","NAVVersion2013R2Plus",$NAVVersion2013R2Plus)
EndFunc

Func _ObjectsLoad()
	$linescount = 0
	If Not FileExists($ListFile) Then
		Return
	EndIf
	$file = FileOpen($ListFile, 0)
	;~ считаем количество линий и забиваем массив
	While 1
		$listlines[$linescount] = FileReadLine($file)
		If @error = -1 Then
			ExitLoop
		Else
			$s = StringUpper(StringMid($listlines[$linescount],1,3))
			For $k = 0 to 8
				If $s = $objects[$k][0] Then
					$linescount += 1
					ExitLoop
				EndIf
			Next
		EndIf
	Wend
	If $linescount = 0 Then
		Return
	EndIf
	;~ приводим к нормальному виду
	For $i = 0 To $linescount - 1
		$listlines[$i] = StringReplace($listlines[$i], " ", "", 0, 0)
		$listlines[$i] = StringReplace($listlines[$i], "_", "", 0, 0)
		$listlines[$i] = StringReplace($listlines[$i], ".txt", "", 0, 0)
		$listlines[$i] = StringReplace($listlines[$i], ".", "", 0, 0)
		$listlines[$i] = StringReplace($listlines[$i], "Codeunit", "COD", 0, 0)
		$listlines[$i] = StringReplace($listlines[$i], "DataPort", "DAT", 0, 0)
		$listlines[$i] = StringReplace($listlines[$i], "Form", "FOR", 0, 0)
		$listlines[$i] = StringReplace($listlines[$i], "MenuSuite", "MEN", 0, 0)
		$listlines[$i] = StringReplace($listlines[$i], "Report", "REP", 0, 0)
		$listlines[$i] = StringReplace($listlines[$i], "Query", "QUE", 0, 0)
		$listlines[$i] = StringReplace($listlines[$i], "Table", "TAB", 0, 0)
		$listlines[$i] = StringReplace($listlines[$i], "XMLPort", "XML", 0, 0)
		$listlines[$i] = StringReplace($listlines[$i], "Page", "PAG", 0, 0)
	Next
	;~ сортируем по объектам для удобства
	_ArraySort($listlines,0,0,$linescount - 1)
	;~ разбиваем на объекты...
	$objtypes = 0
	$i = 0
	$j = 0
	$l = 0
	$o = StringUpper(StringMid($listlines[$i],1,3))
	$objlist[$objtypes][0] = $o
	$objlist[$objtypes][1] = 0
	$objlist[$objtypes][2] = ""
	While $i < $linescount
		While $o = $objlist[$objtypes][0]
			$obj[$objtypes][$j] = Number(StringMid($listlines[$i],4))
			$i += 1
			$j += 1
			$o = StringUpper(StringMid($listlines[$i],1,3))
		WEnd
		$objlist[$objtypes][1] = $j - 1
		$j = 0
		$objtypes += 1
		If $objtypes < 9 Then
			$objlist[$objtypes][0] = $o
		Endif
	WEnd
	$objtypes = $objtypes - 1
	;сортируем методом вставки по номерам
	For $k = 0 to $objtypes
		For $l = 0 To 8
			If $objlist[$k][0] = $objects[$l][0] Then
				$objlist[$k][2] = $objects[$l][1]
				ExitLoop
			EndIf
		Next
		$x = 0
		For $i = 0 to $objlist[$k][1] ; цикл проходов, i - номер прохода
			$x = $obj[$k][$i]
			;поиск места элемента в готовой последовательности
			$j = $i - 1
			While $j >= 0 And $obj[$k][$j] > $x
				$obj[$k][$j + 1] = $obj[$k][$j] ;сдвигаем элемент направо, пока не дошли
				$j -= 1
			WEnd
			;место найдено, вставить элемент
			$obj[$k][$j + 1] = $x
		Next
	Next
EndFunc

Func _LoadFile()
	GUICtrlSetData($Input_file, $ListFile)
	$timer_init = TimerInit()
	_ObjectsLoad()
	$timer_diff = TimerDiff($timer_init)
	$message  = "File was successfully processed" & @CRLF
	$message &= "Elapsed time: " & Int($timer_diff) / 1000 & " seconds" & @CRLF
	$message &= "The total number of objects to process: " & $linescount
	TrayTip($title, $message, 10, 1)
	GUICtrlSetData($Label_objects, "The total number of objects to process: " & $linescount)
	If $linescount = 0 Then
		Return
	EndIf
	GUICtrlSetState($Checkbox_objects,$GUI_CHECKED)
	_ReadyCheck()
EndFunc

Func _MarkIt($t)
	Local $c = -1
	While $c <> $t
		Sleep($Delay)
		Send("^{INS}")
		$c = Number(ClipGet())
		If $c = $t Then
			Send("^{F1}")
		EndIf
		If $c > $t Then
			Return
		EndIf
		Send("{DOWN}")
	WEnd
EndFunc

Func _zzz()
	If $NAVOldtitle <> "" Then
		WinSetState($NAVtitle,"Status Bar",@SW_ENABLE)
		WinSetTitle($NAVtitle,"Status Bar",$NAVOldtitle)
	EndIf
	$message = "The work of the tool has been cancelled by the user"
	MsgBox(16, $title, $message)
	_SaveIni()
	Exit(0)
EndFunc

Func _ReadyCheck()
	If  GUICtrlRead($Checkbox_objects) = $GUI_CHECKED _
	And GUICtrlRead($Checkbox_window) = $GUI_CHECKED _
	And $listlines <> 0 _
	Then
		GUICtrlSetState($Button_start, $GUI_ENABLE)
	Else
		GUICtrlSetState($Button_start, $GUI_DISABLE)
	EndIf
EndFunc

Func _Proceed()
	GUISetState(@SW_HIDE)
	WinActivate($NAVtitle)
	$timer_init = TimerInit()
	Send("+{F12}")
	If $ClearFilter = True Then
	;filter objects (if filters do not need to be saved)
		Send('!va')
;~ 		Send("^+{F7}")
		For $i = 0 To $objtypes
			Send("!" & $objlist[$i][2] & "^{HOME}{HOME}{RIGHT}")
			Sleep($Delay)
			$filterstring = ""
			$k = 0
			For $j = 0 To $objlist[$i][1] + 1
				If $k > 120 Or $j >= $objlist[$i][1] + 1 Then
					$filterstring = StringTrimRight($filterstring, 1)
					If $filterstring = "" Then
						ContinueLoop
					EndIf
					ClipPut($filterstring)
					Send("{F7}")
					Send("+{INS}{ENTER}")
					Send("^a")
					Send("^{F1}")
					Send('!va')
;~ 					Send("^+{F7}")
					Send("{DOWN}{UP}")
					$filterstring = ""
					$j -= 1
					$k = 0
				Else
					$filterstring &= $obj[$i][$j] & "|"
					$k += 1
				EndIf
			Next
		Next
	Else
	;check every object (if need to leave filters)
		For $i = 0 To $objtypes
			Send("!" & $objlist[$i][2] & "^{HOME}{HOME}{RIGHT}")
			Sleep($Delay)
			For $j = 0 To $objlist[$i][1]
				_MarkIt($obj[$i][$j])
			Next
		Next
	EndIf
	Send("!a")
	If ($AddTextToVL = True And $TextToAddToVL <> "") Or $SetModified = True Then
		Send("!vm")
	EndIf
	If $AddTextToVL = True And $TextToAddToVL <> "" Then
		Local $maxlen
		If $NAVVersion2013R2Plus = True Then
			$maxlen = 248
		Else
			$maxlen = 80
		EndIf
		Send("^{HOME}{HOME}{RIGHT 4}")
		While 1
			$state = WinGetState($NAVtitle)
			If Not BitAnd($state, 8) Then
				ExitLoop
			EndIf
			ClipPut("")
			Send("^{INS}")
			Sleep($Delay)
			$s = ClipGet()
			If $s = "" Then
				ClipPut($TextToAddToVL)
				Send("{F2}")
				Sleep($Delay)
				Send("+{INS}")
				Send("{DOWN}")
				ContinueLoop
			EndIf
			If StringInStr($s, $TextToAddToVL, 1, -1) = 0 Then
				$str = $s & "," & $TextToAddToVL
				WinSetState($NAVtitle,"Status Bar",@SW_DISABLE)
				While StringLen($str) > $maxlen
					$message  = "The length of the Version List field must not exceed " & $maxlen & " characters (now its length is " & StringLen($str) & " characters)." & @CRLF
					$message &= "In the text box below, edit the version list to meet the rule." & @CRLF & @CRLF
					$message &= "If you click Cancel, the version list for this object will NOT be changed."
					$str = InputBox($title, $message, $str, "", 550, 160)
					If $str = "" Then
						If @error = 1 Then
							$str = $s
							ExitLoop
						EndIf
					EndIf
				WEnd
				WinSetState($NAVtitle,"Status Bar",@SW_ENABLE)
				WinActivate($NAVtitle)
				Sleep($Delay)
				If $str <> $s Then
					ClipPut($str)
					Send("{F2}")
					Sleep($Delay)
					Send("+{INS}")
				EndIf
			EndIf
			Send("{DOWN}")
		WEnd
		Send("{ENTER}{DEL}{UP}")
	EndIf
	If $SetModified = True Then
		Send("^{HOME}{HOME}{RIGHT 3}{F7}")
		If $SetModifiedYes = True Then
			ClipPut("No")
		Else
			ClipPut("Yes")
		EndIf
		Send("+{INS}{ENTER}")
		While 1
			$state = WinGetState($NAVtitle)
			If Not BitAnd($state, 8) Then
				ExitLoop
			EndIf
			Send("{F2} {DOWN}")
		WEnd
		Send("{ENTER} {F7}{DEL}{ENTER}^{HOME}")
	EndIf
	$timer_diff = TimerDiff($timer_init)
	$sec = Int($timer_diff) / 1000
	$min = 0
	if $sec > 60 Then
		$min = int($sec / 60)
		$sec = $sec - $min * 60
	endif
	$message  = "Object marking completed" & @CRLF
	$message &= "Elapsed Time: " & $min & " minutes " & $sec & " seconds"
	TrayTip($title, $message, 10, 1)
	If ($AddTextToVL = True And $TextToAddToVL <> "") Or $SetModified = True Then
		If $SetMarkedOnly = False Then
			Send("!vm^a")
		ElseIf $SetOpenExport = True Then
			Send("!ft")
		EndIf
	ElseIf $SetMarkedOnly = True Then
			Send("!vm^a")
		If $SetOpenExport = True Then
			Send("!ft")
		EndIf
	EndIf
	WinSetTitle($NAVtitle,"Status Bar",$NAVOldtitle)
	WinWaitActive($NAVOldtitle)
	$NAVOldtitle = ""
	$NAVtitle = ""
	GUICtrlSetState($Checkbox_window, $GUI_UNCHECKED)
	GUICtrlSetState($Button_start, $GUI_DISABLE)
	GUISetState(@SW_SHOW)
EndFunc
;the end