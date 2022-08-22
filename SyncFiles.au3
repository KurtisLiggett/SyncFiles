#AutoIt3Wrapper_Run_Au3Stripper=y

#Region doc-info
;------------------------------------------------------------------------------
;	Copyright 2022 Kurtis Liggett
;
;	This program is free software: you can redistribute it and/or modify it
;	under the terms of the GNU General Public License as published by the Free
;	Software Foundation, either version 3 of the License, or (at your option)
;	any later version.
;
;	This program is distributed in the hope that it will be useful, but WITHOUT
;	ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
;	FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
;	more details.
;
;	You should have received a copy of the GNU General Public License along
;	with this program. If not, see <https://www.gnu.org/licenses/>.
;------------------------------------------------------------------------------

;==============================================================================
;
; Name........... SyncFiles
;
; Description.... SyncFiles is an easy to use file and folder synchronization tool.
;
; Author......... kurtykurtyboy
;
; Credits........ Many thanks to examples and UDFs in the forums
;				  - LarsJ: example of virtual listviews and item colors
;		 		    - https://www.autoitscript.com/forum/topic/168707-listview-item-subitem-background-colour/#comment-1234009
;				  - Yashied: Copy UDF
;				    - https://www.autoitscript.com/forum/topic/121833-copy-udf/
;				  - Ward: MemoryDll UDF (and BinaryCall UDF)
;				    - https://www.autoitscript.com/forum/topic/77463-embed-dlls-in-script-and-call-functions-from-memory-memorydll-udf/
;				  - Ward: Json UDF
;				    - https://www.autoitscript.com/forum/topic/148114-a-non-strict-json-udf-jsmn/
;				  - Many other random snippets from the forums
;
;==============================================================================
#EndRegion doc-info

#Region includes-and-options
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <GuiListView.au3>
#include <GuiStatusBar.au3>
#include <WinAPISys.au3>
#include <WinAPIGdi.au3>
#include <GDIPlus.au3>
#include <Date.au3>
#include <Array.au3>
#include "include\_GetFileListRec.au3"
#include "include\MemoryDll.au3"
#include "include\Json.au3"
#include "include\Copy\Copy_binary.au3"
#include "include\GuiFlatButton.au3"

#NoTrayIcon
Opt("GUIOnEventMode", 1)
Opt("GUIResizeMode", $GUI_DOCKALL)
;~ Opt("TrayIconHide", 0)
;~ Opt("TrayAutoPause", 0)
;~ Opt("TrayOnEventMode", 1)
;~ Opt("TrayMenuMode", 3)
#EndRegion includes-and-options

#Region global-vars
;copy DLL does not work with X64!
#AutoIt3Wrapper_UseX64=N
#AutoIt3Wrapper_Icon=icon.ico
#AutoIt3Wrapper_OutFile=..\SyncFiles 0.9.0.exe
#AutoIt3Wrapper_Res_Fileversion=0.9.0

;App info
Global $guiName = "SyncFiles"
Global $version = "0.9.0"
Global $date = "8/19/2022"

;GUI/control widths and heights used for layout and sizing
Global $guiWidth = 900
Global $guiHeight = 605
Global $guiPosX = -1
Global $guiPosY = -1
Global $widthMid = 100
Global $spacingMidH = 0
Global $spacingSideH = 5
Global $browseButtonW = 21
Global $browseButtonH = 21
Global $lvSizeColWidth = 85
Global $profilesWidth = 200
Global $lvY = 33
Global $lvBottomOffset = 78
Global $lvHeaderHeight, $menuH, $statusbarHeight = 42
Global $lvScrollBarWidth = 25
Global $toolbarHeight = 25

;GUI controls
Global $hGUI, $hGuiAbout
Global $lv_profiles, $h_lv_profiles
Global $input_Left, $button_LeftSelect
Global $input_Right, $button_RightSelect
Global $lv_results, $h_lv_results, $h_lv_results_header
Global $button_analyze, $button_sync
Global $radio_mirrorLeft, $radio_mirrorRight, $radio_sync
Global $label_fileCount
Global $hStatusbar, $progressBar, $lvMenu, $h_lvMenu
Global $button_new, $button_save, $button_delete

;other
Global $aAccelKeys[1][2]
Global $tText = DllStructCreate("wchar[4096]")
Global $hotTracking
Global $objOptions
Global $sOptionsFile = @ScriptDir & "\userdata.json"
Global $aLastDirs[2] = ["", ""]

;listview data storage
Global $aLeftFileListFinal[0][3]
Global $aRightFileListFinal[0][3]
Global $aActionFinal[0], $aActionOverrides[0]
Global $aLvDataArray[0], $aLvColorArray[0]

;copy control
Global $syncNow, $copyAbort

;action strings
Global Enum $action_NoChange, $action_CopyRight, $action_CopyLeft, $action_UpdateRight, $action_UpdateLeft, $action_Unknown, $action_Error, $action_DeleteLeft, $action_DeleteRight, $action_NotEqual
Global $actionString_NoChange = "=", $actionString_CopyRight = "[+] ->", $actionString_CopyLeft = "<- [+]", $actionString_UpdateRight = "->", $actionString_UpdateLeft = "<-"
Global $actionString_Unknown = "?", $actionString_DeleteLeft = "<- [-]", $actionString_DeleteRight = "[-] ->", $actionString_NotEqual = "? ≠ ?"
#EndRegion global-vars


;call the main function
_main()


;------------------------------------------------------------------------------
; Title...........:	_main
; Description.....:	create GUI then run the main program loop
;------------------------------------------------------------------------------
Func _main()
	_GDIPlus_Startup() ;initialize GDI+

	;load options from file
	Local $sOptionsFileData = FileRead($sOptionsFile)
	If @error Then
		$sOptionsFileData = ""
	EndIf
	$objOptions = Json_Decode($sOptionsFileData)
	$guiPosX = _Json_Get($objOptions, ".Options.WindowPosX", $guiPosX)
	$guiPosY = _Json_Get($objOptions, ".Options.WindowPosY", $guiPosY)
	$guiWidth = _Json_Get($objOptions, ".Options.WindowWidth", $guiWidth)
	$guiHeight = _Json_Get($objOptions, ".Options.WindowHeight", $guiHeight)
	$aLastDirs[0] = _Json_Get($objOptions, ".LastDir.DirLeft", $aLastDirs[0])
	$aLastDirs[1] = _Json_Get($objOptions, ".LastDir.DirRight", $aLastDirs[1])

	;create the main GUI
	_guiCreate()

	;load profiles from file
	_LvProfiles_update()

	;show GUI and start program
	GUISetState(@SW_SHOWNORMAL)
	GUIRegisterMsg($WM_MOVE, "_WM_MOVE")
	GUIRegisterMsg($WM_SIZING, "_WM_SIZING")
	GUIRegisterMsg($WM_SIZE, "_WM_SIZE")
	GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

	Local $aCursorInfo, $aHitTest
	Local $headerBottomPos = $lvY + $lvHeaderHeight

	While 1
		;monitor to turn tooltip off
		If $hotTracking Then
			$aCursorInfo = GUIGetCursorInfo($hGUI)
			If $aCursorInfo[4] = $lv_results Then
				Local $aHitTest = _GUICtrlListView_SubItemHitTest($h_lv_results)
				If $aHitTest[0] = -1 Or $aHitTest[1] <> 2 Then
					$hotTracking = False
					ToolTip('')
				EndIf

				;check if over the header
				If $aHitTest[0] = 0 Then
					If $aCursorInfo[1] < $headerBottomPos Then
						$hotTracking = False
						ToolTip('')
					EndIf
				EndIf
			Else
				$hotTracking = False
				ToolTip('')
			EndIf
		EndIf

		;update the status bar
		_setStatusMessage()

		;execute copy and progress updates
		_FileCopy_Process()

		Sleep(10)
	WEnd
EndFunc   ;==>_main

#Region Synchronize-and-Copy
;------------------------------------------------------------------------------
; Title...........:	_onSynchronize
; Description.....:	open the copy dll and start the synchronization process
; Event...........: Synchronize button
;------------------------------------------------------------------------------
Func _onSynchronize()
	If Not $syncNow Then
		If Not _Copy_OpenDll() Then
			If @error <> 6 Then
				ConsoleWrite("error " & @error & @CRLF)
				_setStatusMessage("Error " & @error & ": Copy DLL not found!")
				Return SetError(1)
			EndIf
		EndIf

		$syncNow = True
		GUICtrlSetData($button_sync, "Abort Sync")
		;continue the copy logic in loop to track progress
		;without locking up the program
	Else
		$copyAbort = True
	EndIf
EndFunc   ;==>_onSynchronize

;------------------------------------------------------------------------------
; Title...........:	_FileCopy_Process
; Description.....:	called from main while loop to copy files and update progress bar
;------------------------------------------------------------------------------
Func _FileCopy_Process()
	Static Local $timer2
	Static Local $sDirLeft, $sDirRight
	Static Local $copyNone, $copyErrorFlag
	Static Local $_FileCopy_hProgressProc
	Local $copyResult, $copyState, $copyNext
	Static Local $copyIndex, $copyStarted, $copyTotalCount, $progressValuePrev
	Static Local $copyTotalSize, $copyCompletedSize, $partialCopySize, $progressValue
	Static Local $sFileCopyName, $bProcessed

	If $syncNow Then
		If Not $copyStarted Then
			$timer2 = TimerInit()
			$sDirLeft = GUICtrlRead($input_Left)
			$sDirRight = GUICtrlRead($input_Right)
			_setStatusMessage("Synchronizing folders...")

			;calculate total size to be copied
			$copyTotalSize = 0
			$copyCompletedSize = 0
			$partialCopySize = 0
			$progressValue = 0
			$copyErrorFlag = 0
			$copyIndex = 0
			$copyNone = True
			$bProcessed = False
			For $i = 0 To UBound($aLeftFileListFinal) - 1
				Switch $aActionOverrides[$i]
					Case $action_CopyLeft, $action_UpdateLeft, $action_DeleteRight
						$copyTotalSize += Number($aRightFileListFinal[$i][2])
						$copyNone = False
					Case $action_CopyRight, $action_UpdateRight, $action_DeleteLeft
						$copyTotalSize += Number($aLeftFileListFinal[$i][2])
						$copyNone = False
				EndSwitch
			Next
			$copyTotalCount = UBound($aLeftFileListFinal)
			$copyStarted = True
		EndIf

		;perform the copy
		If $copyNone Then
			_setStatusMessage("Synchronizing completed successfully: No change")
			$syncNow = False
		Else
			If $copyIndex < $copyTotalCount And Not $copyAbort Then
				If $bProcessed Then
					$copyNext = True
				Else
					$copyState = _Copy_GetState(0)
					If @error Then
						$copyNext = True
						$bProcessed = True
					Else
						If $copyState[0] Then
;~ 							ConsoleWrite($copyState[1] & " - " & $copyState[2] & @CRLF)
							$partialCopySize = $copyState[1]
						Else
							Switch $copyState[5]
								Case 0
									$copyCompletedSize += $copyState[2]
								Case 1235 ; ERROR_REQUEST_ABORTED
									$copyErrorFlag = True
								Case Else
									$copyErrorFlag = True
							EndSwitch
							$partialCopySize = 0
							$copyNext = True
							$bProcessed = True
						EndIf
					EndIf
				EndIf


				If $copyNext Then
					$bProcessed = False
					Switch $aActionOverrides[$copyIndex]
						Case $action_CopyRight
							$sFileCopyName = "Copy file: " & StringTrimLeft($aLeftFileListFinal[$copyIndex][0], 1)
							If $aLeftFileListFinal[$copyIndex][1] = 1 Then
								DirCreate($sDirRight & $aLeftFileListFinal[$copyIndex][0])
								$bProcessed = True
							Else
								_Copy_CopyFile($sDirLeft & $aLeftFileListFinal[$copyIndex][0], $sDirRight & $aLeftFileListFinal[$copyIndex][0], 0, 0)
							EndIf

						Case $action_CopyLeft
							$sFileCopyName = "Copy file: " & StringTrimLeft($aRightFileListFinal[$copyIndex][0], 1)
							If $aRightFileListFinal[$copyIndex][1] = 1 Then
								DirCreate($sDirLeft & $aRightFileListFinal[$copyIndex][0])
								$bProcessed = True
							Else
								_Copy_CopyFile($sDirRight & $aRightFileListFinal[$copyIndex][0], $sDirLeft & $aRightFileListFinal[$copyIndex][0], 0, 0)
							EndIf

						Case $action_NoChange
							$bProcessed = True

						Case $action_UpdateRight
							$sFileCopyName = "Update file: " & StringTrimLeft($aRightFileListFinal[$copyIndex][0], 1)
							_Copy_CopyFile($sDirLeft & $aLeftFileListFinal[$copyIndex][0], $sDirRight & $aLeftFileListFinal[$copyIndex][0], 0, 0)

						Case $action_UpdateLeft
							$sFileCopyName = "Update file: " & StringTrimLeft($aLeftFileListFinal[$copyIndex][0], 1)
							_Copy_CopyFile($sDirRight & $aRightFileListFinal[$copyIndex][0], $sDirLeft & $aRightFileListFinal[$copyIndex][0], 0, 0)

						Case $action_DeleteLeft
							$sFileCopyName = "Delete file: " & StringTrimLeft($aLeftFileListFinal[$copyIndex][0], 1)
							If $aLeftFileListFinal[$copyIndex][1] = 1 Then
								DirRemove($sDirLeft & $aLeftFileListFinal[$copyIndex][0], 1)
								$bProcessed = True
							Else
								FileDelete($sDirLeft & $aLeftFileListFinal[$copyIndex][0])
								$copyCompletedSize += $aLeftFileListFinal[$copyIndex][2]
								$bProcessed = True
							EndIf

						Case $action_DeleteRight
							$sFileCopyName = "Delete file: " & StringTrimLeft($aRightFileListFinal[$copyIndex][0], 1)
							If $aRightFileListFinal[$copyIndex][1] = 1 Then
								DirRemove($sDirRight & $aRightFileListFinal[$copyIndex][0], 1)
								$bProcessed = True
							Else
								FileDelete($sDirRight & $aRightFileListFinal[$copyIndex][0])
								$copyCompletedSize += $aRightFileListFinal[$copyIndex][2]
								$bProcessed = True
							EndIf

						Case Else    ;unknown -- no action
							$bProcessed = True

					EndSwitch
					$copyIndex += 1
				EndIf

				$progressValue = Round($guiWidth * (($copyCompletedSize + $partialCopySize) / $copyTotalSize))
				If $progressValue <> $progressValuePrev Then
					_setStatusMessage("Synchronizing folders...(" & Round($progressValue / $guiWidth * 100) & "%)" & @TAB & $sFileCopyName, True)
					GUICtrlSetPos($progressBar, Default, Default, $progressValue, Default)
					$progressValuePrev = $progressValue
				EndIf

				;finished copying
			Else
				;check it aborted first
				If $copyAbort Then
					_Copy_Abort(0)
					$copyAbort = False
				EndIf

				GUICtrlSetData($button_sync, "Synchronize")

				;re-analyze the folders
				_setStatusMessage("Re-analyizing")
				_onAnalyze()
				If Not $copyErrorFlag Then
					_setStatusMessage("Synchronizing completed successfully")
				Else
					_setStatusMessage("Synchronizing completed with errors")
				EndIf
				GUICtrlSetPos($progressBar, Default, Default, 0, Default)
				$syncNow = False
			EndIf
		EndIf

		If Not $syncNow Then
			$copyStarted = False
			GUICtrlSetData($button_sync, "Synchronize")
			ConsoleWrite("Done: " & TimerDiff($timer2) & @CRLF)
			_Copy_CloseDll()
		EndIf

	EndIf
EndFunc   ;==>_FileCopy_Process
#EndRegion Synchronize-and-Copy


#Region GUI-creation-and-position
;------------------------------------------------------------------------------
; Title...........:	_guiCreate
; Description.....:	Create the main GUI
;------------------------------------------------------------------------------
Func _guiCreate()
	Local $colRightX = ($guiWidth - $profilesWidth) / 2 + $widthMid / 2 + $spacingMidH - $lvScrollBarWidth / 2 + $profilesWidth
	Local $midX = ($guiWidth - $profilesWidth) / 2 - $widthMid / 2 + $profilesWidth
	Local $lvW = ($guiWidth - $profilesWidth) / 2 - $widthMid / 2 - $spacingSideH - $spacingMidH
	Local $lvH = $guiHeight - $lvY - $lvBottomOffset - $statusbarHeight
	Local $colX = $profilesWidth + $spacingSideH

	$hGUI = GUICreate($guiName, $guiWidth, $guiHeight, $guiPosX, $guiPosY, BitOR($GUI_SS_DEFAULT_GUI, $WS_SIZEBOX, $WS_MAXIMIZEBOX), $WS_EX_ACCEPTFILES)
	GUISetOnEvent($GUI_EVENT_CLOSE, "_onExitMain")
	GUISetOnEvent($GUI_EVENT_DROPPED, "_onDropped")
	GUISetOnEvent($GUI_EVENT_MAXIMIZE, "_onMaximized")
	GUISetOnEvent($GUI_EVENT_RESTORE, "_onRestore")

	Local $menu_file = GUICtrlCreateMenu("File")
	GUICtrlCreateMenuItem("New", $menu_file)
	GUICtrlSetOnEvent(-1, "_onProfileNewItem")
	GUICtrlCreateMenuItem("Save", $menu_file)
	GUICtrlSetOnEvent(-1, "_onProfileSave")
	GUICtrlCreateMenuItem("Save As...", $menu_file)
	GUICtrlSetOnEvent(-1, "_onProfileSaveAs")
	GUICtrlCreateMenuItem("Rename", $menu_file)
	GUICtrlSetOnEvent(-1, "_onProfileRename")
	GUICtrlCreateMenuItem("Delete", $menu_file)
	GUICtrlSetOnEvent(-1, "_onProfileDelete")
	GUICtrlCreateMenuItem("", $menu_file)
	GUICtrlCreateMenuItem("Exit", $menu_file)
	GUICtrlSetOnEvent(-1, "_onMenuExit")
	Local $menu_tools = GUICtrlCreateMenu("Tools")
	GUICtrlCreateMenuItem("Analyze", $menu_tools)
	GUICtrlSetOnEvent(-1, "_onAnalyze")
	GUICtrlCreateMenuItem("Synchronize", $menu_tools)
	GUICtrlSetOnEvent(-1, "_onSynchronize")
	Local $menu_help = GUICtrlCreateMenu("Help")
	GUICtrlCreateMenuItem("About", $menu_help)
	GUICtrlSetOnEvent(-1, "_onMenuAbout")

	$menuH = _WinAPI_GetSystemMetrics($SM_CYMENU)

	Local $top_border = GUICtrlCreateLabel("", 0, 0, $guiWidth, 1)
	GUICtrlSetBkColor(-1, 0xBBBBBB)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT)
	GUICtrlSetState(-1, $GUI_DISABLE)

	;profiles
	$lv_profiles = GUICtrlCreateListView("Profiles", 0, $toolbarHeight, $profilesWidth, $guiHeight - $statusbarHeight - 5 - $toolbarHeight, BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $LVS_SINGLESEL))
	GUICtrlSetResizing($lv_profiles, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH)
	$h_lv_profiles = GUICtrlGetHandle($lv_profiles)
	_GUICtrlListView_SetExtendedListViewStyle($lv_profiles, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_DOUBLEBUFFER))
	_GUICtrlListView_SetColumnWidth($lv_profiles, 0, $profilesWidth - $lvScrollBarWidth)

	;profile buttons
	Local $tbButtonOffset = 2
	Local $menuColor = _WinAPI_GetSysColor($COLOR_MENUBAR)
	Local $hoverColor = $menuColor * 0.9
	Local $buttonSpace = 0
	Local $aColorsEx = _
			[$menuColor, 0xFCFCFC, $menuColor, _     ; normal 		: Background, Text, Border
			$menuColor, 0xFCFCFC, $menuColor, _      ; focus 		: Background, Text, Border
			$hoverColor, 0xFCFCFC, $hoverColor, _    ; hover 		: Background, Text, Border
			$menuColor, 0xFCFCFC, $menuColor]        ; selected 	: Background, Text, Border

	;background
	GUICtrlCreateLabel("", 0, 0 - 1, $profilesWidth - 1, 22 + 2)
	GUICtrlSetBkColor(-1, $menuColor)
	GUICtrlSetState(-1, $GUI_DISABLE)

	$button_new = GuiFlatButton_Create("", 2, $tbButtonOffset, 22, 22, $BS_TOOLBUTTON)
	GUICtrlSetTip(-1, "Create new profile")
	GUICtrlSetOnEvent(-1, "_onProfileNewItem")
	GuiFlatButton_SetColorsEx(-1, $aColorsEx)
	_WinAPI_DeleteObject(_SendMessage(GUICtrlGetHandle($button_new), $BM_SETIMAGE, $IMAGE_ICON, GetIconData(0)))

	$button_save = GuiFlatButton_Create("", 2 + 22 + $buttonSpace, $tbButtonOffset, 22, 22, $BS_TOOLBUTTON)
	GUICtrlSetTip(-1, "Save profile")
	GUICtrlSetOnEvent(-1, "_onProfileSave")
	GuiFlatButton_SetColorsEx(-1, $aColorsEx)
	_WinAPI_DeleteObject(_SendMessage(GUICtrlGetHandle($button_save), $BM_SETIMAGE, $IMAGE_ICON, GetIconData(1)))

	;divider
	GUICtrlCreateLabel("", 2 + 22 * 2 + 3, 0, 1, 26)
	GUICtrlSetBkColor(-1, 0xBBBBBB)
	GUICtrlSetState(-1, $GUI_DISABLE)

	$button_delete = GuiFlatButton_Create("", 2 + 22 * 2 + 6, $tbButtonOffset, 22, 22, $BS_TOOLBUTTON)
	GUICtrlSetTip(-1, "Delete profile")
	GUICtrlSetOnEvent(-1, "_onProfileDelete")
	GuiFlatButton_SetColorsEx(-1, $aColorsEx)
	_WinAPI_DeleteObject(_SendMessage(GUICtrlGetHandle($button_delete), $BM_SETIMAGE, $IMAGE_ICON, GetIconData(2)))

	;right line
	GUICtrlCreateLabel("", $profilesWidth - 1, 0 - 1, 1, 26)
	GUICtrlSetBkColor(-1, 0x888888)
	GUICtrlSetState(-1, $GUI_DISABLE)

	;bottom line
;~ 	GUICtrlCreateLabel("", 0, 22 + 2, $profilesWidth - 1, 1)
;~ 	GUICtrlSetBkColor(-1, 0xAAAAAA)
;~ 	GUICtrlSetState(-1, $GUI_DISABLE)

	;other
	$input_Left = GUICtrlCreateInput($aLastDirs[0], $colX, 5, $lvW - $browseButtonW - 2 + $lvScrollBarWidth / 2, 21)
	GUICtrlSetState(-1, $GUI_DROPACCEPTED)
	$button_LeftSelect = GUICtrlCreateButton("...", $colX + $lvW - $browseButtonW + $lvScrollBarWidth / 2, 5, $browseButtonW, $browseButtonH)
	GUICtrlSetOnEvent(-1, "_onLeftSelect")

	$input_Right = GUICtrlCreateInput($aLastDirs[1], $colRightX, 5, $lvW - $browseButtonW - 2 + $lvScrollBarWidth / 2, 21)
	GUICtrlSetState(-1, $GUI_DROPACCEPTED)
	$button_RightSelect = GUICtrlCreateButton("...", $colRightX + $lvW - $browseButtonW + $lvScrollBarWidth / 2, 5, $browseButtonW, $browseButtonH)
	GUICtrlSetOnEvent(-1, "_onRightSelect")

	$lv_results = GUICtrlCreateListView("File Name|Size|Action|File Name|Size", $colX, $lvY, $guiWidth - $spacingSideH * 2 - $profilesWidth, $lvH, $LVS_OWNERDATA)
	GUICtrlSetResizing($lv_results, $GUI_DOCKTOP + $GUI_DOCKBOTTOM + $GUI_DOCKLEFT + $GUI_DOCKRIGHT)
	$h_lv_results = GUICtrlGetHandle($lv_results)
	$h_lv_results_header = _GUICtrlListView_GetHeader($h_lv_results)
	$lvHeaderHeight = _WinAPI_GetWindowHeight(_GUICtrlListView_GetHeader($h_lv_results))
	_GUICtrlListView_SetExtendedListViewStyle($lv_results, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT, $LVS_EX_DOUBLEBUFFER))
	_GUICtrlListView_SetColumnWidth($lv_results, 0, $lvW - $lvSizeColWidth - $lvScrollBarWidth / 2)
	_GUICtrlListView_SetColumnWidth($lv_results, 1, $lvSizeColWidth)
	_GUICtrlListView_SetColumnWidth($lv_results, 2, $widthMid - 4)
	_GUICtrlListView_SetColumnWidth($lv_results, 3, $lvW - $lvSizeColWidth - $lvScrollBarWidth / 2)
	_GUICtrlListView_SetColumnWidth($lv_results, 4, $lvSizeColWidth)
	_GUICtrlListView_JustifyColumn($lv_results, 1, 1)
	_GUICtrlListView_JustifyColumn($lv_results, 2, 2)
	_GUICtrlListView_JustifyColumn($lv_results, 4, 1)

	;create the listview context menu
	$lvMenu = GUICtrlCreateContextMenu($lv_results)
	$h_lvMenu = GUICtrlGetHandle($lvMenu)
	Local $lvMenu_CopyLeft = GUICtrlCreateMenuItem("<- Copy to left", $lvMenu)
	GUICtrlSetOnEvent(-1, "_onLvMenuCopyLeft")
	Local $lvMenu_CopyRight = GUICtrlCreateMenuItem("Copy to right ->", $lvMenu)
	GUICtrlSetOnEvent(-1, "_onLvMenuCopyRight")
	GUICtrlCreateMenuItem("", $lvMenu)
	Local $lvMenu_DeleteLeft = GUICtrlCreateMenuItem("<- Delete left", $lvMenu)
	GUICtrlSetOnEvent(-1, "_onLvMenuDeleteLeft")
	Local $lvMenu_DeleteRight = GUICtrlCreateMenuItem("Delete right ->", $lvMenu)
	GUICtrlSetOnEvent(-1, "_onLvMenuDeleteRight")
	GUICtrlCreateMenuItem("", $lvMenu)
	Local $lvMenu_NoChange = GUICtrlCreateMenuItem("No change", $lvMenu)
	GUICtrlSetOnEvent(-1, "_onLvMenuNoChange")

	Local $radioWidth = 115
	$radio_mirrorLeft = GUICtrlCreateRadio("<- Mirror to left", ($guiWidth - $profilesWidth) / 2 + $profilesWidth - $radioWidth / 2 - $radioWidth - 5, $lvY + $lvH + 3, $radioWidth)
	GUICtrlSetResizing(-1, $GUI_DOCKBOTTOM + $GUI_DOCKHCENTER + $GUI_DOCKHEIGHT + $GUI_DOCKWIDTH)
	GUICtrlSetOnEvent(-1, "_onRadioMirrorLeft")
	$radio_sync = GUICtrlCreateRadio("<- Synchronize ->", ($guiWidth - $profilesWidth) / 2 + $profilesWidth - $radioWidth / 2, $lvY + $lvH + 3, $radioWidth)
	GUICtrlSetResizing(-1, $GUI_DOCKBOTTOM + $GUI_DOCKHCENTER + $GUI_DOCKHEIGHT + $GUI_DOCKWIDTH)
	GUICtrlSetOnEvent(-1, "_onRadioSync")
	$radio_mirrorRight = GUICtrlCreateRadio("Mirror to right ->", ($guiWidth - $profilesWidth) / 2 + $profilesWidth + $radioWidth / 2 + 5, $lvY + $lvH + 3, $radioWidth)
	GUICtrlSetResizing(-1, $GUI_DOCKBOTTOM + $GUI_DOCKHCENTER + $GUI_DOCKHEIGHT + $GUI_DOCKWIDTH)
	GUICtrlSetOnEvent(-1, "_onRadioMirrorRight")
	GUICtrlSetState($radio_sync, $GUI_CHECKED)

	$button_analyze = GUICtrlCreateButton("Analyze", $guiWidth - 335, $guiHeight - 55 - $statusbarHeight + 5, 161, 41)
	GUICtrlSetOnEvent(-1, "_onAnalyze")
	GUICtrlSetResizing(-1, $GUI_DOCKBOTTOM + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT + $GUI_DOCKWIDTH)
	GUICtrlSetTip(-1, "Analyze folders")
	$button_sync = GUICtrlCreateButton("Synchronize", $guiWidth - 170, $guiHeight - 55 - $statusbarHeight + 5, 161, 41)
	GUICtrlSetResizing(-1, $GUI_DOCKBOTTOM + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT + $GUI_DOCKWIDTH)
	GUICtrlSetOnEvent(-1, "_onSynchronize")
	GUICtrlSetTip(-1, "Synchronize folders")

	$label_fileCount = GUICtrlCreateLabel("", $profilesWidth + 5, $guiHeight - $statusbarHeight - 20, 200, 25)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetResizing(-1, $GUI_DOCKBOTTOM + $GUI_DOCKLEFT + $GUI_DOCKHEIGHT + $GUI_DOCKWIDTH)

	$progressBar = GUICtrlCreateLabel("", 0, $guiHeight - $statusbarHeight - 6, 0, 5)
	GUICtrlSetBkColor(-1, 0x3CB043)
	GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)

	$hStatusbar = _GUICtrlStatusBar_Create($hGUI)
	GUICtrlSetResizing($hStatusbar, $GUI_DOCKBOTTOM + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT)

	;set up accelerators
	Local Const $accel_CtrlA = GUICtrlCreateDummy()
	$aAccelKeys[0][0] = '^a'
	$aAccelKeys[0][1] = $accel_CtrlA
	GUISetAccelerators($aAccelKeys)
	GUICtrlSetOnEvent($accel_CtrlA, "_onLvSelectAll")


	;tray menu
	TrayCreateItem("About")
	TrayItemSetOnEvent(-1, "_onMenuAbout")
	TrayCreateItem("")
	TrayCreateItem("Exit")
	TrayItemSetOnEvent(-1, "_onMenuExit")
	TraySetToolTip($guiName)
EndFunc   ;==>_guiCreate

;------------------------------------------------------------------------------
; Title...........:	_setControlSizes
; Description.....:	Resize controls to fit GUI size
;------------------------------------------------------------------------------
Func _setControlSizes()
	Local $aPos = WinGetClientSize($hGUI, "")
	$guiWidth = $aPos[0]
	$guiHeight = $aPos[1]

	Local $colRightX = ($guiWidth - $profilesWidth) / 2 + $widthMid / 2 + $spacingMidH - $lvScrollBarWidth / 2 + $profilesWidth
	Local $midX = ($guiWidth - $profilesWidth) / 2 - $widthMid / 2 + $profilesWidth
	Local $lvW = ($guiWidth - $profilesWidth) / 2 - $widthMid / 2 - $spacingSideH - $spacingMidH
	Local $lvH = $guiHeight - $lvY - $lvBottomOffset - $statusbarHeight
	Local $colX = $profilesWidth + $spacingSideH

	GUICtrlSetPos($input_Left, $colX, 5, $lvW - $browseButtonW - 2 + $lvScrollBarWidth / 2, 21)
	GUICtrlSetPos($button_LeftSelect, $colX + $lvW - $browseButtonW + $lvScrollBarWidth / 2, 5, $browseButtonW, $browseButtonH)

	GUICtrlSetPos($input_Right, $colRightX, 5, $lvW - $browseButtonW - 2 + $lvScrollBarWidth / 2, 21)
	GUICtrlSetPos($button_RightSelect, $colRightX + $lvW - $browseButtonW + $lvScrollBarWidth / 2, 5, $browseButtonW, $browseButtonH)

	Local $colSizeLeft = _GUICtrlListView_GetColumnWidth($h_lv_results, 1)
	Local $colSizeRight = _GUICtrlListView_GetColumnWidth($h_lv_results, 4)
	_GUICtrlListView_SetColumnWidth($lv_results, 0, $lvW - $colSizeLeft - $lvScrollBarWidth / 2)
	_GUICtrlListView_SetColumnWidth($lv_results, 1, $colSizeLeft)
	_GUICtrlListView_SetColumnWidth($lv_results, 2, $widthMid - 4)
	_GUICtrlListView_SetColumnWidth($lv_results, 3, $lvW - $colSizeRight - $lvScrollBarWidth / 2)
	_GUICtrlListView_SetColumnWidth($lv_results, 4, $colSizeRight)

	_GUICtrlStatusBar_Resize($hStatusbar)
EndFunc   ;==>_setControlSizes

Func WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
	Local $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
	Local $hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
	Local $iCode = DllStructGetData($tNMHDR, "Code")

	Switch $hWndFrom
		Case $h_lv_results
			Switch $iCode
				Case $LVN_HOTTRACK
					Local $aSubItemHitTest = _GUICtrlListView_SubItemHitTest($h_lv_results)
					If $aSubItemHitTest[0] <> -1 And $aSubItemHitTest[1] = 2 Then
						Local $action = $aActionOverrides[$aSubItemHitTest[0]]
						Switch $action
							Case $action_CopyRight
								$itemText = "Copy to right"

							Case $action_CopyLeft
								$itemText = "Copy to left"

							Case $action_NoChange
								$itemText = "No action"

							Case $action_UpdateRight
								$itemText = "Update right"

							Case $action_UpdateLeft
								$itemText = "Update left"

							Case $action_NotEqual
								$itemText = "Warning: No action" & @CRLF & "Dates match, but file sizes are different"

							Case Else
								$itemText = "unknown"
						EndSwitch

						ToolTip($itemText)
						$hotTracking = True
					Else
						ToolTip('')
					EndIf

					;virtual listview and color change from larsj
					;https://www.autoitscript.com/forum/topic/168707-listview-item-subitem-background-colour/#comment-1234009
				Case $LVN_GETDISPINFOW
					; Fill virtual listview
					Local $tNMLVDISPINFO = DllStructCreate($tagNMLVDISPINFO, $lParam)
					If BitAND(DllStructGetData($tNMLVDISPINFO, "Mask"), $LVIF_TEXT) Then
						Local $sItem = $aLvDataArray[DllStructGetData($tNMLVDISPINFO, "Item")][DllStructGetData($tNMLVDISPINFO, "SubItem")]
						DllStructSetData($tText, 1, $sItem)
						DllStructSetData($tNMLVDISPINFO, "TextMax", StringLen($sItem))
						DllStructSetData($tNMLVDISPINFO, "Text", DllStructGetPtr($tText))
					EndIf

				Case $NM_CUSTOMDRAW
					Local $tNMLVCUSTOMDRAW = DllStructCreate($tagNMLVCUSTOMDRAW, $lParam)
					Local $dwDrawStage = DllStructGetData($tNMLVCUSTOMDRAW, "dwDrawStage")

					Switch $dwDrawStage              ; Holds a value that specifies the drawing stage

						Case $CDDS_PREPAINT
							; Before the paint cycle begins
							Return $CDRF_NOTIFYITEMDRAW ; Notify the parent window of any item-related drawing operations

						Case $CDDS_ITEMPREPAINT
							; Before painting an item
							Return $CDRF_NOTIFYSUBITEMDRAW ; Notify the parent window of any subitem-related drawing operations

						Case BitOR($CDDS_ITEMPREPAINT, $CDDS_SUBITEM)
							; Before painting a subitem
							;Local $iItem = DllStructGetData($tNMLVCUSTOMDRAW, "dwItemSpec")                ; Item index
							;Local $iSubItem = DllStructGetData($tNMLVCUSTOMDRAW, "iSubItem")               ; Subitem index
							;DllStructSetData( $tNMLVCUSTOMDRAW, "ClrTextBk", $aColors[$iItem][$iSubItem] ) ; Backcolor of item/subitem
							DllStructSetData($tNMLVCUSTOMDRAW, "ClrTextBk", $aLvColorArray[DllStructGetData($tNMLVCUSTOMDRAW, "dwItemSpec")][DllStructGetData($tNMLVCUSTOMDRAW, "iSubItem")])
							Return $CDRF_NEWFONT     ; $CDRF_NEWFONT must be returned after changing font or colors

					EndSwitch
			EndSwitch
	EndSwitch
	Return $GUI_RUNDEFMSG
	#forceref $hWnd, $iMsg, $wParam
EndFunc   ;==>WM_NOTIFY

;------------------------------------------------------------------------------
; Title...........:	_WM_SIZING
; Description.....:	Resize controls, while drag-resizing
;------------------------------------------------------------------------------
Func _WM_SIZING($hWnd, $iMsg, $wParam, $lParam)
	#forceref $iMsg, $wParam, $lParam
	If $hWnd = $hGUI Then
;~ 		_setControlSizes()
	EndIf

;~ 	Return $GUI_RUNDEFMSG
EndFunc   ;==>_WM_SIZING

;------------------------------------------------------------------------------
; Title...........:	_WM_SIZE
; Description.....:	Resize controls, after size event (such as aero snap)
;------------------------------------------------------------------------------
Func _WM_SIZE($hWnd, $iMsg, $wParam, $lParam)
	#forceref $iMsg, $wParam, $lParam
	If $hWnd = $hGUI Then
		_setControlSizes()
	EndIf

	Local $aClientSize = WinGetClientSize($hGUI)
	Local $aPos = WinGetPos($hGUI)

	;don't save if minimized
	If $aClientSize[0] <> 0 Then
		Json_Put($objOptions, ".Options.WindowPosX", $aPos[0])
		Json_Put($objOptions, ".Options.WindowPosY", $aPos[1])
		Json_Put($objOptions, ".Options.WindowWidth", $aClientSize[0])
		Json_Put($objOptions, ".Options.WindowHeight", $aClientSize[1] + $menuH)
	EndIf

	Local $Json = Json_Encode($objOptions, $Json_PRETTY_PRINT)
	Local $hFile = FileOpen($sOptionsFile, $FO_OVERWRITE)
	FileWrite($hFile, $Json)
	FileClose($hFile)

	Return $GUI_RUNDEFMSG
EndFunc   ;==>_WM_SIZE

;------------------------------------------------------------------------------
; Title...........: _WM_MOVE
; Description.....: Set the resize flag to ignore primary click event when moving GUI
;					This prevents controls from getting selected after a move
; Events..........: Called while dragging window to move
;------------------------------------------------------------------------------
Func _WM_MOVE($hWnd, $Msg, $wParam, $lParam)
	If $hWnd <> $hGUI Then Return $GUI_RUNDEFMSG

	Local $aClientSize = WinGetClientSize($hGUI)
	Local $aPos = WinGetPos($hGUI)

	;don't save if minimized
	If $aClientSize[0] <> 0 Then
		Json_Put($objOptions, ".Options.WindowPosX", $aPos[0])
		Json_Put($objOptions, ".Options.WindowPosY", $aPos[1])
		Json_Put($objOptions, ".Options.WindowWidth", $aClientSize[0])
		Json_Put($objOptions, ".Options.WindowHeight", $aClientSize[1] + $menuH)
	EndIf

	Local $Json = Json_Encode($objOptions, $Json_PRETTY_PRINT)
	Local $hFile = FileOpen($sOptionsFile, $FO_OVERWRITE)
	FileWrite($hFile, $Json)
	FileClose($hFile)

	Return $GUI_RUNDEFMSG
EndFunc   ;==>_WM_MOVE

;------------------------------------------------------------------------------
; Title...........:	_onMaximized
; Description.....:	Resize controls, after maximize event
;------------------------------------------------------------------------------
Func _onMaximized()
	_setControlSizes()
EndFunc   ;==>_onMaximized

;------------------------------------------------------------------------------
; Title...........:	_onRestore
; Description.....:	Resize controls, after restore event
;------------------------------------------------------------------------------
Func _onRestore()
	_setControlSizes()
EndFunc   ;==>_onRestore

;------------------------------------------------------------------------------
; Title...........:	_onExitMain
; Description.....:	call exit routine
; Events..........: close button
;------------------------------------------------------------------------------
Func _onExitMain()
	_ExitMain()
EndFunc   ;==>_onExitMain

;------------------------------------------------------------------------------
; Title...........:	_onMenuExit
; Description.....:	call exit routine
; Events..........: menu item
;------------------------------------------------------------------------------
Func _onMenuExit()
	_ExitMain()
EndFunc   ;==>_onMenuExit

;------------------------------------------------------------------------------
; Title...........:	_ExitMain
; Description.....:	Clean up and exit the program
;------------------------------------------------------------------------------
Func _ExitMain()
	If $syncNow Then
		_Copy_Abort(0)
		$syncNow = False
		_Copy_CloseDll()
	EndIf
	GUIDelete()
	_GDIPlus_Shutdown()
	Exit
EndFunc   ;==>_ExitMain

#EndRegion GUI-creation-and-position


#Region Profile-Handling
Func _LvProfiles_update()
	Local $iSelected = ControlListView($hGUI, "", $lv_profiles, "GetSelected")
	_GUICtrlListView_DeleteAllItems($h_lv_profiles)
	Local $aProfileNames
	Local $oProfiles = _Json_Get($objOptions, ".Profiles")
	If IsObj($oProfiles) Then
		$aProfileNames = Json_ObjGetKeys($oProfiles)
		If IsArray($aProfileNames) Then
			Local $profileName
			For $sName In $aProfileNames
				$profileName = _profileStringDecode($sName)
				GUICtrlCreateListViewItem($profileName, $lv_profiles)
				GUICtrlSetOnEvent(-1, "_onLvProfileSelect")
			Next
			ControlListView($hGUI, "", $lv_profiles, "Select", $iSelected)
		Else
			;no profiles found
		EndIf
	Else
		;no profile section
	EndIf
EndFunc   ;==>_LvProfiles_update

Func _onLvProfileSelect()
	Local $profileName = StringReplace(GUICtrlRead(GUICtrlRead($lv_profiles)), "|", "")
	$profileName = _profileStringEncode($profileName)
	Local $dirLeft = Json_Get($objOptions, '["Profiles"]["' & $profileName & '"]["DirLeft"]')
	Local $dirRight = Json_Get($objOptions, '["Profiles"]["' & $profileName & '"]["DirRight"]')
	Local $radioAction = Json_Get($objOptions, '["Profiles"]["' & $profileName & '"]["radioAction"]')

	GUICtrlSetData($input_Left, $dirLeft)
	GUICtrlSetData($input_Right, $dirRight)
	Switch $radioAction
		Case "MirrorLeft"
			GUICtrlSetState($radio_mirrorLeft, $GUI_CHECKED)
		Case "MirrorRight"
			GUICtrlSetState($radio_mirrorRight, $GUI_CHECKED)
		Case Else
			GUICtrlSetState($radio_sync, $GUI_CHECKED)
	EndSwitch
EndFunc   ;==>_onLvProfileSelect

Func _LvProfiles_save($sName = "")
	Local $profileName = $sName
	If $sName = "" Then
		$profileName = StringReplace(GUICtrlRead(GUICtrlRead($lv_profiles)), "|", "")
	EndIf
	If $profileName = "0" Or $profileName = "" Then
		_setStatusMessage("Select a profile to save.")
		Return SetError(-1, 0, 0)
	EndIf
	$profileName = _profileStringEncode($profileName)

	Local $actionSelect
	If BitAND(GUICtrlRead($radio_mirrorLeft), $GUI_CHECKED) = $GUI_CHECKED Then
		$actionSelect = "MirrorLeft"
	ElseIf BitAND(GUICtrlRead($radio_mirrorRight), $GUI_CHECKED) = $GUI_CHECKED Then
		$actionSelect = "MirrorRight"
	Else
		$actionSelect = "Sync"
	EndIf

	Local $oProfile = Json_Get($objOptions, '["Profiles"]["' & $profileName & '"]')
	If IsObj($oProfile) Then
		$oProfile.Item("DirLeft") = GUICtrlRead($input_Left)
		$oProfile.Item("DirRight") = GUICtrlRead($input_Right)
		$oProfile.Item("radioAction") = $actionSelect
	Else
		Json_Put($objOptions, '["Profiles"]["' & $profileName & '"]["DirLeft"]', GUICtrlRead($input_Left))
		Json_Put($objOptions, '["Profiles"]["' & $profileName & '"]["DirRight"]', GUICtrlRead($input_Right))
		Json_Put($objOptions, '["Profiles"]["' & $profileName & '"]["radioAction"]', $actionSelect)
	EndIf

	Local $Json = Json_Encode($objOptions, $Json_PRETTY_PRINT)
	Local $hFile = FileOpen($sOptionsFile, $FO_OVERWRITE)
	FileWrite($hFile, $Json)
	Return 1
EndFunc   ;==>_LvProfiles_save

Func _LvProfiles_delete($sName = "")
	Local $profileName = $sName
	If $sName = "" Then
		$profileName = StringReplace(GUICtrlRead(GUICtrlRead($lv_profiles)), "|", "")
	EndIf
	If $profileName = "0" Or $profileName = "" Then
		_setStatusMessage("Select a profile to delete.")
		Return SetError(-1, 0, 0)
	EndIf
	$profileName = _profileStringEncode($profileName)
	Local $oProfiles = Json_Get($objOptions, ".Profiles")
	Json_ObjDelete($oProfiles, $profileName)

	Local $Json = Json_Encode($objOptions, $Json_PRETTY_PRINT)
	Local $hFile = FileOpen($sOptionsFile, $FO_OVERWRITE)
	FileWrite($hFile, $Json)
	Return 1
EndFunc   ;==>_LvProfiles_delete

Func _onProfileNewItem()
	Local $aMousePos = MouseGetPos()
	Local $sInput = InputBox("New Profile", "Enter a new profile name.", Default, Default, -1, 125, $aMousePos[0], $aMousePos[1])
	If @error Then Return SetError(-1, 0, 0)
	Local $oProfiles = Json_Get($objOptions, ".Profiles")
	If IsObj($oProfiles) Then
		While 1
			If $oProfiles.Exists($sInput) Then
				$sInput = InputBox("Rename Profile", "Profile already exists!" & @CRLF & "Enter a new profile name.", Default, Default, -1, 140, $aMousePos[0], $aMousePos[1])
				If @error Then Return SetError(-2, 0, 0)
			Else
				ExitLoop
			EndIf
		WEnd
	EndIf
	Local $ret = _LvProfiles_save($sInput)
	If $ret Then
		_LvProfiles_update()
		Local $iSelected = ControlListView($hGUI, "", $lv_profiles, "GetItemCount")
		ControlListView($hGUI, "", $lv_profiles, "Select", $iSelected - 1)
		_setStatusMessage("Profile saved successfully.")
	Else
;~ 		_setStatusMessage("Error: failed to save the profile.")
	EndIf
EndFunc   ;==>_onProfileNewItem

Func _onProfileSave()
	Local $ret = _LvProfiles_save()
	If $ret Then
		_LvProfiles_update()
		_setStatusMessage("Profile saved successfully.")
	Else
;~ 		_setStatusMessage("Error: failed to save the profile.")
	EndIf
EndFunc   ;==>_onProfileSave

Func _onProfileSaveAs()
	_onProfileNewItem()
EndFunc   ;==>_onProfileSaveAs

Func _onProfileDelete()
	_LvProfiles_delete()
	_LvProfiles_update()
EndFunc   ;==>_onProfileDelete

Func _onProfileRename()
	Local $profileName = StringReplace(GUICtrlRead(GUICtrlRead($lv_profiles)), "|", "")
	If $profileName = "0" Then
		MsgBox(1, "Error", "No profile is selected!")
		Return -3
	EndIf
	$profileName = _profileStringEncode($profileName)
	Local $aMousePos = MouseGetPos()
	Local $sInput = InputBox("Rename Profile", "Enter a new profile name.", Default, Default, -1, 125, $aMousePos[0], $aMousePos[1])
	If @error Then Return -1
	Local $oProfiles = Json_Get($objOptions, ".Profiles")
	While 1
		If $oProfiles.Exists($sInput) Then
			$sInput = InputBox("Rename Profile", "Profile already exists!" & @CRLF & "Enter a new profile name.", Default, Default, -1, 140, $aMousePos[0], $aMousePos[1])
			If @error Then Return -2
		Else
			ExitLoop
		EndIf
	WEnd

	If $sInput = "0" Or $sInput = "" Then Return -1
	Local $oProfiles = Json_Get($objOptions, ".Profiles")
	If IsObj($oProfiles) Then
		$oProfiles.Key($profileName) = $sInput
	EndIf

	Local $Json = Json_Encode($objOptions, $Json_PRETTY_PRINT)
	Local $hFile = FileOpen($sOptionsFile, $FO_OVERWRITE)
	FileWrite($hFile, $Json)

	_LvProfiles_update()
	Return 1
EndFunc   ;==>_onProfileRename

Func _profileStringEncode($sName)
	$sName = StringReplace($sName, ".", "{dot}")
	$sName = StringReplace($sName, "\", "{bs}")
	$sName = StringReplace($sName, "/", "{fs}")
	Return $sName
EndFunc   ;==>_profileStringEncode

Func _profileStringDecode($sName)
	$sName = StringReplace($sName, "{dot}", ".")
	$sName = StringReplace($sName, "{bs}", "\")
	$sName = StringReplace($sName, "{fs}", "/")
	Return $sName
EndFunc   ;==>_profileStringDecode

Func _Json_Get(ByRef $obj, $data, $defaultValue = 0)
	Local $val = Json_Get($obj, $data)
	If @error Then
		Return $defaultValue
	Else
		Return $val
	EndIf
EndFunc   ;==>_Json_Get

Func _onLvSelectAll()
	Local $hFocus = ControlGetHandle($hGUI, "", ControlGetFocus($hGUI))
	If $hFocus = $h_lv_results Then
		ControlListView($hGUI, "", $lv_results, "SelectAll")
	Else
		GUISetAccelerators(0)
		Send("^a")
		GUISetAccelerators($aAccelKeys)
	EndIf
EndFunc   ;==>_onLvSelectAll
#EndRegion Profile-Handling


#Region Radio-buttons
;------------------------------------------------------------------------------
; Title...........:	_onRadioMirrorLeft
; Description.....:	update the listview action and row colors
; Event...........: clicking a radio button
;------------------------------------------------------------------------------
Func _onRadioMirrorLeft()
	$aActionOverrides = $aActionFinal
	Local $actionString
	For $i = 0 To UBound($aActionFinal) - 1
		Switch $aActionFinal[$i]
			Case $action_CopyRight
				$aActionOverrides[$i] = $action_DeleteLeft
				$actionString = $actionString_DeleteLeft

			Case $action_CopyLeft
				$actionString = $actionString_CopyLeft

			Case $action_NoChange
				$actionString = $actionString_NoChange

			Case $action_UpdateRight
				$aActionOverrides[$i] = $action_UpdateLeft
				$actionString = $actionString_UpdateLeft

			Case $action_UpdateLeft
				$actionString = $actionString_UpdateLeft

			Case $action_DeleteRight
				$actionString = $actionString_DeleteRight

			Case $action_DeleteLeft
				$actionString = $actionString_DeleteLeft

			Case $action_NotEqual
				$actionString = $actionString_NotEqual

			Case Else
				$actionString = $actionString_Unknown

		EndSwitch

		$aLvDataArray[$i][2] = $actionString
		_setRowColor($lv_results, $i, $aActionOverrides[$i])
	Next

	GUICtrlSendMsg($lv_results, $LVM_SETITEMCOUNT, UBound($aLvDataArray), 0)
EndFunc   ;==>_onRadioMirrorLeft

Func _onRadioMirrorRight()
	$aActionOverrides = $aActionFinal
	Local $actionString
	For $i = 0 To UBound($aActionFinal) - 1
		Switch $aActionFinal[$i]
			Case $action_CopyRight
				$actionString = $actionString_CopyRight

			Case $action_CopyLeft
				$aActionOverrides[$i] = $action_DeleteRight
				$actionString = $actionString_DeleteRight

			Case $action_NoChange
				$actionString = $actionString_NoChange

			Case $action_UpdateRight
				$actionString = $actionString_UpdateRight

			Case $action_UpdateLeft
				$aActionOverrides[$i] = $action_UpdateRight
				$actionString = $actionString_UpdateRight

			Case $action_DeleteRight
				$actionString = $actionString_DeleteRight

			Case $action_DeleteLeft
				$actionString = $actionString_DeleteLeft

			Case $action_NotEqual
				$actionString = $actionString_NotEqual

			Case Else
				$actionString = $actionString_Unknown

		EndSwitch

		$aLvDataArray[$i][2] = $actionString
		_setRowColor($lv_results, $i, $aActionOverrides[$i])
	Next

	GUICtrlSendMsg($lv_results, $LVM_SETITEMCOUNT, UBound($aLvDataArray), 0)
EndFunc   ;==>_onRadioMirrorRight

Func _onRadioSync()
	$aActionOverrides = $aActionFinal
	Local $actionString
	For $i = 0 To UBound($aActionFinal) - 1
		Switch $aActionFinal[$i]
			Case $action_CopyRight
				$actionString = $actionString_CopyRight

			Case $action_CopyLeft
				$actionString = $actionString_CopyLeft

			Case $action_NoChange
				$actionString = $actionString_NoChange

			Case $action_UpdateRight
				$actionString = $actionString_UpdateRight

			Case $action_UpdateLeft
				$actionString = $actionString_UpdateLeft

			Case $action_DeleteRight
				$actionString = $actionString_DeleteRight

			Case $action_DeleteLeft
				$actionString = $actionString_DeleteLeft

			Case $action_NotEqual
				$actionString = $actionString_NotEqual

			Case Else
				$actionString = $actionString_Unknown

		EndSwitch

		$aLvDataArray[$i][2] = $actionString
		_setRowColor($lv_results, $i, $aActionOverrides[$i])
	Next

	GUICtrlSendMsg($lv_results, $LVM_SETITEMCOUNT, UBound($aLvDataArray), 0)
EndFunc   ;==>_onRadioSync
#EndRegion Radio-buttons


#Region Listview-menu-items
;------------------------------------------------------------------------------
; Title...........:	_onLvMenuCopyLeft
; Description.....:	update the listview action and row colors
; Event...........: Listview context menu item
;------------------------------------------------------------------------------
Func _onLvMenuCopyLeft()
	LvMenuAction($action_CopyLeft)
EndFunc   ;==>_onLvMenuCopyLeft

Func _onLvMenuCopyRight()
	LvMenuAction($action_CopyRight)
EndFunc   ;==>_onLvMenuCopyRight

Func _onLvMenuDeleteLeft()
	LvMenuAction($action_DeleteLeft)
EndFunc   ;==>_onLvMenuDeleteLeft

Func _onLvMenuDeleteRight()
	LvMenuAction($action_DeleteRight)
EndFunc   ;==>_onLvMenuDeleteRight

Func _onLvMenuNoChange()
	LvMenuAction($action_NoChange)
EndFunc   ;==>_onLvMenuNoChange

;------------------------------------------------------------------------------
; Title...........:	LvMenuAction
; Description.....:	update the listview action and row colors
;					called by one of the context menu items
;------------------------------------------------------------------------------
Func LvMenuAction($selAction)
	Local $aSelectedIndices = _GUICtrlListView_GetSelectedIndices($h_lv_results, True)
	If Not IsArray($aSelectedIndices) Or $aSelectedIndices[0] = 0 Then
		ConsoleWrite("no selection" & @CRLF)
		Return 1
	EndIf

	Local $index, $actionString
	For $i = 1 To $aSelectedIndices[0]
		$index = $aSelectedIndices[$i]
		$aActionOverrides[$index] = $selAction
		Switch $aActionOverrides[$index]
			Case $action_CopyRight
				$actionString = $actionString_CopyRight

			Case $action_CopyLeft
				$actionString = $actionString_CopyLeft

			Case $action_NoChange
				$actionString = $actionString_NoChange

			Case $action_UpdateRight
				$actionString = $actionString_UpdateRight

			Case $action_UpdateLeft
				$actionString = $actionString_UpdateLeft

			Case $action_DeleteRight
				$actionString = $actionString_DeleteRight

			Case $action_DeleteLeft
				$actionString = $actionString_DeleteLeft

			Case Else
				$actionString = $actionString_Unknown

		EndSwitch

		$aLvDataArray[$index][2] = $actionString
		_setRowColor($lv_results, $index, $aActionOverrides[$index], True)
	Next

	GUICtrlSendMsg($lv_results, $LVM_SETITEMCOUNT, UBound($aLvDataArray), 0)
EndFunc   ;==>LvMenuAction
#EndRegion Listview-menu-items


#Region Analyze
;------------------------------------------------------------------------------
; Title...........:	_onAnalyze
; Description.....:	Analyze the Left and Right folders
; Events..........: Analyze button clicked
;------------------------------------------------------------------------------
Func _onAnalyze()
	Local $timer2 = TimerInit()
	Local $timer = TimerInit()

	GUICtrlSetState($button_analyze, $GUI_DISABLE)

	_LogMessage()
	_LogMessage("---")

	;clear the arrays to delete listview items
	Local $aLvEmpty[0][5]
	$aLvDataArray = $aLvEmpty
	$aLvColorArray = $aLvEmpty
	GUICtrlSendMsg($lv_results, $LVM_SETITEMCOUNT, 0, 0)
	GUICtrlSetData($label_fileCount, "")

	;grab the selected dir names
	Local $sDirLeft = GUICtrlRead($input_Left)
	Local $sDirRight = GUICtrlRead($input_Right)
	Local $iRadioAction

	;save dirs to previous (startup)
	Json_Put($objOptions, ".LastDir.DirLeft", $sDirLeft)
	Json_Put($objOptions, ".LastDir.DirRight", $sDirRight)
	Local $Json = Json_Encode($objOptions, $Json_PRETTY_PRINT)
	Local $hFile = FileOpen($sOptionsFile, $FO_OVERWRITE)
	FileWrite($hFile, $Json)

	;get radio selection
	Local Enum $sel_mirrorLeft, $sel_mirrorRight, $sel_sync
	If BitAND(GUICtrlRead($radio_mirrorLeft), $GUI_CHECKED) = $GUI_CHECKED Then
		$iRadioAction = $sel_mirrorLeft
	ElseIf BitAND(GUICtrlRead($radio_mirrorRight), $GUI_CHECKED) = $GUI_CHECKED Then
		$iRadioAction = $sel_mirrorRight
	Else
		$iRadioAction = $sel_sync
	EndIf

	If Not FileExists($sDirLeft) Then
		_setStatusMessage("Error: Folder on the left does not exist!")
		GUICtrlSetState($button_analyze, $GUI_ENABLE)
		Return
	EndIf
	If Not FileExists($sDirRight) Then
		_setStatusMessage("Error: Folder on the right does not exist!")
		GUICtrlSetState($button_analyze, $GUI_ENABLE)
		Return
	EndIf

	;get Left file/folder list
	_setStatusMessage("Getting Left list of files...")
	Local $aLeftFileList = _GetFileListRec($sDirLeft)
	If @error Then
		ConsoleWrite(@error & @CRLF)
		GUICtrlSetState($button_analyze, $GUI_ENABLE)
		Return
	EndIf

	;get Right file/folder list
	_setStatusMessage("Getting Right list of files...")
	Local $aRightFileList = _GetFileListRec($sDirRight)
	If @error Then
		ConsoleWrite(@error & @CRLF)
		GUICtrlSetState($button_analyze, $GUI_ENABLE)
		Return
	EndIf
	_LogMessage("Get list of files")

	;combine the file lists
	Local $oLeftFileList = ObjCreate("Scripting.Dictionary")
	Local $oRightFileList = ObjCreate("Scripting.Dictionary")
	Local $aFileListCombined[500], $iIndexCombined
	For $i = 0 To UBound($aLeftFileList) - 1
		$oLeftFileList.Item($aLeftFileList[$i][0]) = $aLeftFileList[$i][1]
		If $iIndexCombined > UBound($aFileListCombined) - 1 Then
			ReDim $aFileListCombined[UBound($aFileListCombined) * 2]
		EndIf
		$aFileListCombined[$iIndexCombined] = $aLeftFileList[$i][0]
		$iIndexCombined += 1
	Next
	For $i = 0 To UBound($aRightFileList) - 1
		$oRightFileList.Item($aRightFileList[$i][0]) = $aRightFileList[$i][1]
		If $iIndexCombined > UBound($aFileListCombined) - 1 Then
			ReDim $aFileListCombined[UBound($aFileListCombined) * 2]
		EndIf
		$aFileListCombined[$iIndexCombined] = $aRightFileList[$i][0]
		$iIndexCombined += 1
	Next
	ReDim $aFileListCombined[$iIndexCombined]
	_LogMessage("Create storage arrays")

	;sort the list and remove duplicates
	$aFileListCombined = _ArrayUnique($aFileListCombined, 0, 0, 0, 0)
	_LogMessage("Array unique")

	__ArrayDualPivotSortByFolder($aFileListCombined, 0, UBound($aFileListCombined) - 1)
	_LogMessage("Array sort")

	Local $iSize = UBound($aFileListCombined)
	GUICtrlSetData($label_fileCount, " Files found: " & _StringAddThousandsSep($iSize))


	;preparation complete, now analyze the files
	_setStatusMessage("Analyzing files...")


	;analyze and rebuild the Left and Right arrays for display
	Local $aListTemp[$iSize][3], $aDiffTemp[$iSize]
	$aLeftFileListFinal = $aListTemp
	$aRightFileListFinal = $aListTemp
	$aActionFinal = $aDiffTemp

	;find matching items
	Local $index, $foundInLeft, $foundInRight
	Local $aDateLeft, $aDateRight, $sDateEmpty[6]
	Local $sIndent, $sFileNameLeft, $sFileNameRight, $actionString
	Local $lvData[$iSize][5]

	_LogMessage()

	;open dll
	Local $kernel32_DLL = DllOpen("kernel32.dll")

	For $sFileItem In $aFileListCombined
		;check for matches
		If $oLeftFileList.Exists($sFileItem) Then
			$foundInLeft = True
		Else
			$foundInLeft = False
		EndIf
		If $oRightFileList.Exists($sFileItem) Then
			$foundInRight = True
		Else
			$foundInRight = False
		EndIf
;~ 		_LogMessage("  check exists")


		If $foundInLeft Then
			$aLeftFileListFinal[$index][0] = $sFileItem
			$aLeftFileListFinal[$index][1] = $oLeftFileList.Item($sFileItem)
			$aLeftFileListFinal[$index][2] = FileGetSize($sDirLeft & $sFileItem)
		Else
			$aLeftFileListFinal[$index][0] = ""
			$aLeftFileListFinal[$index][1] = 2
			$aLeftFileListFinal[$index][2] = ""
		EndIf

		If $foundInRight Then
			$aRightFileListFinal[$index][0] = $sFileItem
			$aRightFileListFinal[$index][1] = $oRightFileList.Item($sFileItem)
			$aRightFileListFinal[$index][2] = FileGetSize($sDirRight & $sFileItem)
		Else
			$aRightFileListFinal[$index][0] = ""
			$aRightFileListFinal[$index][1] = 2
			$aRightFileListFinal[$index][2] = ""
		EndIf
;~ 		_LogMessage("  get file size")

		If $foundInLeft And Not $foundInRight Then
			If $oLeftFileList.Item($sFileItem) = 1 Then
				$aDateLeft = $sDateEmpty
				$aLeftFileListFinal[$index][2] = ""
			Else
				$aDateLeft = FileGetTime($sDirLeft & $sFileItem)
			EndIf
			$aActionFinal[$index] = $action_CopyRight
		ElseIf Not $foundInLeft And $foundInRight Then
			If $oRightFileList.Item($sFileItem) = 1 Then
				$aDateRight = $sDateEmpty
				$aRightFileListFinal[$index][2] = ""
			Else
				$aDateRight = FileGetTime($sDirRight & $sFileItem)
			EndIf
			$aActionFinal[$index] = $action_CopyLeft
		ElseIf $foundInLeft And $foundInRight Then
			If $oRightFileList.Item($sFileItem) = 1 Then
				;if DIR, don't care about date
				$aActionFinal[$index] = $action_NoChange
				$aLeftFileListFinal[$index][2] = ""
				$aRightFileListFinal[$index][2] = ""
			Else
				If $oRightFileList.Item($sFileItem) = 1 Then
					$aDateLeft = $sDateEmpty
					$aDateRight = $sDateEmpty
					$aActionFinal[$index] = $action_NoChange
				Else
					$aDateLeft = FileGetTime($sDirLeft & $sFileItem)
					$aDateRight = FileGetTime($sDirRight & $sFileItem)
					For $x = 0 To 5
						If $aDateLeft[$x] <> $aDateRight[$x] Then
							If $aDateLeft[$x] > $aDateRight[$x] Then
								$aActionFinal[$index] = $action_UpdateRight
							Else
								$aActionFinal[$index] = $action_UpdateLeft
							EndIf
							ExitLoop
						Else
							If $aLeftFileListFinal[$index][2] = $aRightFileListFinal[$index][2] Then
								$aActionFinal[$index] = $action_NoChange
							Else
								$aActionFinal[$index] = $action_NotEqual
							EndIf
						EndIf
					Next
				EndIf
			EndIf
		Else
			$aActionFinal[$index] = $action_Unknown
		EndIf
;~ 		_LogMessage("  compare time")


		;;build the listview

		;set indent levels
		$sIndent = ""
		StringReplace($sFileItem, "\", "")
		For $i = 1 To @extended - 1
			$sIndent &= "     "
		Next

		;get left side file name
		$sFileNameLeft = StringRegExpReplace($aLeftFileListFinal[$index][0], '(?m).*\\(.*?)$', '$1')
		If $aLeftFileListFinal[$index][1] = 1 Then
			$sFileNameLeft = $sFileNameLeft & "\"
		EndIf

		;get right side file name
		$sFileNameRight = StringRegExpReplace($aRightFileListFinal[$index][0], '(?m).*\\(.*?)$', '$1')
		If $aRightFileListFinal[$index][1] = 1 Then
			$sFileNameRight = $sFileNameRight & "\"
		EndIf

		;get action text
		Switch $aActionFinal[$index]
			Case $action_CopyRight
				$actionString = $actionString_CopyRight
			Case $action_CopyLeft
				$actionString = $actionString_CopyLeft
			Case $action_NoChange
				$actionString = $actionString_NoChange
			Case $action_UpdateRight
				$actionString = $actionString_UpdateRight
			Case $action_UpdateLeft
				$actionString = $actionString_UpdateLeft
			Case $action_DeleteLeft
				$actionString = $actionString_DeleteLeft
			Case $action_DeleteRight
				$actionString = $actionString_DeleteRight
			Case $action_NotEqual
				$actionString = $actionString_NotEqual
			Case Else
				$actionString = $actionString_Unknown
		EndSwitch

		;update lv data array (used to create items later)
		$lvData[$index][0] = $sIndent & $sFileNameLeft
		$lvData[$index][1] = _StringAddThousandsSep($aLeftFileListFinal[$index][2])
		$lvData[$index][2] = $actionString
		$lvData[$index][3] = $sIndent & $sFileNameRight
		$lvData[$index][4] = _StringAddThousandsSep($aRightFileListFinal[$index][2])

		$index += 1

;~ 		_LogMessage("  build listview")
	Next

	DllClose($kernel32_DLL)
	_LogMessage("File compare")

	$aActionOverrides = $aActionFinal
	$aLvDataArray = $lvData
	Local $aColorsTemp[UBound($aLvDataArray)][5]
	$aLvColorArray = $aColorsTemp

	;set the action and background colors
	Switch $iRadioAction
		Case $sel_sync
			_onRadioSync()
		Case $sel_mirrorLeft
			_onRadioMirrorLeft()
		Case $sel_mirrorRight
			_onRadioMirrorRight()
	EndSwitch

	_LogMessage("Set action and colors")

	;force the listview to update itself
	GUICtrlSendMsg($lv_results, $LVM_SETITEMCOUNT, UBound($aLvDataArray), 0)

	_LogMessage("Update listview")

	ConsoleWrite("Done: " & TimerDiff($timer2) & @CRLF & @CRLF)
	_setStatusMessage("Folders analyzed successfully")
	GUICtrlSetState($button_analyze, $GUI_ENABLE)
EndFunc   ;==>_onAnalyze

;------------------------------------------------------------------------------
; Title...........:	_setRowColor
; Description.....:	update the row color array based on selected action
;------------------------------------------------------------------------------
Func _setRowColor($lv, $index, $action, $override = False)
	Local $colorDefault = 0xFFFFFF
	Local $colorCopy = 0x74ff74
	Local $colorUpdate = 0xffe1bd
	Local $colorNoChange = 0xF9F9F9
	Local $colorDelete = 0xb3b3ff
	Local $colorOverride = 0x00FFFF
	Local $colorUnknown = 0x00FFA5
	Local $colorError = 0x0000FF

	If $override Then
		Switch $action
			Case $action_DeleteLeft
				$aLvColorArray[$index][0] = $colorDelete
				$aLvColorArray[$index][1] = $colorDelete
				$aLvColorArray[$index][2] = $colorDelete
				$aLvColorArray[$index][3] = $colorDefault
				$aLvColorArray[$index][4] = $colorDefault
			Case $action_DeleteRight
				$aLvColorArray[$index][0] = $colorDefault
				$aLvColorArray[$index][1] = $colorDefault
				$aLvColorArray[$index][2] = $colorDelete
				$aLvColorArray[$index][3] = $colorDelete
				$aLvColorArray[$index][4] = $colorDelete
			Case Else
				$aLvColorArray[$index][0] = $colorOverride
				$aLvColorArray[$index][1] = $colorOverride
				$aLvColorArray[$index][2] = $colorOverride
				$aLvColorArray[$index][3] = $colorOverride
				$aLvColorArray[$index][4] = $colorOverride
		EndSwitch
	Else
		Switch $action
			Case $action_CopyLeft
				$aLvColorArray[$index][0] = $colorDefault
				$aLvColorArray[$index][1] = $colorDefault
				$aLvColorArray[$index][2] = $colorCopy
				$aLvColorArray[$index][3] = $colorCopy
				$aLvColorArray[$index][4] = $colorCopy
			Case $action_CopyRight
				$aLvColorArray[$index][0] = $colorCopy
				$aLvColorArray[$index][1] = $colorCopy
				$aLvColorArray[$index][2] = $colorCopy
				$aLvColorArray[$index][3] = $colorDefault
				$aLvColorArray[$index][4] = $colorDefault
			Case $action_UpdateLeft
				$aLvColorArray[$index][0] = $colorUpdate
				$aLvColorArray[$index][1] = $colorUpdate
				$aLvColorArray[$index][2] = $colorUpdate
				$aLvColorArray[$index][3] = $colorUpdate
				$aLvColorArray[$index][4] = $colorUpdate
			Case $action_UpdateRight
				$aLvColorArray[$index][0] = $colorUpdate
				$aLvColorArray[$index][1] = $colorUpdate
				$aLvColorArray[$index][2] = $colorUpdate
				$aLvColorArray[$index][3] = $colorUpdate
				$aLvColorArray[$index][4] = $colorUpdate
			Case $action_DeleteLeft
				$aLvColorArray[$index][0] = $colorDelete
				$aLvColorArray[$index][1] = $colorDelete
				$aLvColorArray[$index][2] = $colorDelete
				$aLvColorArray[$index][3] = $colorDefault
				$aLvColorArray[$index][4] = $colorDefault
			Case $action_DeleteRight
				$aLvColorArray[$index][0] = $colorDefault
				$aLvColorArray[$index][1] = $colorDefault
				$aLvColorArray[$index][2] = $colorDelete
				$aLvColorArray[$index][3] = $colorDelete
				$aLvColorArray[$index][4] = $colorDelete
			Case $action_NoChange
				$aLvColorArray[$index][0] = $colorNoChange
				$aLvColorArray[$index][1] = $colorNoChange
				$aLvColorArray[$index][2] = $colorNoChange
				$aLvColorArray[$index][3] = $colorNoChange
				$aLvColorArray[$index][4] = $colorNoChange
			Case $action_NotEqual, $action_Unknown
				$aLvColorArray[$index][0] = $colorUnknown
				$aLvColorArray[$index][1] = $colorUnknown
				$aLvColorArray[$index][2] = $colorUnknown
				$aLvColorArray[$index][3] = $colorUnknown
				$aLvColorArray[$index][4] = $colorUnknown
			Case $action_Error
				$aLvColorArray[$index][0] = $colorError
				$aLvColorArray[$index][1] = $colorError
				$aLvColorArray[$index][2] = $colorError
				$aLvColorArray[$index][3] = $colorError
				$aLvColorArray[$index][4] = $colorError
		EndSwitch
	EndIf
EndFunc   ;==>_setRowColor
#EndRegion Analyze


#Region Helpers
;------------------------------------------------------------------------------
; Title...........:	_StringAddThousandsSep
; Description.....:	add commas to number string
; Author..........: Melkey
; Source..........: from https://www.autoitscript.com/forum/topic/113446-replacement-for-_stringaddthousandssep/?tab=comments#comment-793903
;------------------------------------------------------------------------------
Func _StringAddThousandsSep($sString, $sThousands = ",", $sDecimal = ".")
	Local $aNumber, $sLeft, $sResult = "", $iNegSign = "", $DolSgn = ""
	If Number(StringRegExpReplace($sString, "[^0-9\-.+]", "\1")) < 0 Then $iNegSign = "-" ; Allows for a negative value
	If StringRegExp($sString, "\$") And StringRegExpReplace($sString, "[^0-9]", "\1") <> "" Then $DolSgn = "$" ; Allow for Dollar sign
	$aNumber = StringRegExp($sString, "(\d+)\D?(\d*)", 1)
	If UBound($aNumber) = 2 Then
		$sLeft = $aNumber[0]
		While StringLen($sLeft)
			$sResult = $sThousands & StringRight($sLeft, 3) & $sResult
			$sLeft = StringTrimRight($sLeft, 3)
		WEnd
		$sResult = StringTrimLeft($sResult, 1) ; Strip leading thousands separator
		If $aNumber[1] <> "" Then $sResult &= $sDecimal & $aNumber[1] ; Add decimal
	EndIf
	Return $iNegSign & $DolSgn & $sResult ; Adds minus or "" (nothing)and Adds $ or ""
EndFunc   ;==>_StringAddThousandsSep

;------------------------------------------------------------------------------
; Title...........:	_onLeftSelect
; Description.....:	call function to open folder select dialog
; Events..........: button click
;------------------------------------------------------------------------------
Func _onLeftSelect()
	_selectFolder($input_Left, $input_Right, "left")
EndFunc   ;==>_onLeftSelect

;------------------------------------------------------------------------------
; Title...........:	_onRightSelect
; Description.....:	call function to open folder select dialog
; Events..........: button click
;------------------------------------------------------------------------------
Func _onRightSelect()
	_selectFolder($input_Right, $input_Left, "right")
EndFunc   ;==>_onRightSelect

;------------------------------------------------------------------------------
; Title...........:	_selectFolder
; Description.....:	Open the folder select dialog and set the associated input
;------------------------------------------------------------------------------
Func _selectFolder($iCtrlIDSelected, $iCtrlIDOther, $sText)
	Local $sStartDir = GUICtrlRead($iCtrlIDSelected)
	If $sStartDir <> "" Then
		Local $aStartDir = StringRegExp($sStartDir, '(.*\\)', $STR_REGEXPARRAYGLOBALMATCH)
		$sStartDir = $aStartDir[0]
		ConsoleWrite($sStartDir & @CRLF)
	Else
		$sStartDir = GUICtrlRead($iCtrlIDOther)
		If $sStartDir <> "" Then
			Local $aStartDir = StringRegExp($sStartDir, '(.*\\)', $STR_REGEXPARRAYGLOBALMATCH)
			$sStartDir = $aStartDir[0]
			ConsoleWrite($sStartDir & @CRLF)
		EndIf
	EndIf

	Local $sDirSelected = FileSelectFolder("Select " & $sText & " folder", $sStartDir)
	If $sDirSelected <> "" Then
		GUICtrlSetData($iCtrlIDSelected, $sDirSelected)
	EndIf
EndFunc   ;==>_selectFolder

;------------------------------------------------------------------------------
; Title...........:	_onDropped
; Description.....:	Check file/folder dropped onto input box, only allow folders
;------------------------------------------------------------------------------
Func _onDropped()
;~ 	Local $data = GUICtrlRead(@GUI_DropId)
	Local $filename = @GUI_DragFile
	Local $attribute = FileGetAttrib($filename) ; Retrieve the file/dir attributes
	If Not StringInStr($attribute, "D") Then ; If the attribute string contains the letter 'D' then it is a DIR.
		GUICtrlSetData(@GUI_DropId, StringRegExpReplace($filename, "(?m)(.*)\\.*?$", "$1"))
	Else
		GUICtrlSetData(@GUI_DropId, $filename)
	EndIf
EndFunc   ;==>_onDropped

;------------------------------------------------------------------------------
; Title...........:	_setStatusMessage
; Description.....:	set the statusbar message with optional delay off
;------------------------------------------------------------------------------
Func _setStatusMessage($sMessage = "", $hold = False)
	Static Local $timer = -1, $startTimer
	Local $statusDelay = 3000

	If $sMessage = "" Then
		If $timer <> -1 Then
			If TimerDiff($timer) > $statusDelay Then
				_GUICtrlStatusBar_SetText($hStatusbar, "")
				$timer = -1
			EndIf
		EndIf
	Else
		_GUICtrlStatusBar_SetText($hStatusbar, $sMessage)
		If $hold Then
			$timer = -1
		Else
			$startTimer = True
			$timer = TimerInit()
		EndIf
	EndIf
EndFunc   ;==>_setStatusMessage

Func _LogMessage($sMessage = -1)
	Static Local $timer

	If $sMessage = -1 Then
		$timer = TimerInit()
	Else
		Local $iTimeMs = TimerDiff($timer)
		Local $iHr, $iMin, $iSec, $iMs
		_TicksToTime($iTimeMs, $iHr, $iMin, $iSec)
		$iMs = $iTimeMs - ($iHr * 3600000 + $iMin * 60000 + $iSec * 1000)
		Local $sDiff = StringFormat("%.2d:%.2d:%.2d:%08.4f", ($iHr), ($iMin), ($iSec), ($iMs))
		If $sMessage = "---" Then
			ConsoleWrite("------------------------------------------" & @CRLF)
		ElseIf $sMessage = @CRLF Then
			ConsoleWrite(@CRLF)
		Else
			ConsoleWrite("(" & $sDiff & ")" & @TAB & $sMessage & @CRLF)
		EndIf
		$timer = TimerInit()
	EndIf
EndFunc   ;==>_LogMessage
#EndRegion Helpers


#Region Icon-data-and-retrieval
;------------------------------------------------------------------------------
; Title...........:	GetIconData
; Description.....:	decode the icon data and convert to hIcon
;------------------------------------------------------------------------------
Func GetIconData($iconSel)
	Local $icondData
	Switch $iconSel
		Case 0 ; NEW
			$icondData = '' & _
					'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABHNCSVQICAgIfAhk' & _
					'iAAAAAlwSFlzAAAKVQAAClUB2A38aQAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3Nj' & _
					'YXBlLm9yZ5vuPBoAAAFMSURBVDiNlZO9TgJBFIXPuTMj/m0WQ1QKIBT6IGthLEwM' & _
					'j6E1USqsfAMjr0FlKb6FhWIEjdmGhEVWYmMsYGWWHUycaubk3HtPvtwhHKcYeB1/' & _
					'a2PP1qLh5DG8iw4WvdrVYN3Llc+bhyVbu2refrm84hJdh6BTdyZYLFRils76M0FS' & _
					'rGhAdwDoYuDdb/prZVsslfI+IbNiDSUG1epuXtdMFyCEBECMRp997td2XhuXRylg' & _
					'pEBxOjlJoESn3zSoX7TeMgx+iy3jvMks0exOMg1RqJxGRQMRk2kMWA3SxbZRZ7UE' & _
					'LAgWA7/j+auVhDpJlCvb/tnpccGedt1qD3ovYQTI1Adi+BH3dNiJghBRisP3iTwp' & _
					'MQWbwXP3ffjQ7qXWG1iySCShuDInzykw5x44VRB6gfyynXMnADPQ/vUX4vGkX2/c' & _
					'5BJYEGIcT/ou7w9MzD6ryePdiwAAAABJRU5ErkJggg=='
		Case 1 ; SAVE
			$icondData = '' & _
					'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABHNCSVQICAgIfAhk' & _
					'iAAAAAlwSFlzAAAKVwAAClcBp/M/4AAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3Nj' & _
					'YXBlLm9yZ5vuPBoAAAHjSURBVDiNjZFNaxNRFIafO8l0Mkm0CcFpoxkoYrtIRVJw' & _
					'UbB0Zaw2de0fCN0kmKX/QGhdBSsUf4ErJWLFhRtpxYWLgBZUkJpqJLZuajV08nWv' & _
					'izTNN/pu7r0v530451zB5GQO276EgFO7X0Px+flfQso6PRJCaH+U0je3Nv1lw1Ng' & _
					'ZyfB0dE3wbX4F215eQLg/MZzcg/WerPcujKHZVl4xsd5D3y/voBcvfeOfH5JQ9P6' & _
					'Ar0Kh8NcjEap1Y8b03VE5jbY9ob7n+ljNeo9U7lcOqGQ+V+A/b0feL1eNI+n09YB' & _
					'ugBVx+Hh+nof4HJiCdVoUDFNatvbTVMpHVBdgOLiDe4fHva3MGa174nF5imlDlTb' & _
					'ACkxPn7AXakMHUVTiqphUInNAKobEHn2lLXpc5wNhoYCGij2Cp9J5x5TvLqgKzp2' & _
					'EKg63JyNDQ23VB3zE/j0kmJzB22A04A7L/JMWGeYmY7y6vWbgYCa41BBb43QBnhO' & _
					'B1i5u3pSOBtPDASUy2UepdMgVf83tpTNZimVSl2ebdukUqlOSwwF+Hw+gsFgs0oI' & _
					'MpkMhmEM7MiNlACM+P0nZjKZHFgM4EiJ2zSbDyURTE09IRKJXRgdHZmzLM1VrztD' & _
					'08DvWk17e3BgFn7uK1XY3foLS4ebi0+2J4UAAAAASUVORK5CYII='
		Case 2 ; DELETE
			$icondData = '' & _
					'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABHNCSVQICAgIfAhk' & _
					'iAAAAAlwSFlzAAAKIwAACiMBfBPMxQAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3Nj' & _
					'YXBlLm9yZ5vuPBoAAAKNSURBVDiNpZNLTxNhFIbPXCDTrzOdKdAbSjpjFyRIaTVp' & _
					'iJsWosGU8BMsBRcksua3uCSUCom/wLjRQsKqIm1NQFBkwIDUXphxhoHO5asLJcEE' & _
					'iAlnfZ7kvOfJS0wD+PyC8JYA0A8VJZ0HOIUbJgPgviMIb9oA7E9FeUqNCcLai4mJ' & _
					'+EAodPfL0VE6YppLFQDrOjjAcSvZ0dHEo2g0tHlwMEYSGOtnqopFjqOyiUQ8gNBK' & _
					'BsB9JYzQSjaRiIscR52pKiYw1qn7rdbrzcPDdB/L+u+xLB1xu4Nfm83xftte3gAw' & _
					'AQBmAFAPwxSmBgfjktdLy40GXlhb+1TXtCfExYLH5Vp9NjAQizAM/b3ZxAu7u+W6' & _
					'aaZoAIfv7CxMRiIPxK4uWj4/x4tbW5WGYSTnATTi4sRZABYxTGEyHI6JFEXLqooX' & _
					'q9UyAOCM3x+TBIGWbRu/2t8v661W8iWADgBAXM45C8Cijo5CxueLi+02tWcYDgCA' & _
					'hBAlE4STr9VKhmWNXMDXapqj6Q8lj8dSGKatMEy77PE4czT98arnkjc5/5+5dQTq' & _
					'H5hhCpOSFJMQovdNE+dPTioVy/oRZlm/JAh0mOeDn3V9POY4S8W/ikmAPxrdCL3L' & _
					'RqNxqaeHljHG+Vqt0rTt1C/bTubr9dIuxrbk85HTQ0MxFqHV5wAcAAA1A4B4jitM' & _
					'DQ8/FL1eSlZVnN/erjRMMzkPoK0DWIOOs7Sj6+m+7m6/FAxSkWAwsF2rpftNc5l6' & _
					'7PG8nx4ZSYg+H7XXbDq5jY1S9fw8mbuUcx3Aitj28jdFSYcDgYDY20tJoZB/5/g4' & _
					'RQJJul08T8qa5uSKxVLVMFJXNTIPcFo1jFSuWCzJmua4eJ4kSNJN3LbOvwFR0zR/' & _
					'yIVS+wAAAABJRU5ErkJggg=='
		Case 3 ; SEARCH
			$icondData = '' & _
					'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABHNCSVQICAgIfAhk' & _
					'iAAAAAlwSFlzAAAPgAAAD4ABMkKt4wAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3Nj' & _
					'YXBlLm9yZ5vuPBoAAAD5SURBVDiNndE9L0RBFMbx390lIkLUEpUoVBo0EtFsoVIs' & _
					'ofUFfAy1aCUKjUazicaqFD6DToioSRSbeAnFnI11171711PNzJnnf87Mk/lRhi1s' & _
					'4C32NbzgCI9KNIlT7ISxVzM4RrPInIV5vqwDDrD+V2EbuwPMpOecoZ4vnOgfu0h7' & _
					'WMtT3/FVEXCD5TygqhlepQ//BRgZAjCHhzzgWYqqijbRzh/OSjkP0gIOi4pNKeda' & _
					'ifkuuo/3FrqZ3uIzIFPoYBSL2MdK3G3E+hwf9OdfxyqWMI17XOIpOrcCciX9R6fw' & _
					'wQUaw4UU/TUmhgWISdoBaf0H0IW00PgGTogouhCe70QAAAAASUVORK5CYII='
	EndSwitch

	Return _getMemoryAsIcon(_Base64Decode($icondData))
EndFunc   ;==>GetIconData

;------------------------------------------------------------------------------
; Title...........:	GetPicData
; Description.....:	decode the Pic data
;------------------------------------------------------------------------------
Func GetPicData($iconSel)
	Local $icondData
	Switch $iconSel
		Case 0    ; LOGO
			$icondData = "" & _
					"iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAMAAADXqc3KAAACdlBMVEUAAAAfJCUA" & _
					"AAADAAMODw4AAAA5Pz0dHyMxMUcXMxcnOykAAACGhs+YmM6ensiDg5xaWog/iEKC" & _
					"gY1YWF8rRy5XV3BxcnpIWEw3XzpISE0dWh0jIy8xMTkoPSoXORcdHR8bGyQLAAoJ" & _
					"Bwlo525Z41445Tmrq+Orq+Ofn+GTk+BG4Um3t91k0muKityZmd7CydPExNh5ecCa" & _
					"zKc7vj0wxDF20X9lv2xe02OfwqxjyGk5zTqxrsihsq2srMoyzzOUlMotsS1kwmp7" & _
					"e8Shrq99tIdzt3xaXXlJdE9vb69TwFelpbpeumWMjL0tuS0utC98fLeCgrhupXZy" & _
					"nHptbaxra6iCgqM+qUFZWXNemmR9fadzjnw2kjkomCg7Yj50dJVpaZR1g35SlVce" & _
					"ZR54eIlkfmtzdH9ra347O1k7PU1HR11GekpeY2Z4eIIieSJeXoFUVHFUVGRpaXJG" & _
					"RmtYVmUvbzFFRVRFRWA+R0ILKApCQlQoKC4NDQqrq/+mpv+/v/+1tf/i4v/d3f/Y" & _
					"2P/Fxf+ysv+vr/+jo/+5uf5X/FtO+1FA/0LS0v/Pz//MzP/IyP+srPDY7uvI7tq+" & _
					"8c6r8bmj87CR9ZyK9ZSE+I1v+nVq+nFm+2xi+mhc+mFH/krn5//U1P+8vP+fn/6g" & _
					"oP3W1vvd3fjg4PbV1fbc3PXFxfWmpvXh7vSmpvS9vfOWlu/Dw+vX3OrU7+fHx+XR" & _
					"2+OkpOPP7eHO3+DLy97G5tfExdbB79HA6tC48Me31cey8MGw8r6q4rij7LCc8Kib" & _
					"86eV9KGU7J9/4Yh+94d494B48IB0+3tz+Hpx7Xha9l5Z6l1T+1ZK/E1E/0Y99D6S" & _
					"FbXQAAAAfHRSTlMAJAcEGAFFQj87Ngv6+Ou/npiEfX1ya2RgWlM7OTArKh0QDv7+" & _
					"/v38/Pz8+/v6+fn4+Pj4+Pf39/X19fT08fHq6uno4+Pj4+Pi4eDg3dnY1NPSysnH" & _
					"xcW/vrq4sbGwraupp6ejn5aTk4+NhoWEgoF1b2xsaWJVTEtJQikKH6H/FgAAAYRJ" & _
					"REFUKM9iQALVTAzYgXS6AnYJDsNgDuw6eBY6FWEax1ImoFnbpiuAZhxjXoDeypra" & _
					"2ra1IRysCGElUZ/++vrW1pqa2rZFHuwI5XFqdQva6+tbQDI8btliYoUyIPGq0DXN" & _
					"dXV17Zv7W1pW+2fYbN20focwyFb+Fb19zc3qkZm87RrxjOJ7NkxYpS/JwMCctq67" & _
					"p7ePT5Rbhtc+n5spatfELROc5RkY2Gyb5nf38JUyMLCHVzIwiBtN2jlxewITYAzM" & _
					"ycsam5q0CxgYZGPZgcpcJu/dPcminIFB3quhsXExPzezhLcZF7OE68Epk/ftTwH6" & _
					"X9JgXkPDxmIWQfOTliKJplMPH5pyIIgR6CbhbUuXL7EWCTs7a6ayzonp06ZNPRLI" & _
					"BfIEW46QUKq7cWfHnNmzZpw6fmz60QhZuM8rHLrmdp4DSs2ccdoxiwURVMxSfhcv" & _
					"nO/sOKPiKcjFihK4cjGqXXO1oksYocIIoJhr12XFiS36WKV8TRASqMYlwSQwjFNC" & _
					"4gAAd8l1qUwtRpYAAAAASUVORK5CYII="

	EndSwitch

	Return _Base64Decode($icondData)
EndFunc   ;==>GetPicData

;------------------------------------------------------------------------------
; Title...........:	_getMemoryAsPic
; Description.....:	convert binary memory bitmap to PIC
;------------------------------------------------------------------------------
Func _setMemoryAsPic($idPic, $name)
	$hBmp = _GDIPlus_BitmapCreateFromMemory(Binary($name), 1)
	_WinAPI_DeleteObject(GUICtrlSendMsg($idPic, 0x0172, 0, $hBmp))
	_WinAPI_DeleteObject($hBmp)
	Return 0
EndFunc   ;==>_setMemoryAsPic

;------------------------------------------------------------------------------
; Title...........:	_getMemoryAsIcon
; Description.....:	convert binary memory bitmap to HICON
;------------------------------------------------------------------------------
Func _getMemoryAsIcon($name)
	$Bmp = _GDIPlus_BitmapCreateFromMemory(Binary($name))
	$hIcon = _GDIPlus_HICONCreateFromBitmap($Bmp)
	_GDIPlus_ImageDispose($Bmp)
	Return $hIcon
EndFunc   ;==>_getMemoryAsIcon

;encode/decode functions by trancexx
;https://www.autoitscript.com/forum/topic/81332-_base64encode-_base64decode/
Func _Base64Encode($input)

	$input = Binary($input)

	Local $struct = DllStructCreate("byte[" & BinaryLen($input) & "]")

	DllStructSetData($struct, 1, $input)

	Local $strc = DllStructCreate("int")

	Local $a_Call = DllCall("Crypt32.dll", "int", "CryptBinaryToString", _
			"ptr", DllStructGetPtr($struct), _
			"int", DllStructGetSize($struct), _
			"int", 1, _
			"ptr", 0, _
			"ptr", DllStructGetPtr($strc))

	If @error Or Not $a_Call[0] Then
		Return SetError(1, 0, "") ; error calculating the length of the buffer needed
	EndIf

	Local $a = DllStructCreate("char[" & DllStructGetData($strc, 1) & "]")

	$a_Call = DllCall("Crypt32.dll", "int", "CryptBinaryToString", _
			"ptr", DllStructGetPtr($struct), _
			"int", DllStructGetSize($struct), _
			"int", 1, _
			"ptr", DllStructGetPtr($a), _
			"ptr", DllStructGetPtr($strc))

	If @error Or Not $a_Call[0] Then
		Return SetError(2, 0, "") ; error encoding
	EndIf

	Return DllStructGetData($a, 1)

EndFunc   ;==>_Base64Encode

Func _Base64Decode($input_string)

	Local $struct = DllStructCreate("int")

	$a_Call = DllCall("Crypt32.dll", "int", "CryptStringToBinary", _
			"str", $input_string, _
			"int", 0, _
			"int", 1, _
			"ptr", 0, _
			"ptr", DllStructGetPtr($struct, 1), _
			"ptr", 0, _
			"ptr", 0)

	If @error Or Not $a_Call[0] Then
		Return SetError(1, 0, "") ; error calculating the length of the buffer needed
	EndIf

	Local $a = DllStructCreate("byte[" & DllStructGetData($struct, 1) & "]")

	$a_Call = DllCall("Crypt32.dll", "int", "CryptStringToBinary", _
			"str", $input_string, _
			"int", 0, _
			"int", 1, _
			"ptr", DllStructGetPtr($a), _
			"ptr", DllStructGetPtr($struct, 1), _
			"ptr", 0, _
			"ptr", 0)

	If @error Or Not $a_Call[0] Then
		Return SetError(2, 0, "") ; error decoding
	EndIf

	Return DllStructGetData($a, 1)

EndFunc   ;==>_Base64Decode
#EndRegion Icon-data-and-retrieval


#Region About-dialog
;------------------------------------------------------------------------------
; Title...........:	_onMenuAbout
; Description.....:	create and display the about dialog
;------------------------------------------------------------------------------
Func _onMenuAbout()
	$w = 275
	$h = 196

	$hGuiAbout = GUICreate("About " & "SyncFiles", $w, $h, Default, Default, $WS_CAPTION, -1, $hGUI)
	GUISetOnEvent($GUI_EVENT_CLOSE, "_onExitAbout")

	; top section

	GUICtrlCreateLabel("", 0, 0, $w, $h - 32)
	GUICtrlSetBkColor(-1, 0xFFFFFF)
	GUICtrlSetState(-1, $GUI_DISABLE)

	GUICtrlCreateLabel("", 0, $h - 32, $w, 1)
	GUICtrlSetBkColor(-1, 0x000000)

	;icon
	Local $pic = GUICtrlCreatePic("", 10, 11, 24, 24)
	_setMemoryAsPic($pic, GetPicData(0))

	GUICtrlCreateLabel("SyncFiles", 42, 11, $w - 15)
	GUICtrlSetFont(-1, 16, 800)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)

	GUICtrlCreateLabel("Author:", 10, 45, 60, -1)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlCreateLabel("kurtykurtyboy", 60, 45, 65, -1)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)

	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlCreateLabel("Version:", 10, 61, 60, -1)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlCreateLabel($version, 60, 61, 65, -1)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)

	GUICtrlCreateLabel("Date:", 10, 77, 60, -1)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlCreateLabel($date, 60, 77, 65, -1)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)

	GUICtrlCreateLabel("License:", 10, 93, 60, -1)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlCreateLabel("GNU GPL v3", 60, 93, 65, -1)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)

	Local $desc = "SyncFiles is an easy to use file and folder synchronization tool."
	GUICtrlCreateLabel($desc, 10, 121, $w - 16, 50)
	GUICtrlSetFont(-1, 9)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)

	; bottom section
	GUICtrlCreateButton("OK", $w - 55, $h - 27, 50, 22)
	GUICtrlSetOnEvent(-1, "_onExitAbout")

	GUISetState(@SW_SHOW, $hGuiAbout)
EndFunc   ;==>_onMenuAbout

;------------------------------------------------------------------------------
; Title...........:	_onExitAbout
; Description.....:	close the about dialog
;------------------------------------------------------------------------------
Func _onExitAbout()
	GUIDelete($hGuiAbout)
	GUISetState(@SW_SHOWNORMAL, $hGUI)
EndFunc   ;==>_onExitAbout
#EndRegion About-dialog
