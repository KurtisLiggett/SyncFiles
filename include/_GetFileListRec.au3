#include-Once

#include <WinAPIFiles.au3>
#include <Array.au3>

;------------------------------------------------------------------------------------------------------------------------------------------------
; Title...........:	_GetFileListRec
; Description.....:	Get array of directory and file names (recursive)
; Syntax..........: __ArrayDualPivotSort ( ByRef $aArray, $iPivot_Left, $iPivot_Right [, $bLeftMost = True ] )
; Parameters .....: $sFilePath  - The base file path to search
;                   $optRelative  - return relative paths
; Return values ..: 2D Array of files
;						[0] - filename
;						[1] - 0=file, 1=directory
; Author..........: kurtykurtyboy
;					Heavily influenced by guinness. Idea by Belini and others.
; Source..........: https://www.autoitscript.com/forum/topic/143676-list-files-in-folders-and-subfolders-quickly/?tab=comments#comment-1232304
;------------------------------------------------------------------------------------------------------------------------------------------------
Func _GetFileListRec($sFilePath, $optRelative = True)
	Local $aFileList[100][2]
	Local $arrayIndex = 0

	Local $sFilePath2 = _WinAPI_PathAddBackslash($sFilePath) ; Add backslash
	If _WinAPI_PathIsDirectory($sFilePath2) <> $FILE_ATTRIBUTE_DIRECTORY Then
		Return SetError(1, 0, '')
	EndIf

	_GetFileListRec_Process($aFileList, $arrayIndex, $sFilePath, $sFilePath, True)

	ReDim $aFileList[$arrayIndex][2]
	Return $aFileList
EndFunc   ;==>_GetFileListRec

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: _GetFileListRec_Process
; Description ...: Takes a 2D array and searches for files and folders, adding to the array as needed.
;				   Calls itself recursively when needed.
; Return values .: None
; Author ........: kurtykurtyboy
; ===============================================================================================================================
Func _GetFileListRec_Process(ByRef $aFileList, ByRef $iIndex, $sBasePath, $sFilePath, $optRelative = True)
	$sFilePath = _WinAPI_PathAddBackslash($sFilePath) ; Add backslash

	Local $hFileFind = FileFindFirstFile($sFilePath & '*')
	If $hFileFind = -1 Then ; File not found
		Return SetError(2, 0, '')
	EndIf

	Local $sFileName = ''
	While True
		If $iIndex > UBound($aFileList) - 1 Then
			ReDim $aFileList[$iIndex + UBound($aFileList) * 2][2]
		EndIf

		$sFileName = FileFindNextFile($hFileFind)
		If @error Then
			ExitLoop
		EndIf

		If @extended Then ; Is directory.
			If $optRelative Then
				$aFileList[$iIndex][0] = StringReplace($sFilePath & $sFileName, $sBasePath, "", 0, 1)
			Else
				$aFileList[$iIndex][0] = $sFilePath & $sFileName
			EndIf
			$aFileList[$iIndex][1] = 1 ;dir
			$iIndex += 1
			_GetFileListRec_Process($aFileList, $iIndex, $sBasePath, $sFilePath & $sFileName, $optRelative)
		Else
			If $optRelative Then
				$aFileList[$iIndex][0] = StringReplace($sFilePath & $sFileName, $sBasePath, "", 0, 1)
			Else
				$aFileList[$iIndex][0] = $sFilePath & $sFileName
			EndIf
			$aFileList[$iIndex][1] = 0 ;dir
			$iIndex += 1
		EndIf
	WEnd
	FileClose($hFileFind)
EndFunc   ;==>_GetFileListRec_Process



; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __ArrayDualPivotSortByFolder
; Description ...: Helper function for sorting 1D arrays
; Syntax.........: __ArrayDualPivotSort ( ByRef $aArray, $iPivot_Left, $iPivot_Right [, $bLeftMost = True ] )
; Parameters ....: $aArray  - Array to sort
;                  $iPivot_Left  - Index of the array to start sorting at
;                  $iPivot_Right - Index of the array to stop sorting at
;                  $bLeftMost    - Indicates if this part is the leftmost in the range
; Return values .: None
; Author ........: Erik Pilsits
; Modified.......: Melba23
; Modified.......: kurtykurtyboy - use custom compare function
; Remarks .......: For Internal Use Only
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __ArrayDualPivotSortByFolder(ByRef $aArray, $iPivot_Left, $iPivot_Right, $bLeftMost = True)
	If $iPivot_Left > $iPivot_Right Then Return
	Local $iLength = $iPivot_Right - $iPivot_Left + 1
	Local $i, $j, $k, $iAi, $iAk, $iA1, $iA2, $iLast
	If $iLength < 45 Then ; Use insertion sort for small arrays - value chosen empirically
		If $bLeftMost Then
			$i = $iPivot_Left
			While $i < $iPivot_Right
				$j = $i
				$iAi = $aArray[$i + 1]
				While __ArrayDualPivotSortByFolder_Compare($iAi, $aArray[$j])
					$aArray[$j + 1] = $aArray[$j]
					$j -= 1
					If $j + 1 = $iPivot_Left Then ExitLoop
				WEnd
				$aArray[$j + 1] = $iAi
				$i += 1
			WEnd
		Else
			While 1
				If $iPivot_Left >= $iPivot_Right Then Return 1
				$iPivot_Left += 1
				If __ArrayDualPivotSortByFolder_Compare($aArray[$iPivot_Left], $aArray[$iPivot_Left - 1]) Then ExitLoop
			WEnd
			While 1
				$k = $iPivot_Left
				$iPivot_Left += 1
				If $iPivot_Left > $iPivot_Right Then ExitLoop
				$iA1 = $aArray[$k]
				$iA2 = $aArray[$iPivot_Left]
				If __ArrayDualPivotSortByFolder_Compare($iA1, $iA2) Then
					$iA2 = $iA1
					$iA1 = $aArray[$iPivot_Left]
				EndIf
				$k -= 1
				While __ArrayDualPivotSortByFolder_Compare($iA1, $aArray[$k])
					$aArray[$k + 2] = $aArray[$k]
					$k -= 1
				WEnd
				$aArray[$k + 2] = $iA1
				While __ArrayDualPivotSortByFolder_Compare($iA1, $aArray[$k])
					$aArray[$k + 1] = $aArray[$k]
					$k -= 1
				WEnd
				$aArray[$k + 1] = $iA2
				$iPivot_Left += 1
			WEnd
			$iLast = $aArray[$iPivot_Right]
			$iPivot_Right -= 1
			While __ArrayDualPivotSortByFolder_Compare($iLast, $aArray[$iPivot_Right])
				$aArray[$iPivot_Right + 1] = $aArray[$iPivot_Right]
				$iPivot_Right -= 1
			WEnd
			$aArray[$iPivot_Right + 1] = $iLast
		EndIf
		Return 1
	EndIf

	Local $iSeventh = BitShift($iLength, 3) + BitShift($iLength, 6) + 1
	Local $iE1, $iE2, $iE3, $iE4, $iE5, $t
	$iE3 = Ceiling(($iPivot_Left + $iPivot_Right) / 2)
	$iE2 = $iE3 - $iSeventh
	$iE1 = $iE2 - $iSeventh
	$iE4 = $iE3 + $iSeventh
	$iE5 = $iE4 + $iSeventh
	If __ArrayDualPivotSortByFolder_Compare($aArray[$iE2], $aArray[$iE1]) Then
		$t = $aArray[$iE2]
		$aArray[$iE2] = $aArray[$iE1]
		$aArray[$iE1] = $t
	EndIf
	If __ArrayDualPivotSortByFolder_Compare($aArray[$iE3], $aArray[$iE2]) Then
		$t = $aArray[$iE3]
		$aArray[$iE3] = $aArray[$iE2]
		$aArray[$iE2] = $t
		If __ArrayDualPivotSortByFolder_Compare($t, $aArray[$iE1]) Then
			$aArray[$iE2] = $aArray[$iE1]
			$aArray[$iE1] = $t
		EndIf
	EndIf
	If __ArrayDualPivotSortByFolder_Compare($aArray[$iE4], $aArray[$iE3]) Then
		$t = $aArray[$iE4]
		$aArray[$iE4] = $aArray[$iE3]
		$aArray[$iE3] = $t
		If __ArrayDualPivotSortByFolder_Compare($t, $aArray[$iE2]) Then
			$aArray[$iE3] = $aArray[$iE2]
			$aArray[$iE2] = $t
			If __ArrayDualPivotSortByFolder_Compare($t, $aArray[$iE1]) Then
				$aArray[$iE2] = $aArray[$iE1]
				$aArray[$iE1] = $t
			EndIf
		EndIf
	EndIf
	If __ArrayDualPivotSortByFolder_Compare($aArray[$iE5], $aArray[$iE4]) Then
		$t = $aArray[$iE5]
		$aArray[$iE5] = $aArray[$iE4]
		$aArray[$iE4] = $t
		If __ArrayDualPivotSortByFolder_Compare($t, $aArray[$iE3]) Then
			$aArray[$iE4] = $aArray[$iE3]
			$aArray[$iE3] = $t
			If __ArrayDualPivotSortByFolder_Compare($t, $aArray[$iE2]) Then
				$aArray[$iE3] = $aArray[$iE2]
				$aArray[$iE2] = $t
				If __ArrayDualPivotSortByFolder_Compare($t, $aArray[$iE1]) Then
					$aArray[$iE2] = $aArray[$iE1]
					$aArray[$iE1] = $t
				EndIf
			EndIf
		EndIf
	EndIf
	Local $iLess = $iPivot_Left
	Local $iGreater = $iPivot_Right
	If (($aArray[$iE1] <> $aArray[$iE2]) And ($aArray[$iE2] <> $aArray[$iE3]) And ($aArray[$iE3] <> $aArray[$iE4]) And ($aArray[$iE4] <> $aArray[$iE5])) Then
		Local $iPivot_1 = $aArray[$iE2]
		Local $iPivot_2 = $aArray[$iE4]
		$aArray[$iE2] = $aArray[$iPivot_Left]
		$aArray[$iE4] = $aArray[$iPivot_Right]
		Do
			$iLess += 1
		Until __ArrayDualPivotSortByFolder_Compare($aArray[$iLess], $iPivot_1, True, True)
		Do
			$iGreater -= 1
		Until __ArrayDualPivotSortByFolder_Compare($aArray[$iGreater], $iPivot_2, False, True)
		$k = $iLess
		While $k <= $iGreater
			$iAk = $aArray[$k]
			If __ArrayDualPivotSortByFolder_Compare($iAk, $iPivot_1) Then
				$aArray[$k] = $aArray[$iLess]
				$aArray[$iLess] = $iAk
				$iLess += 1
			ElseIf __ArrayDualPivotSortByFolder_Compare($iAk, $iPivot_2, True) Then
				While __ArrayDualPivotSortByFolder_Compare($aArray[$iGreater], $iPivot_2, True)
					$iGreater -= 1
					If $iGreater + 1 = $k Then ExitLoop 2
				WEnd
				If __ArrayDualPivotSortByFolder_Compare($aArray[$iGreater], $iPivot_1) Then
					$aArray[$k] = $aArray[$iLess]
					$aArray[$iLess] = $aArray[$iGreater]
					$iLess += 1
				Else
					$aArray[$k] = $aArray[$iGreater]
				EndIf
				$aArray[$iGreater] = $iAk
				$iGreater -= 1
			EndIf
			$k += 1
		WEnd
		$aArray[$iPivot_Left] = $aArray[$iLess - 1]
		$aArray[$iLess - 1] = $iPivot_1
		$aArray[$iPivot_Right] = $aArray[$iGreater + 1]
		$aArray[$iGreater + 1] = $iPivot_2
		__ArrayDualPivotSortByFolder($aArray, $iPivot_Left, $iLess - 2, True)
		__ArrayDualPivotSortByFolder($aArray, $iGreater + 2, $iPivot_Right, False)
		If ($iLess < $iE1) And ($iE5 < $iGreater) Then
			While $aArray[$iLess] = $iPivot_1
				$iLess += 1
			WEnd
			While $aArray[$iGreater] = $iPivot_2
				$iGreater -= 1
			WEnd
			$k = $iLess
			While $k <= $iGreater
				$iAk = $aArray[$k]
				If $iAk = $iPivot_1 Then
					$aArray[$k] = $aArray[$iLess]
					$aArray[$iLess] = $iAk
					$iLess += 1
				ElseIf $iAk = $iPivot_2 Then
					While $aArray[$iGreater] = $iPivot_2
						$iGreater -= 1
						If $iGreater + 1 = $k Then ExitLoop 2
					WEnd
					If $aArray[$iGreater] = $iPivot_1 Then
						$aArray[$k] = $aArray[$iLess]
						$aArray[$iLess] = $iPivot_1
						$iLess += 1
					Else
						$aArray[$k] = $aArray[$iGreater]
					EndIf
					$aArray[$iGreater] = $iAk
					$iGreater -= 1
				EndIf
				$k += 1
			WEnd
		EndIf
		__ArrayDualPivotSortByFolder($aArray, $iLess, $iGreater, False)
	Else
		Local $iPivot = $aArray[$iE3]
		$k = $iLess
		While $k <= $iGreater
			If $aArray[$k] = $iPivot Then
				$k += 1
				ContinueLoop
			EndIf
			$iAk = $aArray[$k]
			If __ArrayDualPivotSortByFolder_Compare($iAk, $iPivot) Then
				$aArray[$k] = $aArray[$iLess]
				$aArray[$iLess] = $iAk
				$iLess += 1
			Else
				While __ArrayDualPivotSortByFolder_Compare($aArray[$iGreater], $iPivot, True)
					$iGreater -= 1
				WEnd
				If __ArrayDualPivotSortByFolder_Compare($aArray[$iGreater], $iPivot) Then
					$aArray[$k] = $aArray[$iLess]
					$aArray[$iLess] = $aArray[$iGreater]
					$iLess += 1
				Else
					$aArray[$k] = $iPivot
				EndIf
				$aArray[$iGreater] = $iAk
				$iGreater -= 1
			EndIf
			$k += 1
		WEnd
		__ArrayDualPivotSortByFolder($aArray, $iPivot_Left, $iLess - 1, True)
		__ArrayDualPivotSortByFolder($aArray, $iGreater + 1, $iPivot_Right, False)
	EndIf
EndFunc   ;==>__ArrayDualPivotSortByFolder

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __ArrayDualPivotSortByFolder_Compare
; Description ...: Helper function for comparing 2 file paths, while keeping the folder hierarchy together
; Syntax.........: __ArrayDualPivotSortByFolder_Compare ( $sLeft, $sRight [, $bGreater = False [, $bIncludeEqual = False]] )
; Parameters ....: $aArray	- Array to sort
;                  $sLeft	- 1st string to compare
;                  $sRight	- 2nd string to compare
;                  $bGreater		- Use > compare:  check if $sLeft > $sRight.  Default(0) is $sLeft < $sRight
;				   $bIncludeEqual	- Include "=" in comparison: check if $sLeft = $sRight. Default(0) is don't include
; Return values .: Compare evaluates to true:	1
;				   Compare evaluates to false:	0
; Author ........: kurtykurtyboy
; Remarks .......: For Internal Use Only
; ===============================================================================================================================
Func __ArrayDualPivotSortByFolder_Compare($sLeft, $sRight, $bGreater = False, $bIncludeEqual = False)
	Local $aLevelsLeft = StringSplit($sLeft, "\")
	Local $aLevelsRight = StringSplit($sRight, "\")

	Local $iEnd = ($aLevelsLeft[0] < $aLevelsRight[0]) ? $aLevelsLeft[0] : $aLevelsRight[0]
	;every path starts with "\", so skip element [1]
	For $i = 2 To $iEnd
		If $aLevelsLeft[0] = $i And $aLevelsRight[0] = $i Then
			;both sides have no more levels
			If $aLevelsLeft[$i] < $aLevelsRight[$i] Then
				Return Not $bGreater
			Else
				If $bIncludeEqual And $aLevelsLeft[$i] = $aLevelsRight[$i] Then
					Return 1
				Else
					Return $bGreater
				EndIf
			EndIf
		ElseIf $aLevelsLeft[0] = $i And $aLevelsRight[0] > $i Then
			;left side no more levels, right side has more
			If $aLevelsLeft[$i] > $aLevelsRight[$i] Then
				Return $bGreater
			Else
				If $bIncludeEqual And $aLevelsLeft[$i] = $aLevelsRight[$i] Then
					Return 1
				Else
					Return Not $bGreater
				EndIf
			EndIf
		ElseIf $aLevelsLeft[0] > $i And $aLevelsRight[0] = $i Then
			;left side has more levels, right side does not
			If $aLevelsLeft[$i] < $aLevelsRight[$i] Then
				Return Not $bGreater
			Else
				If $bIncludeEqual And $aLevelsLeft[$i] = $aLevelsRight[$i] Then
					Return 1
				Else
					Return $bGreater
				EndIf
			EndIf
		Else
			;both have more levels to check
			If $aLevelsLeft[$i] < $aLevelsRight[$i] Then
				Return Not $bGreater
			ElseIf $aLevelsLeft[$i] > $aLevelsRight[$i] Then
				Return $bGreater
			Else
				If $bIncludeEqual And $aLevelsLeft[$i] = $aLevelsRight[$i] Then
					Return 1
				Else
					ContinueLoop
				EndIf
			EndIf
		EndIf
	Next
EndFunc   ;==>__ArrayDualPivotSortByFolder_Compare
