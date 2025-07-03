Opt("TrayAutoPause", 0)
Opt("TrayIconHide", 1)
Global Const $OPT_MATCHSTART = 1
Global Const $UBOUND_DIMENSIONS = 0
Global Const $UBOUND_ROWS = 1
Global Const $UBOUND_COLUMNS = 2
Global Const $STR_NOCASESENSEBASIC = 2
Global Const $STR_STRIPLEADING = 1
Global Const $STR_STRIPTRAILING = 2
Func _ArrayConcatenate(ByRef $aArrayTarget, Const ByRef $aArraySource, $iStart = 0)
If $iStart = Default Then $iStart = 0
If Not IsArray($aArrayTarget) Then Return SetError(1, 0, -1)
If Not IsArray($aArraySource) Then Return SetError(2, 0, -1)
Local $iDim_Total_Tgt = UBound($aArrayTarget, $UBOUND_DIMENSIONS)
Local $iDim_Total_Src = UBound($aArraySource, $UBOUND_DIMENSIONS)
Local $iDim_1_Tgt = UBound($aArrayTarget, $UBOUND_ROWS)
Local $iDim_1_Src = UBound($aArraySource, $UBOUND_ROWS)
If $iStart < 0 Or $iStart > $iDim_1_Src - 1 Then Return SetError(6, 0, -1)
Switch $iDim_Total_Tgt
Case 1
If $iDim_Total_Src <> 1 Then Return SetError(4, 0, -1)
ReDim $aArrayTarget[$iDim_1_Tgt + $iDim_1_Src - $iStart]
For $i = $iStart To $iDim_1_Src - 1
$aArrayTarget[$iDim_1_Tgt + $i - $iStart] = $aArraySource[$i]
Next
Case 2
If $iDim_Total_Src <> 2 Then Return SetError(4, 0, -1)
Local $iDim_2_Tgt = UBound($aArrayTarget, $UBOUND_COLUMNS)
If UBound($aArraySource, $UBOUND_COLUMNS) <> $iDim_2_Tgt Then Return SetError(5, 0, -1)
ReDim $aArrayTarget[$iDim_1_Tgt + $iDim_1_Src - $iStart][$iDim_2_Tgt]
For $i = $iStart To $iDim_1_Src - 1
For $j = 0 To $iDim_2_Tgt - 1
$aArrayTarget[$iDim_1_Tgt + $i - $iStart][$j] = $aArraySource[$i][$j]
Next
Next
Case Else
Return SetError(3, 0, -1)
EndSwitch
Return UBound($aArrayTarget, $UBOUND_ROWS)
EndFunc
Func __ArrayDualPivotSort(ByRef $aArray, $iPivot_Left, $iPivot_Right, $bLeftMost = True)
If $iPivot_Left > $iPivot_Right Then Return
Local $iLength = $iPivot_Right - $iPivot_Left + 1
Local $i, $j, $k, $iAi, $iAk, $iA1, $iA2, $iLast
If $iLength < 45 Then
If $bLeftMost Then
$i = $iPivot_Left
While $i < $iPivot_Right
$j = $i
$iAi = $aArray[$i + 1]
While $iAi < $aArray[$j]
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
If $aArray[$iPivot_Left] < $aArray[$iPivot_Left - 1] Then ExitLoop
WEnd
While 1
$k = $iPivot_Left
$iPivot_Left += 1
If $iPivot_Left > $iPivot_Right Then ExitLoop
$iA1 = $aArray[$k]
$iA2 = $aArray[$iPivot_Left]
If $iA1 < $iA2 Then
$iA2 = $iA1
$iA1 = $aArray[$iPivot_Left]
EndIf
$k -= 1
While $iA1 < $aArray[$k]
$aArray[$k + 2] = $aArray[$k]
$k -= 1
WEnd
$aArray[$k + 2] = $iA1
While $iA2 < $aArray[$k]
$aArray[$k + 1] = $aArray[$k]
$k -= 1
WEnd
$aArray[$k + 1] = $iA2
$iPivot_Left += 1
WEnd
$iLast = $aArray[$iPivot_Right]
$iPivot_Right -= 1
While $iLast < $aArray[$iPivot_Right]
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
If $aArray[$iE2] < $aArray[$iE1] Then
$t = $aArray[$iE2]
$aArray[$iE2] = $aArray[$iE1]
$aArray[$iE1] = $t
EndIf
If $aArray[$iE3] < $aArray[$iE2] Then
$t = $aArray[$iE3]
$aArray[$iE3] = $aArray[$iE2]
$aArray[$iE2] = $t
If $t < $aArray[$iE1] Then
$aArray[$iE2] = $aArray[$iE1]
$aArray[$iE1] = $t
EndIf
EndIf
If $aArray[$iE4] < $aArray[$iE3] Then
$t = $aArray[$iE4]
$aArray[$iE4] = $aArray[$iE3]
$aArray[$iE3] = $t
If $t < $aArray[$iE2] Then
$aArray[$iE3] = $aArray[$iE2]
$aArray[$iE2] = $t
If $t < $aArray[$iE1] Then
$aArray[$iE2] = $aArray[$iE1]
$aArray[$iE1] = $t
EndIf
EndIf
EndIf
If $aArray[$iE5] < $aArray[$iE4] Then
$t = $aArray[$iE5]
$aArray[$iE5] = $aArray[$iE4]
$aArray[$iE4] = $t
If $t < $aArray[$iE3] Then
$aArray[$iE4] = $aArray[$iE3]
$aArray[$iE3] = $t
If $t < $aArray[$iE2] Then
$aArray[$iE3] = $aArray[$iE2]
$aArray[$iE2] = $t
If $t < $aArray[$iE1] Then
$aArray[$iE2] = $aArray[$iE1]
$aArray[$iE1] = $t
EndIf
EndIf
EndIf
EndIf
Local $iLess = $iPivot_Left
Local $iGreater = $iPivot_Right
If(($aArray[$iE1] <> $aArray[$iE2]) And($aArray[$iE2] <> $aArray[$iE3]) And($aArray[$iE3] <> $aArray[$iE4]) And($aArray[$iE4] <> $aArray[$iE5])) Then
Local $iPivot_1 = $aArray[$iE2]
Local $iPivot_2 = $aArray[$iE4]
$aArray[$iE2] = $aArray[$iPivot_Left]
$aArray[$iE4] = $aArray[$iPivot_Right]
Do
$iLess += 1
Until $aArray[$iLess] >= $iPivot_1
Do
$iGreater -= 1
Until $aArray[$iGreater] <= $iPivot_2
$k = $iLess
While $k <= $iGreater
$iAk = $aArray[$k]
If $iAk < $iPivot_1 Then
$aArray[$k] = $aArray[$iLess]
$aArray[$iLess] = $iAk
$iLess += 1
ElseIf $iAk > $iPivot_2 Then
While $aArray[$iGreater] > $iPivot_2
$iGreater -= 1
If $iGreater + 1 = $k Then ExitLoop 2
WEnd
If $aArray[$iGreater] < $iPivot_1 Then
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
__ArrayDualPivotSort($aArray, $iPivot_Left, $iLess - 2, True)
__ArrayDualPivotSort($aArray, $iGreater + 2, $iPivot_Right, False)
If($iLess < $iE1) And($iE5 < $iGreater) Then
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
__ArrayDualPivotSort($aArray, $iLess, $iGreater, False)
Else
Local $iPivot = $aArray[$iE3]
$k = $iLess
While $k <= $iGreater
If $aArray[$k] = $iPivot Then
$k += 1
ContinueLoop
EndIf
$iAk = $aArray[$k]
If $iAk < $iPivot Then
$aArray[$k] = $aArray[$iLess]
$aArray[$iLess] = $iAk
$iLess += 1
Else
While $aArray[$iGreater] > $iPivot
$iGreater -= 1
WEnd
If $aArray[$iGreater] < $iPivot Then
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
__ArrayDualPivotSort($aArray, $iPivot_Left, $iLess - 1, True)
__ArrayDualPivotSort($aArray, $iGreater + 1, $iPivot_Right, False)
EndIf
EndFunc
Func _SendMessage($hWnd, $iMsg, $wParam = 0, $lParam = 0, $iReturn = 0, $wParamType = "wparam", $lParamType = "lparam", $sReturnType = "lresult")
Local $aCall = DllCall("user32.dll", $sReturnType, "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, $wParamType, $wParam, $lParamType, $lParam)
If @error Then Return SetError(@error, @extended, "")
If $iReturn >= 0 And $iReturn <= 4 Then Return $aCall[$iReturn]
Return $aCall
EndFunc
Global Const $FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x00000100
Global Const $FORMAT_MESSAGE_IGNORE_INSERTS = 0x00000200
Global Const $FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000
Func _WinAPI_FormatMessage($iFlags, $pSource, $iMessageID, $iLanguageID, ByRef $pBuffer, $iSize, $vArguments)
Local $sBufferType = "struct*"
If IsString($pBuffer) Then $sBufferType = "wstr"
Local $aCall = DllCall("kernel32.dll", "dword", "FormatMessageW", "dword", $iFlags, "struct*", $pSource, "dword", $iMessageID, "dword", $iLanguageID, $sBufferType, $pBuffer, "dword", $iSize, "ptr", $vArguments)
If @error Then Return SetError(@error, @extended, 0)
If Not $aCall[0] Then Return SetError(10, _WinAPI_GetLastError(), 0)
If $sBufferType = "wstr" Then $pBuffer = $aCall[5]
Return $aCall[0]
EndFunc
Func _WinAPI_GetLastError(Const $_iCallerError = @error, Const $_iCallerExtended = @extended)
Local $aCall = DllCall("kernel32.dll", "dword", "GetLastError")
Return SetError($_iCallerError, $_iCallerExtended, $aCall[0])
EndFunc
Func _WinAPI_GetLastErrorMessage(Const $_iCallerError = @error, Const $_iCallerExtended = @extended)
Local $iLastError = _WinAPI_GetLastError()
Local $tBufferPtr = DllStructCreate("ptr")
Local $nCount = _WinAPI_FormatMessage(BitOR($FORMAT_MESSAGE_ALLOCATE_BUFFER, $FORMAT_MESSAGE_FROM_SYSTEM, $FORMAT_MESSAGE_IGNORE_INSERTS), 0, $iLastError, 0, $tBufferPtr, 0, 0)
If @error Then Return SetError(-@error, @extended, "")
Local $sText = ""
Local $pBuffer = DllStructGetData($tBufferPtr, 1)
If $pBuffer Then
If $nCount > 0 Then
Local $tBuffer = DllStructCreate("wchar[" &($nCount + 1) & "]", $pBuffer)
$sText = DllStructGetData($tBuffer, 1)
If StringRight($sText, 2) = @CRLF Then $sText = StringTrimRight($sText, 2)
EndIf
DllCall("kernel32.dll", "handle", "LocalFree", "handle", $pBuffer)
EndIf
Return SetError($_iCallerError, $_iCallerExtended, $sText)
EndFunc
Func _WinAPI_SetLastError($iErrorCode, Const $_iCallerError = @error, Const $_iCallerExtended = @extended)
DllCall("kernel32.dll", "none", "SetLastError", "dword", $iErrorCode)
Return SetError($_iCallerError, $_iCallerExtended, Null)
EndFunc
Global Const $__g_sReportWindowText_Debug = "Debug Window hidden text"
Global $__g_sReportTitle_Debug = "AutoIt Debug Report"
Global $__g_iReportType_Debug = 0
Global $__g_bReportWindowWaitClose_Debug = True, $__g_bReportWindowClosed_Debug = True
Global $__g_hReportEdit_Debug = 0
Global $__g_sReportCallBack_Debug
Global $__g_bReportTimeStamp_Debug = False
Func _DebugReport($sData, $bLastError = False, $bExit = False, Const $_iCallerError = @error, $_iCallerExtended = @extended)
If $__g_iReportType_Debug <= 0 Or $__g_iReportType_Debug > 6 Then Return SetError($_iCallerError, $_iCallerExtended, 0)
Local $iLastError = _WinAPI_GetLastError()
__Debug_ReportWrite($sData, $bLastError, $iLastError)
If $bExit Then Exit
_WinAPI_SetLastError($iLastError)
If $bLastError Then $_iCallerExtended = $iLastError
Return SetError($_iCallerError, $_iCallerExtended, 1)
EndFunc
Func __Debug_ReportWindowCreate()
Local $nOld = Opt("WinDetectHiddenText", $OPT_MATCHSTART)
Local $bExists = WinExists($__g_sReportTitle_Debug, $__g_sReportWindowText_Debug)
If $bExists Then
If $__g_hReportEdit_Debug = 0 Then
$__g_hReportEdit_Debug = ControlGetHandle($__g_sReportTitle_Debug, $__g_sReportWindowText_Debug, "Edit1")
$__g_bReportWindowWaitClose_Debug = False
EndIf
EndIf
Opt("WinDetectHiddenText", $nOld)
$__g_bReportWindowClosed_Debug = False
If Not $__g_bReportWindowWaitClose_Debug Then Return 0
Local Const $WS_OVERLAPPEDWINDOW = 0x00CF0000
Local Const $WS_HSCROLL = 0x00100000
Local Const $WS_VSCROLL = 0x00200000
Local Const $ES_READONLY = 2048
Local Const $EM_LIMITTEXT = 0xC5
Local Const $GUI_HIDE = 32
Local $w = 580, $h = 380
GUICreate($__g_sReportTitle_Debug, $w, $h, -1, -1, $WS_OVERLAPPEDWINDOW)
Local $idLabelHidden = GUICtrlCreateLabel($__g_sReportWindowText_Debug, 0, 0, 1, 1)
GUICtrlSetState($idLabelHidden, $GUI_HIDE)
Local $idEdit = GUICtrlCreateEdit("", 4, 4, $w - 8, $h - 8, BitOR($WS_HSCROLL, $WS_VSCROLL, $ES_READONLY))
$__g_hReportEdit_Debug = GUICtrlGetHandle($idEdit)
GUICtrlSetBkColor($idEdit, 0xFFFFFF)
GUICtrlSendMsg($idEdit, $EM_LIMITTEXT, 0, 0)
GUISetState()
$__g_bReportWindowWaitClose_Debug = True
Return 1
EndFunc
#Au3Stripper_Ignore_Funcs=__Debug_ReportWindowWrite
Func __Debug_ReportWindowWrite($sData)
If $__g_bReportWindowClosed_Debug Then __Debug_ReportWindowCreate()
Local Const $WM_GETTEXTLENGTH = 0x000E
Local Const $EM_SETSEL = 0xB1
Local Const $EM_REPLACESEL = 0xC2
Local $nLen = _SendMessage($__g_hReportEdit_Debug, $WM_GETTEXTLENGTH, 0, 0, 0, "int", "int")
_SendMessage($__g_hReportEdit_Debug, $EM_SETSEL, $nLen, $nLen, 0, "int", "int")
_SendMessage($__g_hReportEdit_Debug, $EM_REPLACESEL, True, $sData, 0, "int", "wstr")
EndFunc
Func __Debug_ReportNotepadCreate()
Local $bExists = WinExists($__g_sReportTitle_Debug)
If $bExists Then
If $__g_hReportEdit_Debug = 0 Then
$__g_hReportEdit_Debug = WinGetHandle($__g_sReportTitle_Debug)
Return 0
EndIf
EndIf
Local $pNotepad = Run("Notepad.exe")
$__g_hReportEdit_Debug = WinWait("[CLASS:Notepad]")
If $pNotepad <> WinGetProcess($__g_hReportEdit_Debug) Then
Return SetError(3, 0, 0)
EndIf
WinActivate($__g_hReportEdit_Debug)
WinSetTitle($__g_hReportEdit_Debug, "", String($__g_sReportTitle_Debug))
Return 1
EndFunc
#Au3Stripper_Ignore_Funcs=__Debug_ReportNotepadWrite
Func __Debug_ReportNotepadWrite($sData)
If $__g_hReportEdit_Debug = 0 Then __Debug_ReportNotepadCreate()
ControlCommand($__g_hReportEdit_Debug, "", "Edit1", "EditPaste", String($sData))
EndFunc
Func __Debug_ReportWrite($sData, $bLastError = False, $iLastError = 0)
Local $sError = ""
If $__g_bReportTimeStamp_Debug And($sData <> "") Then $sData = @YEAR & "/" & @MON & "/" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & " " & $sData
If $bLastError Then
$sError = " LastError = " & $iLastError & " : (" & _WinAPI_GetLastErrorMessage() & ")" & @CRLF
EndIf
$sData &= $sError
$sData = StringReplace($sData, "'", "''")
Local Static $sERROR_CODE = ">Error code:"
If StringInStr($sData, $sERROR_CODE) Then
$sData = StringReplace($sData, $sERROR_CODE, @TAB & $sERROR_CODE)
If(StringInStr($sData, $sERROR_CODE & " 0") = 0) Then
$sData = StringReplace($sData, $sERROR_CODE, $sERROR_CODE & @TAB & @TAB & @TAB & @TAB)
EndIf
EndIf
Execute($__g_sReportCallBack_Debug & "'" & $sData & "')")
Return
EndFunc
Global Const $FLTA_FILESFOLDERS = 0
Global Const $FLTAR_FILESFOLDERS = 0
Global Const $FLTAR_NOHIDDEN = 4
Global Const $FLTAR_NOSYSTEM = 8
Global Const $FLTAR_NOLINK = 16
Global Const $FLTAR_NORECUR = 0
Global Const $FLTAR_NOSORT = 0
Global Const $FLTAR_RELPATH = 1
Func _FileListToArray($sFilePath, $sFilter = "*", $iFlag = $FLTA_FILESFOLDERS, $bReturnPath = False)
Local $sDelimiter = "|", $sFileList = "", $sFileName = "", $sFullPath = ""
$sFilePath = StringRegExpReplace($sFilePath, "[\\/]+$", "") & "\"
If $iFlag = Default Then $iFlag = $FLTA_FILESFOLDERS
If $bReturnPath Then $sFullPath = $sFilePath
If $sFilter = Default Then $sFilter = "*"
If Not FileExists($sFilePath) Then Return SetError(1, 0, 0)
If StringRegExp($sFilter, "[\\/:><\|]|(?s)^\s*$") Then Return SetError(2, 0, 0)
If Not($iFlag = 0 Or $iFlag = 1 Or $iFlag = 2) Then Return SetError(3, 0, 0)
Local $hSearch = FileFindFirstFile($sFilePath & $sFilter)
If @error Then Return SetError(4, 0, 0)
While 1
$sFileName = FileFindNextFile($hSearch)
If @error Then ExitLoop
If($iFlag + @extended = 2) Then ContinueLoop
$sFileList &= $sDelimiter & $sFullPath & $sFileName
WEnd
FileClose($hSearch)
If $sFileList = "" Then Return SetError(4, 0, 0)
Return StringSplit(StringTrimLeft($sFileList, 1), $sDelimiter)
EndFunc
Func _FileListToArrayRec($sFilePath, $sMask = "*", $iReturn = $FLTAR_FILESFOLDERS, $iRecur = $FLTAR_NORECUR, $iSort = $FLTAR_NOSORT, $iReturnPath = $FLTAR_RELPATH)
If Not FileExists($sFilePath) Then Return SetError(1, 1, "")
If $sMask = Default Then $sMask = "*"
If $iReturn = Default Then $iReturn = $FLTAR_FILESFOLDERS
If $iRecur = Default Then $iRecur = $FLTAR_NORECUR
If $iSort = Default Then $iSort = $FLTAR_NOSORT
If $iReturnPath = Default Then $iReturnPath = $FLTAR_RELPATH
If $iRecur > 1 Or Not IsInt($iRecur) Then Return SetError(1, 6, "")
Local $bLongPath = False
If StringLeft($sFilePath, 4) == "\\?\" Then
$bLongPath = True
EndIf
Local $sFolderSlash = ""
If StringRight($sFilePath, 1) = "\" Then
$sFolderSlash = "\"
Else
$sFilePath = $sFilePath & "\"
EndIf
Local $asFolderSearchList[100] = [1]
$asFolderSearchList[1] = $sFilePath
Local $iHide_HS = 0, $sHide_HS = ""
If BitAND($iReturn, $FLTAR_NOHIDDEN) Then
$iHide_HS += 2
$sHide_HS &= "H"
$iReturn -= $FLTAR_NOHIDDEN
EndIf
If BitAND($iReturn, $FLTAR_NOSYSTEM) Then
$iHide_HS += 4
$sHide_HS &= "S"
$iReturn -= $FLTAR_NOSYSTEM
EndIf
Local $iHide_Link = 0
If BitAND($iReturn, $FLTAR_NOLINK) Then
$iHide_Link = 0x400
$iReturn -= $FLTAR_NOLINK
EndIf
Local $iMaxLevel = 0
If $iRecur < 0 Then
StringReplace($sFilePath, "\", "", 0, $STR_NOCASESENSEBASIC)
$iMaxLevel = @extended - $iRecur
EndIf
Local $sExclude_List = "", $sExclude_List_Folder = "", $sInclude_List = "*"
Local $aMaskSplit = StringSplit($sMask, "|")
Switch $aMaskSplit[0]
Case 3
$sExclude_List_Folder = $aMaskSplit[3]
ContinueCase
Case 2
$sExclude_List = $aMaskSplit[2]
ContinueCase
Case 1
$sInclude_List = $aMaskSplit[1]
EndSwitch
Local $sInclude_File_Mask = ".+"
If $sInclude_List <> "*" Then
If Not __FLTAR_ListToMask($sInclude_File_Mask, $sInclude_List) Then Return SetError(1, 2, "")
EndIf
Local $sInclude_Folder_Mask = ".+"
Switch $iReturn
Case 0
Switch $iRecur
Case 0
$sInclude_Folder_Mask = $sInclude_File_Mask
EndSwitch
Case 2
$sInclude_Folder_Mask = $sInclude_File_Mask
EndSwitch
Local $sExclude_File_Mask = ":"
If $sExclude_List <> "" Then
If Not __FLTAR_ListToMask($sExclude_File_Mask, $sExclude_List) Then Return SetError(1, 3, "")
EndIf
Local $sExclude_Folder_Mask = ":"
If $iRecur Then
If $sExclude_List_Folder Then
If Not __FLTAR_ListToMask($sExclude_Folder_Mask, $sExclude_List_Folder) Then Return SetError(1, 4, "")
EndIf
If $iReturn = 2 Then
$sExclude_Folder_Mask = $sExclude_File_Mask
EndIf
Else
$sExclude_Folder_Mask = $sExclude_File_Mask
EndIf
If Not($iReturn = 0 Or $iReturn = 1 Or $iReturn = 2) Then Return SetError(1, 5, "")
If Not($iSort = 0 Or $iSort = 1 Or $iSort = 2) Then Return SetError(1, 7, "")
If Not($iReturnPath = 0 Or $iReturnPath = 1 Or $iReturnPath = 2) Then Return SetError(1, 8, "")
If $iHide_Link Then
Local $tFile_Data = DllStructCreate("struct;align 4;dword FileAttributes;uint64 CreationTime;uint64 LastAccessTime;uint64 LastWriteTime;" & "dword FileSizeHigh;dword FileSizeLow;dword Reserved0;dword Reserved1;wchar FileName[260];wchar AlternateFileName[14];endstruct")
Local $hDLL = DllOpen('kernel32.dll'), $aDLL_Ret
EndIf
Local $asReturnList[100] = [0]
Local $asFileMatchList = $asReturnList, $asRootFileMatchList = $asReturnList, $asFolderMatchList = $asReturnList
Local $bFolder = False, $hSearch = 0, $sCurrentPath = "", $sName = "", $sRetPath = ""
Local $iAttribs = 0, $sAttribs = ''
Local $asFolderFileSectionList[100][2] = [[0, 0]]
While $asFolderSearchList[0] > 0
$sCurrentPath = $asFolderSearchList[$asFolderSearchList[0]]
$asFolderSearchList[0] -= 1
Switch $iReturnPath
Case 1
$sRetPath = StringReplace($sCurrentPath, $sFilePath, "")
Case 2
If $bLongPath Then
$sRetPath = StringTrimLeft($sCurrentPath, 4)
Else
$sRetPath = $sCurrentPath
EndIf
EndSwitch
If $iHide_Link Then
$aDLL_Ret = DllCall($hDLL, 'handle', 'FindFirstFileW', 'wstr', $sCurrentPath & "*", 'struct*', $tFile_Data)
If @error Or Not $aDLL_Ret[0] Then
ContinueLoop
EndIf
$hSearch = $aDLL_Ret[0]
Else
$hSearch = FileFindFirstFile($sCurrentPath & "*")
If $hSearch = -1 Then
ContinueLoop
EndIf
EndIf
If $iReturn = 0 And $iSort And $iReturnPath Then
__FLTAR_AddToList($asFolderFileSectionList, $sRetPath, $asFileMatchList[0] + 1)
EndIf
$sAttribs = ''
While 1
If $iHide_Link Then
$aDLL_Ret = DllCall($hDLL, 'int', 'FindNextFileW', 'handle', $hSearch, 'struct*', $tFile_Data)
If @error Or Not $aDLL_Ret[0] Then
ExitLoop
EndIf
$sName = DllStructGetData($tFile_Data, "FileName")
If $sName = ".." Or $sName = "." Then
ContinueLoop
EndIf
$iAttribs = DllStructGetData($tFile_Data, "FileAttributes")
If $iHide_HS And BitAND($iAttribs, $iHide_HS) Then
ContinueLoop
EndIf
If BitAND($iAttribs, $iHide_Link) Then
ContinueLoop
EndIf
$bFolder = False
If BitAND($iAttribs, 16) Then
$bFolder = True
EndIf
Else
$bFolder = False
$sName = FileFindNextFile($hSearch, 1)
If @error Then
ExitLoop
EndIf
If $sName = ".." Or $sName = "." Then
ContinueLoop
EndIf
$sAttribs = @extended
If StringInStr($sAttribs, "D") Then
$bFolder = True
EndIf
If StringRegExp($sAttribs, "[" & $sHide_HS & "]") Then
ContinueLoop
EndIf
EndIf
If $bFolder Then
Select
Case $iRecur < 0
StringReplace($sCurrentPath, "\", "", 0, $STR_NOCASESENSEBASIC)
If @extended < $iMaxLevel Then
ContinueCase
EndIf
Case $iRecur = 1
If Not StringRegExp($sName, $sExclude_Folder_Mask) Then
__FLTAR_AddToList($asFolderSearchList, $sCurrentPath & $sName & "\")
EndIf
EndSelect
EndIf
If $iSort Then
If $bFolder Then
If StringRegExp($sName, $sInclude_Folder_Mask) And Not StringRegExp($sName, $sExclude_Folder_Mask) Then
__FLTAR_AddToList($asFolderMatchList, $sRetPath & $sName & $sFolderSlash)
EndIf
Else
If StringRegExp($sName, $sInclude_File_Mask) And Not StringRegExp($sName, $sExclude_File_Mask) Then
If $sCurrentPath = $sFilePath Then
__FLTAR_AddToList($asRootFileMatchList, $sRetPath & $sName)
Else
__FLTAR_AddToList($asFileMatchList, $sRetPath & $sName)
EndIf
EndIf
EndIf
Else
If $bFolder Then
If $iReturn <> 1 And StringRegExp($sName, $sInclude_Folder_Mask) And Not StringRegExp($sName, $sExclude_Folder_Mask) Then
__FLTAR_AddToList($asReturnList, $sRetPath & $sName & $sFolderSlash)
EndIf
Else
If $iReturn <> 2 And StringRegExp($sName, $sInclude_File_Mask) And Not StringRegExp($sName, $sExclude_File_Mask) Then
__FLTAR_AddToList($asReturnList, $sRetPath & $sName)
EndIf
EndIf
EndIf
WEnd
If $iHide_Link Then
DllCall($hDLL, 'int', 'FindClose', 'ptr', $hSearch)
Else
FileClose($hSearch)
EndIf
WEnd
If $iHide_Link Then
DllClose($hDLL)
EndIf
If $iSort Then
Switch $iReturn
Case 2
If $asFolderMatchList[0] = 0 Then Return SetError(1, 9, "")
ReDim $asFolderMatchList[$asFolderMatchList[0] + 1]
$asReturnList = $asFolderMatchList
__ArrayDualPivotSort($asReturnList, 1, $asReturnList[0])
Case 1
If $asRootFileMatchList[0] = 0 And $asFileMatchList[0] = 0 Then Return SetError(1, 9, "")
If $iReturnPath = 0 Then
__FLTAR_AddFileLists($asReturnList, $asRootFileMatchList, $asFileMatchList)
__ArrayDualPivotSort($asReturnList, 1, $asReturnList[0])
Else
__FLTAR_AddFileLists($asReturnList, $asRootFileMatchList, $asFileMatchList, 1)
EndIf
Case 0
If $asRootFileMatchList[0] = 0 And $asFolderMatchList[0] = 0 Then Return SetError(1, 9, "")
If $iReturnPath = 0 Then
__FLTAR_AddFileLists($asReturnList, $asRootFileMatchList, $asFileMatchList)
$asReturnList[0] += $asFolderMatchList[0]
ReDim $asFolderMatchList[$asFolderMatchList[0] + 1]
_ArrayConcatenate($asReturnList, $asFolderMatchList, 1)
__ArrayDualPivotSort($asReturnList, 1, $asReturnList[0])
Else
Local $asReturnList[$asFileMatchList[0] + $asRootFileMatchList[0] + $asFolderMatchList[0] + 1]
$asReturnList[0] = $asFileMatchList[0] + $asRootFileMatchList[0] + $asFolderMatchList[0]
__ArrayDualPivotSort($asRootFileMatchList, 1, $asRootFileMatchList[0])
For $i = 1 To $asRootFileMatchList[0]
$asReturnList[$i] = $asRootFileMatchList[$i]
Next
Local $iNextInsertionIndex = $asRootFileMatchList[0] + 1
__ArrayDualPivotSort($asFolderMatchList, 1, $asFolderMatchList[0])
Local $sFolderToFind = ""
For $i = 1 To $asFolderMatchList[0]
$asReturnList[$iNextInsertionIndex] = $asFolderMatchList[$i]
$iNextInsertionIndex += 1
If $sFolderSlash Then
$sFolderToFind = $asFolderMatchList[$i]
Else
$sFolderToFind = $asFolderMatchList[$i] & "\"
EndIf
Local $iFileSectionEndIndex = 0, $iFileSectionStartIndex = 0
For $j = 1 To $asFolderFileSectionList[0][0]
If $sFolderToFind = $asFolderFileSectionList[$j][0] Then
$iFileSectionStartIndex = $asFolderFileSectionList[$j][1]
If $j = $asFolderFileSectionList[0][0] Then
$iFileSectionEndIndex = $asFileMatchList[0]
Else
$iFileSectionEndIndex = $asFolderFileSectionList[$j + 1][1] - 1
EndIf
If $iSort = 1 Then
__ArrayDualPivotSort($asFileMatchList, $iFileSectionStartIndex, $iFileSectionEndIndex)
EndIf
For $k = $iFileSectionStartIndex To $iFileSectionEndIndex
$asReturnList[$iNextInsertionIndex] = $asFileMatchList[$k]
$iNextInsertionIndex += 1
Next
ExitLoop
EndIf
Next
Next
EndIf
EndSwitch
Else
If $asReturnList[0] = 0 Then Return SetError(1, 9, "")
ReDim $asReturnList[$asReturnList[0] + 1]
EndIf
Return $asReturnList
EndFunc
Func __FLTAR_AddFileLists(ByRef $asTarget, $asSource_1, $asSource_2, $iSort = 0)
ReDim $asSource_1[$asSource_1[0] + 1]
If $iSort = 1 Then __ArrayDualPivotSort($asSource_1, 1, $asSource_1[0])
$asTarget = $asSource_1
$asTarget[0] += $asSource_2[0]
ReDim $asSource_2[$asSource_2[0] + 1]
If $iSort = 1 Then __ArrayDualPivotSort($asSource_2, 1, $asSource_2[0])
_ArrayConcatenate($asTarget, $asSource_2, 1)
EndFunc
Func __FLTAR_AddToList(ByRef $aList, $vValue_0, $vValue_1 = -1)
If $vValue_1 = -1 Then
$aList[0] += 1
If UBound($aList) <= $aList[0] Then ReDim $aList[UBound($aList) * 2]
$aList[$aList[0]] = $vValue_0
Else
$aList[0][0] += 1
If UBound($aList) <= $aList[0][0] Then ReDim $aList[UBound($aList) * 2][2]
$aList[$aList[0][0]][0] = $vValue_0
$aList[$aList[0][0]][1] = $vValue_1
EndIf
EndFunc
Func __FLTAR_ListToMask(ByRef $sMask, $sList)
If StringRegExp($sList, "\\|/|:|\<|\>|\|") Then Return 0
$sList = StringReplace(StringStripWS(StringRegExpReplace($sList, "\s*;\s*", ";"), BitOR($STR_STRIPLEADING, $STR_STRIPTRAILING)), ";", "|")
$sList = StringReplace(StringReplace(StringRegExpReplace($sList, "[][$^.{}()+\-]", "\\$0"), "?", "."), "*", ".*?")
$sMask = "(?i)^(" & $sList & ")\z"
Return 1
EndFunc
Global Const $CB_ERR = -1
Global Const $CBS_AUTOHSCROLL = 0x40
Global Const $CBS_DROPDOWN = 0x2
Global Const $CB_ADDSTRING = 0x143
Global Const $CB_FINDSTRING = 0x14C
Global Const $CB_GETCOMBOBOXINFO = 0x164
Global Const $CB_GETCURSEL = 0x147
Global Const $CB_GETLBTEXT = 0x148
Global Const $CB_GETLBTEXTLEN = 0x149
Global Const $CB_RESETCONTENT = 0x14B
Global Const $CB_SETCURSEL = 0x14E
Global Const $CB_SETEDITSEL = 0x142
Global Const $CBN_EDITCHANGE = 5
Global Const $tagGUID = "struct;ulong Data1;ushort Data2;ushort Data3;byte Data4[8];endstruct"
Global $__g_vEnum, $__g_vExt = 0
Global Const $IMAGE_ICON = 1
Global Const $LR_LOADFROMFILE = 0x0010
Global Const $LR_DEFAULTSIZE = 0x0040
Func _WinAPI_LoadImage($hInstance, $sImage, $iType, $iXDesired, $iYDesired, $iLoad)
Local $aCall, $sImageType = "int"
If IsString($sImage) Then $sImageType = "wstr"
$aCall = DllCall("user32.dll", "handle", "LoadImageW", "handle", $hInstance, $sImageType, $sImage, "uint", $iType, "int", $iXDesired, "int", $iYDesired, "uint", $iLoad)
If @error Then Return SetError(@error, @extended, 0)
Return $aCall[0]
EndFunc
Func __EnumWindowsProc($hWnd, $bVisible)
Local $aCall
If $bVisible Then
$aCall = DllCall("user32.dll", "bool", "IsWindowVisible", "hwnd", $hWnd)
If Not $aCall[0] Then
Return 1
EndIf
EndIf
__Inc($__g_vEnum)
$__g_vEnum[$__g_vEnum[0][0]][0] = $hWnd
$aCall = DllCall("user32.dll", "int", "GetClassNameW", "hwnd", $hWnd, "wstr", "", "int", 4096)
$__g_vEnum[$__g_vEnum[0][0]][1] = $aCall[2]
Return 1
EndFunc
Func __Inc(ByRef $aData, $iIncrement = 100)
Select
Case UBound($aData, $UBOUND_COLUMNS)
If $iIncrement < 0 Then
ReDim $aData[$aData[0][0] + 1][UBound($aData, $UBOUND_COLUMNS)]
Else
$aData[0][0] += 1
If $aData[0][0] > UBound($aData) - 1 Then
ReDim $aData[$aData[0][0] + $iIncrement][UBound($aData, $UBOUND_COLUMNS)]
EndIf
EndIf
Case UBound($aData, $UBOUND_ROWS)
If $iIncrement < 0 Then
ReDim $aData[$aData[0] + 1]
Else
$aData[0] += 1
If $aData[0] > UBound($aData) - 1 Then
ReDim $aData[$aData[0] + $iIncrement]
EndIf
EndIf
Case Else
Return 0
EndSelect
Return 1
EndFunc
Func _WinAPI_GUIDFromString($sGUID)
Local $tGUID = DllStructCreate($tagGUID)
If Not _WinAPI_GUIDFromStringEx($sGUID, $tGUID) Then Return SetError(@error, @extended, 0)
Return $tGUID
EndFunc
Func _WinAPI_GUIDFromStringEx($sGUID, $tGUID)
Local $aCall = DllCall("ole32.dll", "long", "CLSIDFromString", "wstr", $sGUID, "struct*", $tGUID)
If @error Then Return SetError(@error, @extended, False)
If $aCall[0] Then Return SetError(10, $aCall[0], False)
Return True
EndFunc
Func _WinAPI_MakeLong($iLo, $iHi)
Return BitOR(BitShift($iHi, -16), BitAND($iLo, 0xFFFF))
EndFunc
Func _WinAPI_DeleteObject($hObject)
Local $aCall = DllCall("gdi32.dll", "bool", "DeleteObject", "handle", $hObject)
If @error Then Return SetError(@error, @extended, False)
Return $aCall[0]
EndFunc
Global Const $__COMBOBOXCONSTANT_EM_GETLINE = 0xC4
Global Const $__COMBOBOXCONSTANT_EM_LINEINDEX = 0xBB
Global Const $__COMBOBOXCONSTANT_EM_LINELENGTH = 0xC1
Global Const $__COMBOBOXCONSTANT_EM_REPLACESEL = 0xC2
Global Const $__COMBOBOXCONSTANT_WM_SETREDRAW = 0x000B
Global Const $tagCOMBOBOXINFO = "dword Size;struct;long EditLeft;long EditTop;long EditRight;long EditBottom;endstruct;" & "struct;long BtnLeft;long BtnTop;long BtnRight;long BtnBottom;endstruct;dword BtnState;hwnd hCombo;hwnd hEdit;hwnd hList"
Func _GUICtrlComboBox_AddString($hWnd, $sText)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $CB_ADDSTRING, 0, $sText, 0, "wparam", "wstr")
EndFunc
Func _GUICtrlComboBox_AutoComplete($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
If Not __GUICtrlComboBox_IsPressed('08') And Not __GUICtrlComboBox_IsPressed("2E") Then
Local $sEditText = _GUICtrlComboBox_GetEditText($hWnd)
If StringLen($sEditText) Then
Local $sInputText
Local $iRet = _GUICtrlComboBox_FindString($hWnd, $sEditText)
If($iRet <> $CB_ERR) Then
_GUICtrlComboBox_GetLBText($hWnd, $iRet, $sInputText)
_GUICtrlComboBox_SetEditText($hWnd, $sInputText)
_GUICtrlComboBox_SetEditSel($hWnd, StringLen($sEditText), StringLen($sInputText))
EndIf
EndIf
EndIf
EndFunc
Func _GUICtrlComboBox_BeginUpdate($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $__COMBOBOXCONSTANT_WM_SETREDRAW, False) = 0
EndFunc
Func _GUICtrlComboBox_EndUpdate($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $__COMBOBOXCONSTANT_WM_SETREDRAW, True) = 0
EndFunc
Func _GUICtrlComboBox_FindString($hWnd, $sText, $iIndex = -1)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $CB_FINDSTRING, $iIndex, $sText, 0, "int", "wstr")
EndFunc
Func _GUICtrlComboBox_GetComboBoxInfo($hWnd, ByRef $tInfo)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
$tInfo = DllStructCreate($tagCOMBOBOXINFO)
Local $iInfo = DllStructGetSize($tInfo)
DllStructSetData($tInfo, "Size", $iInfo)
Return _SendMessage($hWnd, $CB_GETCOMBOBOXINFO, 0, $tInfo, 0, "wparam", "struct*") <> 0
EndFunc
Func _GUICtrlComboBox_GetCurSel($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $CB_GETCURSEL)
EndFunc
Func _GUICtrlComboBox_GetEditText($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Local $tInfo
If _GUICtrlComboBox_GetComboBoxInfo($hWnd, $tInfo) Then
Local $hEdit = DllStructGetData($tInfo, "hEdit")
Local $iLine = 0
Local $iIndex = _SendMessage($hEdit, $__COMBOBOXCONSTANT_EM_LINEINDEX, $iLine)
Local $iLength = _SendMessage($hEdit, $__COMBOBOXCONSTANT_EM_LINELENGTH, $iIndex)
If $iLength = 0 Then Return ""
Local $tBuffer = DllStructCreate("short Len;wchar Text[" & $iLength & "]")
DllStructSetData($tBuffer, "Len", $iLength)
Local $iRet = _SendMessage($hEdit, $__COMBOBOXCONSTANT_EM_GETLINE, $iLine, $tBuffer, 0, "wparam", "struct*")
If $iRet = 0 Then Return SetError(-1, -1, "")
Local $tText = DllStructCreate("wchar Text[" & $iLength & "]", DllStructGetPtr($tBuffer))
Return DllStructGetData($tText, "Text")
Else
Return SetError(-1, -1, "")
EndIf
EndFunc
Func _GUICtrlComboBox_GetLBText($hWnd, $iIndex, ByRef $sText)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Local $iLen = _GUICtrlComboBox_GetLBTextLen($hWnd, $iIndex)
Local $tBuffer = DllStructCreate("wchar Text[" & $iLen + 1 & "]")
Local $iRet = _SendMessage($hWnd, $CB_GETLBTEXT, $iIndex, $tBuffer, 0, "wparam", "struct*")
If($iRet == $CB_ERR) Then Return SetError($CB_ERR, $CB_ERR, $CB_ERR)
$sText = DllStructGetData($tBuffer, "Text")
Return $iRet
EndFunc
Func _GUICtrlComboBox_GetLBTextLen($hWnd, $iIndex)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $CB_GETLBTEXTLEN, $iIndex)
EndFunc
Func _GUICtrlComboBox_ReplaceEditSel($hWnd, $sText)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Local $tInfo
If _GUICtrlComboBox_GetComboBoxInfo($hWnd, $tInfo) Then
Local $hEdit = DllStructGetData($tInfo, "hEdit")
_SendMessage($hEdit, $__COMBOBOXCONSTANT_EM_REPLACESEL, True, $sText, 0, "wparam", "wstr")
EndIf
EndFunc
Func _GUICtrlComboBox_ResetContent($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
_SendMessage($hWnd, $CB_RESETCONTENT)
EndFunc
Func _GUICtrlComboBox_SetCurSel($hWnd, $iIndex = -1)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $CB_SETCURSEL, $iIndex)
EndFunc
Func _GUICtrlComboBox_SetEditSel($hWnd, $iStart, $iStop)
If Not HWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $CB_SETEDITSEL, 0, _WinAPI_MakeLong($iStart, $iStop)) <> -1
EndFunc
Func _GUICtrlComboBox_SetEditText($hWnd, $sText)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
_GUICtrlComboBox_SetEditSel($hWnd, 0, -1)
_GUICtrlComboBox_ReplaceEditSel($hWnd, $sText)
EndFunc
Func __GUICtrlComboBox_IsPressed($sHexKey, $vDLL = 'user32.dll')
Local $a_R = DllCall($vDLL, "short", "GetAsyncKeyState", "int", '0x' & $sHexKey)
If @error Then Return SetError(@error, @extended, False)
Return BitAND($a_R[0], 0x8000) <> 0
EndFunc
Global Const $GUI_EVENT_CLOSE = -3
Global Const $GUI_RUNDEFMSG = 'GUI_RUNDEFMSG'
Global Const $GUI_ENABLE = 64
Global Const $GUI_DISABLE = 128
Global Const $WS_VSCROLL = 0x00200000
Global Const $WM_SETICON = 0x0080
Global Const $WM_COMMAND = 0x0111
Func _WinAPI_EnumProcessThreads($iPID = 0)
If Not $iPID Then $iPID = @AutoItPID
Local $hSnapshot = DllCall('kernel32.dll', 'handle', 'CreateToolhelp32Snapshot', 'dword', 0x00000004, 'dword', 0)
If @error Or Not $hSnapshot[0] Then Return SetError(@error + 10, @extended, 0)
Local Const $tagTHREADENTRY32 = 'dword Size;dword Usage;dword ThreadID;dword OwnerProcessID;long BasePri;long DeltaPri;dword Flags'
Local $tTHREADENTRY32 = DllStructCreate($tagTHREADENTRY32)
Local $aRet[101] = [0]
$hSnapshot = $hSnapshot[0]
DllStructSetData($tTHREADENTRY32, 'Size', DllStructGetSize($tTHREADENTRY32))
Local $aCall = DllCall('kernel32.dll', 'bool', 'Thread32First', 'handle', $hSnapshot, 'struct*', $tTHREADENTRY32)
While Not @error And $aCall[0]
If DllStructGetData($tTHREADENTRY32, 'OwnerProcessID') = $iPID Then
__Inc($aRet)
$aRet[$aRet[0]] = DllStructGetData($tTHREADENTRY32, 'ThreadID')
EndIf
$aCall = DllCall('kernel32.dll', 'bool', 'Thread32Next', 'handle', $hSnapshot, 'struct*', $tTHREADENTRY32)
WEnd
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hSnapshot)
If Not $aRet[0] Then Return SetError(1, 0, 0)
__Inc($aRet, -1)
Return $aRet
EndFunc
Func _WinAPI_EnumProcessWindows($iPID = 0, $bVisible = True)
Local $aThreads = _WinAPI_EnumProcessThreads($iPID)
If @error Then Return SetError(@error, @extended, 0)
Local $hEnumProc = DllCallbackRegister('__EnumWindowsProc', 'bool', 'hwnd;lparam')
Dim $__g_vEnum[101][2] = [[0]]
For $i = 1 To $aThreads[0]
DllCall('user32.dll', 'bool', 'EnumThreadWindows', 'dword', $aThreads[$i], 'ptr', DllCallbackGetPtr($hEnumProc), 'lparam', $bVisible)
If @error Then
ExitLoop
EndIf
Next
DllCallbackFree($hEnumProc)
If Not $__g_vEnum[0][0] Then Return SetError(11, 0, 0)
__Inc($__g_vEnum, -1)
Return $__g_vEnum
EndFunc
Func _WinAPI_CopyIcon($hIcon)
Local $aCall = DllCall("user32.dll", "handle", "CopyIcon", "handle", $hIcon)
If @error Then Return SetError(@error, @extended, 0)
Return $aCall[0]
EndFunc
Func _WinAPI_DestroyIcon($hIcon)
Local $aCall = DllCall("user32.dll", "bool", "DestroyIcon", "handle", $hIcon)
If @error Then Return SetError(@error, @extended, False)
Return $aCall[0]
EndFunc
Func _VersionCompare($sVersion1, $sVersion2)
If $sVersion1 = $sVersion2 Then Return 0
Local $sSubVersion1 = "", $sSubVersion2 = ""
If StringIsAlpha(StringRight($sVersion1, 1)) Then
$sSubVersion1 = StringRight($sVersion1, 1)
$sVersion1 = StringTrimRight($sVersion1, 1)
EndIf
If StringIsAlpha(StringRight($sVersion2, 1)) Then
$sSubVersion2 = StringRight($sVersion2, 1)
$sVersion2 = StringTrimRight($sVersion2, 1)
EndIf
Local $aVersion1 = StringSplit($sVersion1, ".,"), $aVersion2 = StringSplit($sVersion2, ".,")
Local $iPartDifference =($aVersion1[0] - $aVersion2[0])
If $iPartDifference < 0 Then
ReDim $aVersion1[UBound($aVersion2)]
$aVersion1[0] = UBound($aVersion1) - 1
For $i =(UBound($aVersion1) - Abs($iPartDifference)) To $aVersion1[0]
$aVersion1[$i] = "0"
Next
ElseIf $iPartDifference > 0 Then
ReDim $aVersion2[UBound($aVersion1)]
$aVersion2[0] = UBound($aVersion2) - 1
For $i =(UBound($aVersion2) - Abs($iPartDifference)) To $aVersion2[0]
$aVersion2[$i] = "0"
Next
EndIf
For $i = 1 To $aVersion1[0]
If StringIsDigit($aVersion1[$i]) And StringIsDigit($aVersion2[$i]) Then
If Number($aVersion1[$i]) > Number($aVersion2[$i]) Then
Return SetExtended(2, 1)
ElseIf Number($aVersion1[$i]) < Number($aVersion2[$i]) Then
Return SetExtended(2, -1)
ElseIf $i = $aVersion1[0] Then
If $sSubVersion1 > $sSubVersion2 Then
Return SetExtended(3, 1)
ElseIf $sSubVersion1 < $sSubVersion2 Then
Return SetExtended(3, -1)
EndIf
EndIf
Else
If $aVersion1[$i] > $aVersion2[$i] Then
Return SetExtended(1, 1)
ElseIf $aVersion1[$i] < $aVersion2[$i] Then
Return SetExtended(1, -1)
EndIf
EndIf
Next
Return SetExtended(Abs($iPartDifference), 0)
EndFunc
Global Const $HGDI_ERROR = Ptr(-1)
Global Const $INVALID_HANDLE_VALUE = Ptr(-1)
Global Const $KF_EXTENDED = 0x0100
Global Const $KF_ALTDOWN = 0x2000
Global Const $KF_UP = 0x8000
Global Const $LLKHF_EXTENDED = BitShift($KF_EXTENDED, 8)
Global Const $LLKHF_ALTDOWN = BitShift($KF_ALTDOWN, 8)
Global Const $LLKHF_UP = BitShift($KF_UP, 8)
Global $__Binary_Kernel32Dll = DllOpen("kernel32.dll")
Global $__Binary_User32Dll = DllOpen("user32.dll")
Global $__Binary_MsvcrtDll = DllOpen("msvcrt.dll")
Global $_IconImage_PNG_Header = Binary("0x89504E470D0A1A0A")
Global Const $tagPROPERTYKEY = 'struct;ulong Data1;ushort Data2;ushort Data3;byte Data4[8];DWORD pid;endstruct'
Global $tagPROPVARIANT = 'USHORT vt;' & 'WORD wReserved1;' & 'WORD wReserved2;' & 'WORD wReserved3;' & 'LONG;PTR'
Global Const $sIID_IPropertyStore = '{886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99}'
Global Const $VT_EMPTY = 0, $VT_LPWSTR = 31
Func _WindowAppId($hWnd, $appid = Default)
Local $tpIPropertyStore = DllStructCreate('ptr')
_WinAPI_SHGetPropertyStoreForWindow($hWnd, $sIID_IPropertyStore, $tpIPropertyStore)
Local $pPropertyStore = DllStructGetData($tpIPropertyStore, 1)
Local $oPropertyStore = ObjCreateInterface($pPropertyStore, $sIID_IPropertyStore, 'GetCount HRESULT(PTR);GetAt HRESULT(DWORD; PTR);GetValue HRESULT(PTR;PTR);' & 'SetValue HRESULT(PTR;PTR);Commit HRESULT()')
If Not IsObj($oPropertyStore) Then Return SetError(1, 0, '')
Local $tPKEY = _PKEY_AppUserModel_ID()
Local $tPROPVARIANT = DllStructCreate($tagPROPVARIANT)
Local $sAppId
If $appid = Default Then
$oPropertyStore.GetValue(DllStructGetPtr($tPKEY), DllStructGetPtr($tPROPVARIANT))
If DllStructGetData($tPROPVARIANT, 'vt') <> $VT_EMPTY Then
Local $buf = DllStructCreate('wchar[128]')
DllCall('Propsys.dll', 'long', 'PropVariantToString', 'ptr', DllStructGetPtr($tPROPVARIANT), 'ptr', DllStructGetPtr($buf), 'uint', DllStructGetSize($buf))
If Not @error Then
$sAppId = DllStructGetData($buf, 1)
EndIf
EndIf
Else
_WinAPI_InitPropVariantFromString($appId, $tPROPVARIANT)
$oPropertyStore.SetValue(DllStructGetPtr($tPKEY), DllStructGetPtr($tPROPVARIANT))
$oPropertyStore.Commit()
$sAppId = $appid
EndIf
Return SetError(($sAppId == '')*2, 0, $sAppId)
EndFunc
Func _WinAPI_InitPropVariantFromString($sUnicodeString, ByRef $tPROPVARIANT)
DllStructSetData($tPROPVARIANT, 'vt', $VT_LPWSTR)
Local $aRet = DllCall('Shlwapi.dll', 'LONG', 'SHStrDupW', 'WSTR', $sUnicodeString, 'PTR', DllStructGetPtr($tPROPVARIANT) + 8)
If @error Then Return SetError(@error, @extended, False)
Local $bSuccess = $aRet[0] == 0
If(Not $bSuccess) Then $tPROPVARIANT = DllStructCreate($tagPROPVARIANT)
Return SetExtended($aRet[0], $bSuccess)
EndFunc
Func _PKEY_AppUserModel_ID()
Local $tPKEY = DllStructCreate($tagPROPERTYKEY)
_WinAPI_GUIDFromStringEx('{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}', DllStructGetPtr($tPKEY))
DllStructSetData($tPKEY, 'pid', 5)
Return $tPKEY
EndFunc
Func _WinAPI_SHGetPropertyStoreForWindow($hWnd, $sIID, ByRef $tPointer)
Local $tIID = _WinAPI_GUIDFromString($sIID)
Local $pp = IsPtr($tPointer)? $tPointer : DllStructGetPtr($tPointer)
Local $aRet = DllCall('Shell32.dll', 'LONG', 'SHGetPropertyStoreForWindow', 'HWND', $hWnd, 'STRUCT*', $tIID, 'PTR', $pp)
If @error Then Return SetError(@error, @extended, False)
Return SetExtended($aRet[0],($aRet[0] = 0))
EndFunc
Global Const $VERSION = "1.0.0.18"
Global $sTitle="ctSpaces v"&$VERSION&"b"
Global $iWinPoll=125
Global $aFiles[]=[0]
Global $g_sDataDir=@LocalAppDataDir&"\InfinitySys\ctSpaces"
Global $sInstFullPath=StringFormat("%s\%s",$g_sDataDir,@ScriptName)
Global $g_sLog=$g_sDataDir&"\ctSpaces.log"
FileClose(FileOpen($g_sLog,2))
Func _Log($sLine)
If Not FileExists($g_sDataDir) Then DirCreate($g_sDataDir)
If FileGetSize($g_sLog)>1024*1024 Then FileDelete($g_sLog)
FileWriteLine($g_sLog,$sLine)
ConsoleWrite($sLine&@CRLF)
EndFunc
If StringInStr($CmdLineRaw,"~!Install") Then
ctInstall()
Exit 0
ElseIf StringInStr($CmdLineRaw,"~!Version") Then
ConsoleWrite($VERSION&@CRLF)
Exit 0
EndIf
If @Compiled Then
If FileExists($sInstFullPath) Then
If ctUpdate() Then Exit
Else
ctInstall()
EndIf
EndIf
Local $aKeepDefault[]=[0, "Local State", "Last Version", "Last Browser", "FirstLaunchAfterInstallation", "First Run", "DevToolsActivePort", "Default\Shortcuts", "Default\Shortcuts-journal", "Default\Secure Preferences", "Default\Preferences", "Default\Favicons-journal", "Default\Favicons", "Default\Bookmarks", "Default\Extension State", "Default\Extensions", "Default\Local Extension Settings" ]
Local $aKeepActive[]=[0, "client.ico", "client.png", "Local State", "Last Version", "Last Browser", "FirstLaunchAfterInstallation", "First Run", "DevToolsActivePort", "Default\History", "Default\Shortcuts", "Default\Shortcuts-journal", "Default\Secure Preferences", "Default\Preferences", "Default\Favicons-journal", "Default\Favicons", "Default\Bookmarks", "Default\Extension State", "Default\Extensions", "Default\Local Extension Settings", "CertificateRevocation", "AutoLaunchProtocolsComponent", "Default\History", "Default\Web Data", "Default\Web Data-journal", "Default\Login Data", "Default\Login Data-journal", "Default\Favicons", "Default\Favicons-journal", "Default\MediaDeviceSalts", "Default\MediaDeviceSalts-journal", "Default\CdmStorage.db", "Default\CdmStorage.db-journal", "Default\DIPS", "Default\DIPS-journal", "Default\Local Storage", "Default\WebStorage", "Default\Service Worker\Database", "Default\ClientCertificates", "Default\blob_storage", "Default\Session Storage", "Default\IndexedDB", "Default\Network", "Default\Sessions" ]
If Not FileExists($g_sDataDir&'\Sites') Then DirCreate($g_sDataDir&'\Sites')
FileInstall("7za.exe",$g_sDataDir&"\7za.exe",1)
FileInstall("Default.7z",$g_sDataDir&"\Default.7z",0)
Global $aClients=_FileListToArrayRec($g_sDataDir&'\Sites','*',2)
$aKeepDefault[0]=UBound($aKeepDefault,1)-1
$aKeepActive[0]=UBound($aKeepActive,1)-1
Local $iGuiW=256+16,$iGuiH=64+16+4,$iGuiM=4
Opt("GUIOnEventMode",1)
$hGui=GUICreate($sTitle, $iGuiW, $iGuiH)
GUISetFont(9,400,0,"Consolas")
$idClient=GUICtrlCreateCombo("",$iGuiM,25,($iGuiW-$iGuiM*2),25,BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL,$WS_VSCROLL))
$hClient=GUICtrlGetHandle($idClient)
GUICtrlCreateLabel("Select or type the client name:",$iGuiM,$iGuiM,($iGuiW-$iGuiM*2),17)
$idBtnGo=GUICtrlCreateButton("Go",($iGuiW/2)-32,25*2+$iGuiM,64,25,0x0001)
Local $aBtn[][4]=[[0,0], [-1,"I","Set Profile Icon","_GuiSetIcon"], [-1,"R","Reset Selected Profile to Default (complete purge)","_GuiProfReset"], [-1,"U","Update Selected Profile with Default (overwrites bookmarks, settings, etc)","_GuiProfUpd"], [-1,"D","Open Default Profile","_GuiOpenDef"], [-1,"T","Launch temporary profile","_GuiOpenTmp"] ]
$aBtn[0][0]=UBound($aBtn,1)-1
If StringInStr($CmdLineRaw,"~!Trd:P",1) Then
Local $iMax,$iMaxY=UBound($aBtn,2)
Local $aTrd[1][$iMaxY],$aMap=[1,5,3,2,4]
For $i=0 To UBound($aMap,1)-1
$iMax=UBound($aTrd,1)
ReDim $aTrd[$iMax+1][$iMaxY]
For $j=0 To $iMaxY-1
$aTrd[$iMax][$j]=$aBtn[$aMap[$i]][$j]
Next
Next
$aTrd[0][0]=$iMax
$aBtn=$aTrd
EndIf
Local $iBtnS=16,$iBtnM=2
Local $iBtnL=$iGuiW-(($iBtnS+$iBtnM)*($aBtn[0][0]+1))
Local $iBtnT=$iGuiH-$iBtnM-$iBtnS
For $i=1 To $aBtn[0][0]
$aBtn[$i][0]=GUICtrlCreateButton($aBtn[$i][1],$iBtnL+(($iBtnS+$iBtnM)*$i),$iBtnT,$iBtnS,$iBtnS)
GUICtrlSetTip($aBtn[$i][0],$aBtn[$i][2])
GUICtrlSetOnEvent($aBtn[$i][0],$aBtn[$i][3])
Next
GUICtrlSetOnEvent($idBtnGo,"_GuiProfOpen")
_updClients()
GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
GUISetOnEvent($GUI_EVENT_CLOSE, "_GuiClose")
GUISetState(@SW_SHOW)
While Sleep(1000)
WEnd
Func _GuiClose()
Exit
EndFunc
Func _WarnSelClient()
MsgBox(48,$sTitle,"Please select a client first.")
EndFunc
Func _GuiProfOpen()
$sClient=GUICtrlRead($idClient)
If $sClient="" Then Return _WarnSelClient()
_SetUiState(0)
Local $sData=StringFormat("%s\Sites\%s",$g_sDataDir,$sClient)
_initClient($sData,$sClient)
_SetUiState(1)
_updClients()
EndFunc
Func _SetUiState($bEn)
If Not $bEn Then GUISetState(@SW_HIDE,$hGui)
GUICtrlSetState($idClient,$bEn?$GUI_ENABLE:$GUI_DISABLE)
GUICtrlSetState($idBtnGo,$bEn?$GUI_ENABLE:$GUI_DISABLE)
For $i=1 To $aBtn[0][0]
GUICtrlSetState($aBtn[$i][0],$bEn?$GUI_ENABLE:$GUI_DISABLE)
Next
GUICtrlSetState($idClient,256)
If $bEn Then GUISetState(@SW_SHOW,$hGui)
EndFunc
Func _initClient($sData,$sClient)
$g_sErrFunc="_initClient"
Local $sLock=StringFormat("%s\ctSpaces.lock",$sData)
_SetLock($sLock)
If @error Then
If @error=4 Then
MsgBox(48,$sTitle,"Warning: This site is already open in another instance.")
Return
EndIf
Return _Log(StringFormat("_SetLock,%d",@Error))
EndIf
If Not FileExists($sData) Then
_extDef($sData)
If @error Then
MsgBox(16,$sTitle,"Error: An error occurred when extracting profile.")
Return
EndIf
EndIf
Local $sIconIco=StringFormat("%s\client.ico",$sData)
Local $sIconImg=StringFormat("%s\client.png",$sData)
Local $bHasIcon=FileExists($sIconIco)
Local $bHasIconImg=FileExists($sIconImg)
Local $bIcoErr=0
Local $hImage,$hIcon,$bIcon=0
$bHasIcon=FileExists($sIconIco)
If $bHasIcon Then
$hImage=_WinAPI_LoadImage(0,$sIconIco,$IMAGE_ICON,0,0,BitOR($LR_LOADFROMFILE,$LR_DEFAULTSIZE))
If @error Then $bIcoErr=2
$hIcon=_WinAPI_CopyIcon($hImage)
If @error Then $bIcoErr=3
If $bIcoErr=0 Then $bIcon=1
EndIf
$aSearch=_FileListToArrayRec($sData,"*",1,1)
Dim $aFiles[]=[0]
Local $sTitleRegEx="(.*) - Profile 1 - Microsoft.*Edge"
$iPid=Run('"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --user-data-dir="'&$sData&'" --no-first-run --disable-sync --disable-features=SyncPromo  --edge-skip-compat-layer-relaunch --no-service-autorun')
While ProcessExists($iPid)
$aWin=_WinAPI_EnumProcessWindows($iPid,1)
If @error Then
Sleep($iWinPoll)
ContinueLoop
EndIf
For $j=1 To $aWin[0][0]
_WindowAppId($aWin[$j][0],StringFormat("ctSpaces.%s.Default",$sClient))
$sWinTitle=WinGetTitle($aWin[$j][0])
If StringRegExp($sWinTitle,$sTitleRegEx,0) Then
WinSetTitle($aWin[$j][0],"",StringRegExpReplace($sWinTitle,$sTitleRegEx,StringFormat("\1 - %s - ctSpaces",$sClient)))
EndIf
If $bIcon Then
_SendMessage($aWin[$j][0],$WM_SETICON,0,$hIcon)
_SendMessage($aWin[$j][0],$WM_SETICON,1,$hImage)
EndIf
Next
Sleep($iWinPoll)
WEnd
_WinAPI_DestroyIcon($hIcon)
_WinAPI_DeleteObject($hImage)
$aSearch=_FileListToArrayRec($sData,"*",1,1)
If $sData=$g_sDataDir&'\Default' Then
For $j=1 To $aSearch[0]
If Not _flt($aKeepDefault,$aSearch[$j]) Then _icl($aSearch[$j])
Next
Else
For $j=1 To $aSearch[0]
If Not _flt($aKeepActive,$aSearch[$j]) And Not _flt($aKeepActive,$aSearch[$j]) Then _icl($aSearch[$j])
Next
EndIf
For $j=1 To $aFiles[0]
FileDelete($sData&"\"&$aFiles[$j])
Next
_delEmpty($sData)
_UnsetLock($sLock)
EndFunc
Func _extDef($sData)
$g_sErrFunc="_extDef"
$sData=StringStripWS($sData,7)
$sData=StringRegExpReplace($sData,'[\\/:*?"<>|]',"")
$sData=StringRegExpReplace($sData,'[\. ]+$','')
If $sData = '' Or StringRegExp($sData, '(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$') Then Return SetError(1,0,0)
If Not DirCreate($sData) Then
_Log("DirCreate")
EndIf
Local $sCmd='"'&$g_sDataDir&'\7za.exe" x "'&$g_sDataDir&'\Default.7z" -y -o"'&$sData&'"'
Local $iRet=RunWait($sCmd,$g_sDataDir)
If $iRet<>0 Then
_Log(StringFormat("CmdRet:%d;Cmd:%s",$iRet,$sCmd))
Return SetError(2,0,0)
EndIf
Return SetError(0,0,1)
EndFunc
Func _updClients()
Local $iItem=_GUICtrlComboBox_GetCurSel($hClient)
Local $sItem
If $iItem<>-1 And $iItem+1<=$aClients[0] Then
$sItem=$aClients[$iItem+1]
EndIf
$aClients=_FileListToArrayRec($g_sDataDir&'\Sites','*',2)
If @error Then Return
_GUICtrlComboBox_BeginUpdate($hClient)
_GUICtrlComboBox_ResetContent($hClient)
For $i=1 To $aClients[0]
_GUICtrlComboBox_AddString($hClient,$aClients[$i])
Next
If $sItem<>"" Then
$iItem=-1
For $i=1 To $aClients[0]
If $aClients[$i]=$sItem Then
$iItem=$i
EndIf
Next
If $iItem<>-1 Then _GUICtrlComboBox_SetCurSel($hClient,$iItem-1)
EndIf
_GUICtrlComboBox_EndUpdate($hClient)
EndFunc
Func WM_COMMAND($hWnd,$iMsg,$wParam,$lParam)
#forceref $hWnd,$iMsg
Local $hWndCombo=$hClient
If Not IsHWnd($hClient) Then $hWndCombo=GUICtrlGetHandle($hClient)
Local $hWndFrom=$lParam
Local $iIDFrom=BitAND($wParam, 0xFFFF)
Local $iCode=BitShift($wParam, 16)
Switch $hWndFrom
Case $hClient,$hWndCombo
Switch $iCode
Case $CBN_EDITCHANGE
_Edit_Changed()
EndSwitch
EndSwitch
Return $GUI_RUNDEFMSG
EndFunc
Func _Edit_Changed()
_GUICtrlComboBox_AutoComplete($hClient)
EndFunc
Func _icl($sStr)
Local $iMax=UBound($aFiles,1)
For $i=1 To $iMax-1
If $aFiles[$i]=$sStr Then Return
Next
ReDim $aFiles[$iMax+1]
$aFiles[$iMax]=$sStr
$aFiles[0]=$iMax
EndFunc
Func _flt(ByRef $aFilter,$sStr)
For $i=1 To $aFilter[0]
If StringInStr($sStr,$aFilter[$i]) Then Return True
Next
Return False
EndFunc
Func _delEmpty($dir)
$folder_list = _FileListToArray($dir, '*', 2)
If @error <> 4 Then
For $i = 1 To $folder_list[0]
_delEmpty($dir & '\' & $folder_list[$i])
Next
EndIf
FileFindFirstFile($dir & '\*')
If @error Then DirRemove($dir)
FileClose($dir)
EndFunc
Func ctInstall()
If Not @Compiled Then Return
Local $bStartup,$bDesktop,$bStartMenu
If MsgBox(32+4,$sTitle,"Would you like to run at startup?")==6 Then $bStartup=1
If MsgBox(32+4,$sTitle,"Would you like to add to the desktop shortcut?")==6 Then $bDesktop=1
If MsgBox(32+4,$sTitle,"Would you like to add to the Start Menu?")==6 Then $bStartMenu=1
If $bStartup Or $bDesktop Or $bStartMenu Then
FileCopy(@AutoItExe,$g_sDataDir&"\ctSpaces.exe",9)
EndIf
If $bStartup Then
RegWrite("HKCU\Software\Microsoft\Windows\CurrentVersion\Run","ctSpaces","REG_SZ",$g_sDataDir&"\ctSpaces.exe")
EndIf
If $bDesktop Then
FileCreateShortcut($g_sDataDir&"\ctSpaces.exe",@DesktopDir&"\ctSpaces.lnk",$g_sDataDir)
EndIf
If $bStartMenu Then
FileCreateShortcut($g_sDataDir&"\ctSpaces.exe",@ProgramsDir&"\ctSpaces.lnk",$g_sDataDir)
EndIf
If MsgBox(32+4,$sTitle,"Would you like to run now?")==6 Then
Run($sInstFullPath&" ~!PostInstall",$g_sDataDir,@SW_SHOW)
Exit 0
EndIf
Exit 0
EndFunc
Func ctUpdate()
$sOldVer=FileGetVersion($sInstFullPath,"FileVersion")
$sCurVer=FileGetVersion(@AutoItExe,"FileVersion")
If _VersionCompare($sOldVer,$sCurVer)<>-1 Then Return 0
Local $bUpdate=0,$bRun=0
If MsgBox(32+4,$sTitle,"This version is newer than the one that is current installed, would you like to update it?")==6 Then $bUpdate=1
If Not $bUpdate Then
If MsgBox(32+4,$sTitle,"Would you like to run the new version now?")==6 Then Return 0
Return 1
EndIf
If $bUpdate Then
Local $iPid=-1,$aProc=ProcessList(@ScriptName)
For $i=1 To $aProc[0][0]
_Log($aProc[$i][1]&","&@AutoItPID)
If $aProc[$i][1]<>@AutoItPID Then
$iPid=$aProc[$i][1]
EndIf
Next
If $iPid<>-1 And ProcessExists($iPid) Then
_Log("Killing pid: "&$iPid)
ProcessClose($iPid)
ProcessWaitClose($iPid)
EndIf
_Log(FileCopy(@AutoItExe,$sInstFullPath,1))
EndIf
Local $iRet=MsgBox(32+4,$sTitle,"Would you like to run the new version after the update?")
If $iRet=6 Then $bRun=1
If $bRun Then
Run($sInstFullPath&" ~!PostInstall",$g_sDataDir,@SW_SHOW)
EndIf
Return 1
EndFunc
Func _SetLock($sPath)
Local $hFile
If FileExists($sPath) Then
$hFile=FileOpen($sPath)
If $hFile=-1 Then Return SetError(1,0,0)
$vData=FileRead($sPath)
If @error Then Return SetError(2,0,0)
FileClose($hFile)
If Not IsInt($sPath) Then Return SetError(3,0,0)
If @AutoItPID=$vData Then Return SetError(0,0,1)
If ProcessExists($vData) Then Return SetError(4,0,0)
If Not FileDelete($sPath) Then Return SetError(5,0,0)
EndIf
$hFile=FileOpen($sPath,10)
If $hFile=-1 Then Return SetError(6,0,0)
If Not FileWrite($hFile,@AutoItPID) Then Return SetError(7,0,0)
If Not FileClose($hFile) Then Return SetError(8,0,0)
Return SetError(0,1,1)
EndFunc
Func _UnsetLock($sPath)
If Not FileExists($sPath) Then Return SetError(1,0,0)
If Not FileDelete($sPath) Then Return SetError(1,1,0)
Return SetError(0,0,1)
EndFunc
