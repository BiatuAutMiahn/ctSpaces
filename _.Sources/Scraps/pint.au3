$objShell = ObjCreate("Shell.Application")
$objFSO = ObjCreate("Scripting.FileSystemObject")
$strlPath = "C:\Users\User\Desktop\WampServer.lnk"
$strFolder = $objFSO.GetParentFolderName($strlPath)
$strFile = $objFSO.GetFileName($strlPath)
$objFolder = $objShell.Namespace($strFolder)
$objFolderItem = $objFolder.ParseName($strFile)
$colVerbs = $objFolderItem.Verbs
For $itemverb in $objFolderItem.verbs
    If StringReplace($itemverb.name, "&", "") = "Anclar a la barra de tareas" Then $itemverb.DoIt
Next


#include <Constants.au3>

If @OSVersion <> "WIN_10" Then Exit MsgBox ($MB_SYSTEMMODAL,"Error","This function only works on Win 10")
_ListSoftwares()
If @error Then Exit MsgBox ($MB_SYSTEMMODAL,"Error", @error)
_SoftwareVerbs("Microsoft Store")
If @error Then Exit MsgBox ($MB_SYSTEMMODAL,"Error", @error)

Func _ListSoftwares()
  $ObjShell = ObjCreate("Shell.Application")
  $objFolder = $ObjShell.Namespace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}')
  $objFolderItems = $objFolder.Items()
  For $oObj In $objFolderItems
    ConsoleWrite ($oObj.name & @CRLF)
  Next
EndFunc

Func _SoftwareVerbs($sSoftware, $sVerbToApply = "List")
  $ObjShell = ObjCreate("Shell.Application")
  $objFolder = $ObjShell.Namespace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}')
  If Not IsObj ($objFolder) Then Return SetError (1)
  $objFolderItems = $objFolder.Items ()
  Local $bFound = False, $sList
  For $oObj In $objFolderItems
    If $oObj.name = $sSoftware Then
      $bFound = True
      ExitLoop
    EndIf
  Next
  If Not $bFound Then Return SetError (2)
  For $oVerb in $oObj.Verbs ()
    If $sVerbToApply = "List" Then
      $sList &= $oVerb.Name & @CRLF
    ElseIf $oVerb.Name = $sVerbToApply Then
      $oVerb.DoIt ()
      Return
    EndIf
  Next
  If $sVerbToApply = "List" Then
    MsgBox ($MB_SYSTEMMODAL,"",$sList)
  Else
    Return SetError (3)
  EndIf
EndFunc


#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/so /pe ;/rm
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#include <Array.au3>
#include <WinAPIGdi.au3>
#include "CUIAutomation2.au3" ; https://www.autoitscript.com/forum/topic/153520-iuiautomation-ms-framework-automate-chrome-ff-ie/


;by InnI - https://www.autoitscript.com/forum/topic/200541-solved-taskbar-icons-coordinates/?do=findComment&comment=1521328
Func _WinAPI_FindMyIconPosInTaskbar($sFileDescription)
    ; Search taskbars
    Local $ahWnd = WinList("[REGEXPCLASS:Shell_(Secondary)?TrayWnd]")
    ConsoleWrite("Found taskbars " & $ahWnd[0][0] & @CRLF)

    ; Search controls
    Local $ahCtrl[$ahWnd[0][0]][2], $aPos
    For $i = 1 To $ahWnd[0][0]
        $ahCtrl[$i - 1][0] = ControlGetHandle($ahWnd[$i][1], "", "MSTaskListWClass1")
        $ahCtrl[$i - 1][1] = WinGetPos($ahWnd[$i][1])
        $aPos = $ahCtrl[$i - 1][1]
        ConsoleWrite($aPos[0] & ", " & $aPos[1] & ", " & $aPos[2] & ", " & $aPos[3] & @CRLF)
    Next

    ; Get UIAutomation object
    Local $oUIAutomation = ObjCreateInterface($sCLSID_CUIAutomation, $sIID_IUIAutomation, $dtagIUIAutomation), $oElement, $oCondition, $oElementArray, $oButton
    If Not IsObj($oUIAutomation) Then Exit ConsoleWrite("Error create UIA object" & @CRLF)

    ; Create 2D array of buttons [name,left,top,right,bottom]
    Local $aBtnInfo[$ahWnd[0][0]][6], $Count, $pElement, $pCondition, $pElementArray, $iButtons, $vValue, $tPos, $hMonitor, $aMonitorPos
    $tPos = _WinAPI_GetMousePos()
    $hMonitor = _WinAPI_MonitorFromPoint($tPos)
    Local $tRect = DllStructCreate("long Left;long Top;long Right;long Bottom")
    For $n = 0 To UBound($ahCtrl) - 1
        ; Get taskbar element
        $oUIAutomation.ElementFromHandle($ahCtrl[$n][0], $pElement)
        $oElement = ObjCreateInterface($pElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
        ; Get condition (ControlType = Button)
        $oUIAutomation.CreatePropertyCondition($UIA_ControlTypePropertyId, $UIA_ButtonControlTypeId, $pCondition)
        $oCondition = ObjCreateInterface($pCondition, $sIID_IUIAutomationPropertyCondition, $dtagIUIAutomationPropertyCondition)
        ; Find all buttons
        $oElement.FindAll($TreeScope_Children, $oCondition, $pElementArray)
        $oElementArray = ObjCreateInterface($pElementArray, $sIID_IUIAutomationElementArray, $dtagIUIAutomationElementArray)
        $oElementArray.Length($iButtons)
        ;ReDim $aBtnInfo[UBound($aBtnInfo) + $iButtons][6]
        ; Get name and position for each button
        For $i = 0 To $iButtons - 1
            $oElementArray.GetElement($i, $pElement)
            $oButton = ObjCreateInterface($pElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
            $oButton.GetCurrentPropertyValue($UIA_NamePropertyId, $vValue)
            $oButton.CurrentBoundingRectangle($tRect)
            If StringInStr($vValue, $sFileDescription) Then
                $aMonitorPos = _WinAPI_GetMonitorInfo($hMonitor)
                If ($aMonitorPos[1].left = ($ahCtrl[$n][1])[0]) And ($aMonitorPos[1].bottom = ($ahCtrl[$n][1])[1]) Then
                    $aBtnInfo[$Count][0] = $vValue
                    $aBtnInfo[$Count][1] = $tRect.Left
                    $aBtnInfo[$Count][2] = $tRect.Top
                    $aBtnInfo[$Count][3] = $tRect.Right - $tRect.Left
                    $aBtnInfo[$Count][4] = $tRect.Bottom - $tRect.Top
                    $aBtnInfo[$Count][5] = $ahCtrl[$n][1]
                    ExitLoop 2
                EndIf
            EndIf
        Next
    Next
    _ArrayDisplay($aBtnInfo)
EndFunc

_WinAPI_FindMyIconPosInTaskbar("Opera Browser")

