#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile_type=a3x
#AutoIt3Wrapper_Icon=Res\ctdkgrsq.ico
#AutoIt3Wrapper_Outfile_x64=..\ctSpaces.a3x
#AutoIt3Wrapper_Res_Fileversion=1.0.0.24
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
Opt("TrayAutoPause", 0)
Opt("TrayIconHide", 1)

#include <Array.au3>
#include <Debug.au3>
#include <File.au3>
#include <GuiComboBox.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <WinAPIShPath.au3>
#include <SendMessage.au3>
#include <WinAPIProc.au3>
#include <WinAPIRes.au3>
#include <WinAPIIcons.au3>
#include <Math.au3>
#include <misc.au3>
#Include "Includes\IconImage.au3"
#include "Includes\AppUserModelId.au3"

Global Const $VERSION = "1.0.0.24"
Global $sTitle="ctSpaces v"&$VERSION&"b"
Global $iWinPoll=500
Global $aFiles[]=[0]
Global $bA3x=StringRight(@ScriptFullPath,3)="a3x"
Global $g_sDataDir=@LocalAppDataDir&"\InfinitySys\ctSpaces"
;If Not @Compiled Then $g_sDataDir=_WinAPI_PathCanonicalize(@ScriptDir&"\..")
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
;If ctUpdate() Then Exit


Local $aKeepDefault[]=[0, _
  "Local State", _
  "Last Version", _
  "Last Browser", _
  "FirstLaunchAfterInstallation", _
  "First Run", _
  "ctSpaces", _
  "DevToolsActivePort", _
  "Default\Shortcuts", _
  "Default\Shortcuts-journal", _
  "Default\Secure Preferences", _
  "Default\Preferences", _
  "Default\Favicons-journal", _
  "Default\Favicons", _
  "Default\Bookmarks", _
  "Default\Extension State", _
  "Default\Extensions", _
  "Default\Local Extension Settings" _
]

Local $aKeepActive[]=[0, _
  "client.ico", _
  "client.png", _
  "ctSpaces", _
  "Local State", _
  "Last Version", _
  "Last Browser", _
  "FirstLaunchAfterInstallation", _
  "First Run", _
  "DevToolsActivePort", _
  "Default\History", _
  "Default\Shortcuts", _
  "Default\Shortcuts-journal", _
  "Default\Secure Preferences", _
  "Default\Preferences", _
  "Default\Favicons-journal", _
  "Default\Favicons", _
  "Default\Bookmarks", _
  "Default\Extension State", _
  "Default\Extensions", _
  "Default\Local Extension Settings", _
  "CertificateRevocation", _
  "AutoLaunchProtocolsComponent", _
  "Default\History", _
  "Default\Web Data", _
  "Default\Web Data-journal", _
  "Default\Login Data", _
  "Default\Login Data-journal", _
  "Default\Favicons", _
  "Default\Favicons-journal", _
  "Default\MediaDeviceSalts", _
  "Default\MediaDeviceSalts-journal", _
  "Default\CdmStorage.db", _
  "Default\CdmStorage.db-journal", _
  "Default\DIPS", _
  "Default\DIPS-journal", _
  "Default\Local Storage", _
  "Default\WebStorage", _
  "Default\Service Worker\Database", _
  "Default\ClientCertificates", _
  "Default\blob_storage", _
  "Default\Session Storage", _
  "Default\IndexedDB", _
  "Default\Network", _
  "Default\Sessions" _
]
If Not FileExists($g_sDataDir&'\Sites') Then DirCreate($g_sDataDir&'\Sites')
FileInstall("7za.exe",$g_sDataDir&"\7za.exe",1)
_extDef7z()


Global $aClients=_FileListToArrayRec($g_sDataDir&'\Sites','*',2)
$aKeepDefault[0]=UBound($aKeepDefault,1)-1
$aKeepActive[0]=UBound($aKeepActive,1)-1

#Region ### START Koda GUI section ### Form=
Local $iGuiW=256+16,$iGuiH=64+16+4,$iGuiM=4
Opt("GUIOnEventMode",1)
$hGui=GUICreate($sTitle, $iGuiW, $iGuiH)
GUISetFont(9,400,0,"Consolas")
$idClient=GUICtrlCreateCombo("",$iGuiM,25,($iGuiW-$iGuiM*2),25,BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL,$WS_VSCROLL))
$hClient=GUICtrlGetHandle($idClient)
GUICtrlCreateLabel("Select or type the client name:",$iGuiM,$iGuiM,($iGuiW-$iGuiM*2),17)
$idBtnGo=GUICtrlCreateButton("Go",($iGuiW/2)-32,25*2+$iGuiM,64,25,0x0001)
Local $aBtn[][4]=[[0,0], _
  [-1,"I","Set Profile Icon","_GuiSetIcon"], _
  [-1,"R","Reset Selected Profile to Default (complete purge)","_GuiProfReset"], _
  [-1,"U","Update Selected Profile with Default (overwrites bookmarks, settings, etc)","_GuiProfUpd"], _
  [-1,"D","Open Default Profile","_GuiOpenDef"], _
  [-1,"T","Launch temporary profile","_GuiOpenTmp"] _
]
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
_WinAPI_EmptyWorkingSet()
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###
While Sleep(1000)
WEnd

Func _GuiClose()
  Exit
EndFunc

Func _WarnSelClient()
  MsgBox(48,$sTitle,"Please select a client first.")
EndFunc

Func _GuiProfOpen()
  $sClient=_SanitizeName(GUICtrlRead($idClient))
  If $sClient="" Then Return _WarnSelClient()
  _SetUiState(0)
  Local $sData=StringFormat("%s\Sites\%s",$g_sDataDir,$sClient)
  _initClient($sData,$sClient)
  _SetUiState(1)
  _updClients()
EndFunc

Func _GuiSetIcon()
  Local $sClient=_SanitizeName(GUICtrlRead($idClient))
  If $sClient="" Then Return _WarnSelClient()
  Local $sFmts=""
  Local $aFmts=_GdiGetTypes()
  If Not @error Then
    For $i=1 To $aFmts[0]
      $sFmts&="*."&$aFmts[$i]
      If $i<$aFmts[0] Then $sFmts&=';'
    Next
    If $sFmts<>"" Then $sFmts=";"&$sFmts
  EndIf
  $sIcoReq=FileOpenDialog("Choose Profile Icon","",StringFormat("Icon/Image Files (*.ico%s)",$sFmts),3,"",$hGui)
  If @error Then Return
  Local $sExt=StringTrimLeft($sIcoReq,StringInStr($sIcoReq,".",0,-1))
  If Not $sExt="ico" and Not _hasFmt($aFmts,$sExt) Then
    MsgBox(32,$sTitle,"This is not an accepted file type.")
    Return
  EndIf
  $sDestIcon=StringFormat("%s\Sites\%s\client.ico",$g_sDataDir,$sClient)
  If FileExists($sDestIcon) Then
    If Not FileDelete($sDestIcon) Then
      MsgBox(32,$sTitle,"Failed to remove existing icon.")
      Return
    EndIf
  EndIf
  If $sExt<>"ico" Then
    $sDestImage=StringFormat("%s\Sites\%s\client.%s",$g_sDataDir,$sClient,$sExt)
    If Not FileCopy($sIcoReq,$sDestImage,9) Then
      MsgBox(32,$sTitle,"Cannot copy image.")
      Return
    EndIf
    _png2ico($sDestImage,$sDestIcon)
    If @error Then
      MsgBox(32,$sTitle,"There was an error converting the image to an icon.")
      Return
    EndIf
  Else
    If Not FileCopy($sIcoReq,$sDestIcon,9) Then
      MsgBox(32,$sTitle,"Cannot copy icon.")
      Return
    EndIf
  EndIf
  If Not FileExists($sDestIcon) Then
    MsgBox(32,$sTitle,"There was an error preparing the icon.")
    Return
  EndIf
  MsgBox(64,$sTitle,"Icon has been configured.")
  Return
EndFunc

Func _GuiProfReset()
  $sClient=_SanitizeName(GUICtrlRead($idClient))
  If $sClient="" Then Return _WarnSelClient()
  _SetUiState(0)
  Local $sData=StringFormat("%s\Sites\%s",$g_sDataDir,$sClient)
  DirRemove($sData)
  _extDef($sData)
  If @error Then
    MsgBox(16,$sTitle,"Error: An error occurred when extracting profile.")
    Return
  EndIf
  _SetUiState(1)
EndFunc

Func _GuiProfUpd()
  $sClient=_SanitizeName(GUICtrlRead($idClient))
  If $sClient="" Then Return _WarnSelClient()
  _SetUiState(0)
  Local $sData=StringFormat("%s\Sites\%s",$g_sDataDir,$sClient)
  _extDef($sData)
  If @error Then
    MsgBox(16,$sTitle,"Error: An error occurred when extracting profile.")
    Return
  EndIf
  _SetUiState(1)
EndFunc

Func _GuiOpenDef()
  _SetUiState(0)
  Local $sData=$g_sDataDir&'\Default'
  _extDef($sData)
  If @error Then
    MsgBox(16,$sTitle,"Error: An error occurred when extracting profile.")
    Return
  EndIf
  _initClient($sData,"Default")
  If MsgBox(32+1,$sTitle,"Save changes to default profile?",0,$hGui)=1 Then
    FileMove($g_sDataDir&"\Default.7z",StringFormat("%s\_.DefBak\%04d.%02d.%02d,%02d%02d%02d- Default.7z",$g_sDataDir,@YEAR,@MON,@MDAY,@HOUR,@MIN,@SEC),9)
    RunWait('7za.exe a -mx9 -myx9 -mmemusep80 -mtm -mtc -mtr -ms16g -stl -slp "'&$g_sDataDir&'\Default.7z" "'&$sData&'\*"')
  EndIf
  DirRemove($sData,1)
  _SetUiState(1)
EndFunc

Func _GuiOpenTmp()
  _SetUiState(0)
  Local $sData=$g_sDataDir&'\Temp'
  _extDef($sData)
  If @error Then
    MsgBox(16,$sTitle,"Error: An error occurred when extracting profile.")
    Return
  EndIf
  _initClient($sData,"Default")
  DirRemove($sData,1)
  _SetUiState(1)
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
  If Not FileExists($sData) Or Not FileExists(StringFormat("%s\ctSpaces",$sData)) Then
    _extDef($sData)
    If @error Then
      MsgBox(16,$sTitle,"Error: An error occurred when extracting profile.")
      Return
    EndIf
  EndIf
  Local $sLock=StringFormat("%s\ctSpaces.lock",$sData)
  _SetLock($sLock)
  If @error Then
    If @error=4 Then
      MsgBox(48,$sTitle,"Warning: This site is already open in another instance.")
      Return
    EndIf
    Return _Log(StringFormat("_SetLock,%d",@Error))
  EndIf
  _Log($sData)
  _Log(FileExists($sData))
  Local $sIconIco=StringFormat("%s\client.ico",$sData)
  Local $sIconImg=StringFormat("%s\client.png",$sData)
  Local $bHasIcon=FileExists($sIconIco)
  Local $bHasIconImg=FileExists($sIconImg)
  Local $bIcoErr=0
  ;If Not $bHasIcon And $bHasIconImg Then
  ;  _png2ico($sIconImg,$sIconIco)
  ;  If @error Then $bIcoErr=1
  ;EndIf
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
  ;_DebugArrayDisplay($aSearch)
  Dim $aFiles[]=[0]
  _WinAPI_EmptyWorkingSet()
  Local $sTitleRegEx="(.*) - Profile 1 - Microsoft.*Edge"
  $iPid=Run('"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --user-data-dir="'&$sData&'" --no-first-run --disable-sync --disable-features=SyncPromo  --edge-skip-compat-layer-relaunch --no-service-autorun')
  While ProcessExists($iPid)
    ;$hTimer=TimerInit()
    $aWin=_WinGetHandleByPID($iPid)
    If @error Then
      Sleep($iWinPoll)
      ContinueLoop
    EndIf
    For $j=1 To $aWin[0]
      _WindowAppId($aWin[$j],StringFormat("ctSpaces.%s.Default",$sClient))
      $sWinTitle=WinGetTitle($aWin[$j])
      If StringRegExp($sWinTitle,$sTitleRegEx,0) Then
        WinSetTitle($aWin[$j],"",StringRegExpReplace($sWinTitle,$sTitleRegEx,StringFormat("\1 - %s - ctSpaces",$sClient)))
      EndIf
      If $bIcon Then
        _SendMessage($aWin[$j],$WM_SETICON,0,$hIcon)
        _SendMessage($aWin[$j],$WM_SETICON,1,$hImage)
      EndIf
    Next
    ;ConsoleWrite(TimerDiff($hTimer)&@CRLF)
    Sleep($iWinPoll)
  WEnd
  ; Cleanup
  _WinAPI_DestroyIcon($hIcon)
  _WinAPI_DeleteObject($hImage)
  $aSearch=_FileListToArrayRec($sData,"*",1,1)
  If $sData=$g_sDataDir&'\Default' Then
    For $j=1 To $aSearch[0]
        If Not _flt($aKeepDefault,$aSearch[$j]) Then _icl($aSearch[$j])
    Next
  Else
    For $j=1 To $aSearch[0]
        If Not _flt($aKeepActive,$aSearch[$j]) And Not _flt($aKeepActive,$aSearch[$j])  Then _icl($aSearch[$j])
    Next
  EndIf
  ;_DebugArrayDisplay($aFiles)
  For $j=1 To $aFiles[0]
    FileDelete($sData&"\"&$aFiles[$j])
  Next
  _delEmpty($sData)
  _UnsetLock($sLock)
  _WinAPI_EmptyWorkingSet()
EndFunc

Func _extDef($sData)
  $g_sErrFunc="_extDef"
  If Not DirCreate($sData) Then
    _Log("DirCreate")
  EndIf
  Local $sCmd='"'&$g_sDataDir&'\7za.exe" x "'&$g_sDataDir&'\Default.7z" -y -o"'&$sData&'"'
  Local $iRet=RunWait($sCmd,$g_sDataDir)
  If $iRet<>0 Then
    _Log(StringFormat("CmdRet:%d;Cmd:%s",$iRet,$sCmd))
    Return SetError(2,0,0)
  EndIf
  ;If Not FileExists(StringFormat("%s\ctSpaces",$sData)) Then Return SetError(3,0,0)
  FileClose(FileOpen(StringFormat("%s\ctSpaces",$sData),10))
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
  $folder_list=_FileListToArray($dir, '*', 2)
  If @error <> 4 Then
    For $i=1 To $folder_list[0]
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
  Local $sExec=$g_sDataDir&"\ctSpaces.exe"
  If $bStartup Or $bDesktop Or $bStartMenu Then
    FileCopy(@AutoItExe,$g_sDataDir&"\ctSpaces.exe",9)
  EndIf
  Local $sCmd=$g_sDataDir&"\ctSpaces.exe"
  Local $sIcon=$sCmd
  If $bA3x Then
    $sCmd=StringFormat('%s\ctSpaces.exe "%s\ctSpaces.a3x"',$g_sDataDir,$g_sDataDir)
    FileInstall("Res\ctdkgrsq.ico","ctdkgrsq.ico",1)
    $sIcon=$g_sDataDir&"\ctdkgrsq.ico"
  EndIf
  If $bStartup Then
    RegWrite("HKCU\Software\Microsoft\Windows\CurrentVersion\Run","ctSpaces","REG_SZ",$sCmd)
  EndIf
  If $bDesktop Then
    FileCreateShortcut($sCmd,@DesktopDir&"\ctSpaces.lnk",$g_sDataDir,"","",$sIcon)
  EndIf
  If $bStartMenu Then
    FileCreateShortcut($sCmd,@ProgramsDir&"\ctSpaces.lnk",$g_sDataDir,"","",$sIcon)
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
    _Log(StringFormat("FileCopy(%s,%s,1)",@AutoItExe,$sInstFullPath))
    If Not FileCopy(@AutoItExe,$sInstFullPath,1) Then
      MsgBox(16,$sTitle,"Error: unable to copy the update.")
      Return 1
    EndIf
    If _VersionCompare(FileGetVersion($sInstFullPath,"FileVersion"),$sCurVer)<>0 Then
      MsgBox(16,$sTitle,"Error: the update was copied but the version does not match ours.")
      Return 1
    EndIf
  EndIf
  _extDef7z(1)
  Local $iRet=MsgBox(32+4,$sTitle,"Would you like to run the new version after the update?")
  If $iRet=6 Then $bRun=1
  If $bRun Then
    Run($sInstFullPath&" ~!PostInstall",$g_sDataDir,@SW_SHOW)
  EndIf
  Return 1
EndFunc

Func _png2ico($sSrc,$sOut)
  Local $aSizes[]=[16,24,32,48,64,128,256]
  _GDIPlus_Startup()
  Local $hImage=_GDIPlus_ImageLoadFromFile($sSrc)
  If @error Then Return SetError((@Error*10)+0,(@Extended*10)+0,0)
  Local $aSize=_GDIPlus_ImageGetDimension($hImage)
  If @error Then Return SetError((@Error*10)+1,(@Extended*10)+1,0)
  Local $iSize=_Max($aSize[0],$aSize[1])
  Local $hScaled=_GDIPlus_ImageResize($hImage,$iSize,$iSize,4)
  If @error Then Return SetError((@Error*10)+2,(@Extended*10)+2,0)
  _GDIPlus_ImageDispose($hImage)
  Local $iMax=0,$hIcoImage
  For $i=0 To UBound($aSizes)-1
      If $aSizes[$i]>=$iSize Then ContinueLoop
      $hIcoImage=_GDIPlus_ImageResize($hScaled,$aSizes[$i],$aSizes[$i],4)
      If @error Then Return SetError((@Error*10)+3,(@Extended*10)+3,0)
      $aSizes[$i]=_GDIPlus_HICONCreateFromBitmap($hIcoImage)
      If @error Then Return SetError((@Error*10)+4,(@Extended*10)+4,0)
      _GDIPlus_ImageDispose($hIcoImage)
      $iMax+=1
  Next
  _GDIPlus_ImageDispose($hScaled)
  ReDim $aSizes[$iMax]
  _WinAPI_SaveHICONToFile($sOut,$aSizes,1,1)
  _GDIPlus_Shutdown()
  Return SetError(0,0,1)
EndFunc

Func _GdiGetTypes()
  _GDIPlus_Startup()
  Local $iMax,$vExts,$aFmts[]=[0]
  Local $g_aCodex=_GDIPlus_Decoders()
  If @error Then Return SetError(1,0,0)
  For $i=1 To $g_aCodex[0][0]
    $vExts=$g_aCodex[$i][6]
    If $vExts="" Then ContinueLoop
    $vExts=StringRegExp($vExts,"\*\.([^;]+);?",3)
    If @error Then ContinueLoop
    For $j=0 To UBound($vExts,1)-1
      For $k=1 To $aFmts[0]
        If $vExts[$j]=$aFmts[$k] Then ContinueLoop 2
      Next
      If $vExts[$j]="ico" Then ContinueLoop
      $iMax=UBound($aFmts,1)
      ReDim $aFmts[$iMax+1]
      $aFmts[$iMax]=StringLower($vExts[$j])
    Next
    $aFmts[0]=$iMax
  Next
  _GDIPlus_Shutdown()
  Return SetError(0,0,$aFmts)
EndFunc

Func _hasFmt(ByRef $aArray,$v)
    For $i=1 To $aArray[0]
      If $aArray[$i]=$v Then Return 1
    Next
    Return 0
EndFunc

Func _SetLock($sPath)
  Local $hFile
  If FileExists($sPath) Then
    $hFile=FileOpen($sPath)
    If $hFile=-1 Then Return SetError(1,0,0)
    $vData=FileRead($sPath)
    If @error Then Return SetError(2,0,0)
    FileClose($hFile)
    ;_Log($vData)
    If Not StringIsInt($vData) Then Return SetError(3,0,0)
    $vData=Int($vData)
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

Func _WinGetHandleByPID($vProc,$nVisible=1)
  $vProc=ProcessExists($vProc)
  If Not $vProc Then Return SetError(1,0,0)
  Local $aWL=WinList("[CLASS:Chrome_WidgetWin_1]")
  If $aWL[0][0]=0 Then Return SetError(1,0,0)
  Local $iMax=1,$aTemp[$aWL[0][0]+1]
  For $i=1 To $aWL[0][0]
      If $vProc=WinGetProcess($aWL[$i][1]) Then
        $aTemp[$iMax]=$aWL[$i][1]
        $iMax+=1
      EndIf
  Next
  $aTemp[0]=$iMax
  Return $aTemp
EndFunc

Func _SanitizeName($sName)
  $sName=StringStripWS($sName,7)
  $sName=StringRegExpReplace($sName,'[\\/:*?"<>|]',"")
  $sName=StringRegExpReplace($sName,'[\. ]+$','')
  If $sName='' Or StringRegExp($sName, '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$') Then Return SetError(1,0,0)
  Return $sName
EndFunc

Func _extDef7z($bO=0)
  Return FileInstall("Default.7z",$g_sDataDir&"\Default.7z",$bO)
EndFunc

Func _cmdGetOut($sCmd)
  Local $iPid,$vPeek,$vStdout
  $iPid=Run($sCmd,"",@SW_HIDE,0x2)
  While Sleep(1)
    $vPeek=StdoutRead($iPid,1,0)
    If @error Then ExitLoop
    If $vPeek<>"" Then $vStdOut&=StdoutRead($iPid)
  WEnd
  Return StringStripWS($vStdOut,7)
EndFunc
