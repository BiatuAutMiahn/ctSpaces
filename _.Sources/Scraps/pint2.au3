#include <WindowsConstants.au3>
#include <SendMessage.au3>
#include <WinAPIProc.au3>
#include <WinAPIRes.au3>
#include <WinAPIIcons.au3>
#include <WinAPIGdi.au3>
#include <Math.au3>
#include <Array.au3>
#Include "Icons.au3"
#Include "IconImage.au3"
#include "AppUserModelId.au3"
#include <GDIPlus.au3>

#include <GDIPlus.au3>
#include <GuiConstantsEx.au3>

Opt("MustDeclareVars", 1)

; ===============================================================================================================================
; Description ...: Shows how to display a PNG image
; Author ........: Paul Campbell (PaulIA)
; Notes .........:
; ===============================================================================================================================

; ===============================================================================================================================
; Global variables
; ===============================================================================================================================
Global $hGUI, $hImage, $hGraphic

; Create GUI
$hGUI = GUICreate("Show PNG", 240, 240)
GUISetState()

; Load PNG image
_GDIPlus_StartUp()
$hImage   = _GDIPlus_ImageLoadFromFile(@ScriptDir & "\Source.png")

; Draw PNG image
$hGraphic = _GDIPlus_GraphicsCreateFromHWND($hGUI)
_GDIPlus_GraphicsDrawImage($hGraphic, $hImage, 0, 0)

; Loop until user exits
do
until GUIGetMsg() = $GUI_EVENT_CLOSE

; Clean up resources
_GDIPlus_GraphicsDispose($hGraphic)
_GDIPlus_ImageDispose($hImage)
_GDIPlus_ShutDown()

;~ _GDIPlus_Startup()
;~ _ImageAutoCrop(@ScriptDir & "\source.png", @ScriptDir & "\destination.png")

;~ Func _ImageAutoCrop($sSrcImage, $sDstImage = "")
;~     If Not FileExists($sSrcImage) Then Return SetError(1, 0, 0)
;~     If $sDstImage = "" Then $sDstImage = $sSrcImage
;~     Local $hImage = _GDIPlus_ImageLoadFromFile($sSrcImage)
;~     Local $CoordsY = _GetImageCropSize($hImage)
;~     If @error Then Return SetError(2, 0, 0)
;~     Local $hClone = _GDIPlus_ImageClone($hImage)
;~     _GDIPlus_ImageRotateFlip($hClone, 1)
;~     Local $CoordsX = _GetImageCropSize($hClone)
;~     _GDIPlus_ImageDispose($hClone)
;~     Local $hCrop = _GDIPlus_BitmapCloneArea($hImage, $CoordsX[0], $CoordsY[0], $CoordsX[1], $CoordsY[1])
;~     _GDIPlus_ImageDispose($hImage)
;~     Local $bRes = _GDIPlus_ImageSaveToFile($hCrop, $sDstImage)
;~     _GDIPlus_ImageDispose($hCrop)
;~     Return $bRes
;~ EndFunc

;~ Func _GetImageCropSize($hImage)
;~     Local $iWidth = _GDIPlus_ImageGetWidth($hImage)
;~     Local $iHeight = _GDIPlus_ImageGetHeight($hImage)
;~     Local $hClone = _GDIPlus_ImageClone($hImage)
;~     Local $tBitmapData  = _GDIPlus_BitmapLockBits($hClone, 0, 0, $iWidth, $iHeight)
;~     Local $iScan0 = DllStructGetData($tBitmapData , "Scan0")
;~     Local $v_BufferA = DllStructCreate("byte[" & $iHeight * $iWidth * 4 & "]", $iScan0) ; Create DLL structure for all pixels
;~     Local $AllPixels = StringRegExpReplace(DllStructGetData($v_BufferA, 1), "^0x", "")
;~     Local $aPixels = StringRegExp($AllPixels, ".{" & 8 * $iWidth & "}", 3)
;~     Local $iYStart, $iYEnd

;~     For $i = 0 To UBound($aPixels) - 1
;~         If Not $iYStart And StringRegExp($aPixels[$i], "[^F]") Then
;~             $iYStart = $i
;~             ExitLoop
;~         EndIf
;~     Next
;~     If Not $iYStart Then SetError(1, 0, 0)

;~     For $i = UBound($aPixels) - 1 To 0 Step -1
;~         If Not $iYEnd And StringRegExp($aPixels[$i], "[^F]") Then
;~             $iYEnd = $i
;~             ExitLoop
;~         EndIf
;~     Next

;~     _GDIPlus_BitmapUnlockBits($hClone, $tBitmapData ) ; releases the locked region
;~     _GDIPlus_ImageDispose($hClone)

;~     Local $aReturn = [$iYStart, $iYEnd - $iYStart]
;~     Return $aReturn
;~ EndFunc

Func _png2ico($sSrc,$sOut)
  Local $aSizes[]=[16,24,32,48,64,128,256]
  ;Local $hSrc=_IconImage_FromImageFile($sSrc)
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
  If FileExists($sOut) Then FileDelete($sOut)
  _WinAPI_SaveHICONToFile($sOut,$aSizes,1,1)
  _GDIPlus_Shutdown()
  Return SetError(0,0,1)
EndFunc
