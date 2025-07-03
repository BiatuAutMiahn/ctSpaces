$sIcoReq="C:\Temp\test.ico"
$iExt=StringInStr($sIcoReq,".",0,-1)
Local $sExt=StringTrimLeft($sIcoReq,$iExt)
MsgBox(64,"",$sExt)
