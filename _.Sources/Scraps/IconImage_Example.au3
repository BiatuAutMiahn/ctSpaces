; =============================================================================
;  AutoIt IconImage UDF Example (2011.9.21)
;  Purpose: Icon Management And Conversion (Among IconImage/hIcon/.ico)
;  Author: Ward
;  Required:
;	"Icons.au3"			by Yashied
; =============================================================================

#Include <GUIConstantsEx.au3>
#Include <EditConstants.au3>
#Include <StaticConstants.au3>
#Include <ButtonConstants.au3>
#Include <GuiImageList.au3>

#Include "Icons.au3"
#Include "IconImage.au3"


_IconImage_Startup()

Example()

_IconImage_Shutdown()


Func Example()
	Local $RequestSize = 256, $DisplaySize = 256

	; Create window
	Opt("GUIResizeMode", $GUI_DOCKALL)
	Local $MainWin = GUICreate("IconImage UDF Example", 350, 280, -1, -1, BitOR($WS_SYSMENU, $DS_MODALFRAME, $WS_CAPTION), BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW))
	Local $ComboFile = GUICtrlCreateCombo(@AutoItExe, 10, 10, 300, 30)
	Local $ButtonFile = GUICtrlCreateButton("...", 320, 10, 30, 27)
	Local $ButtonRotate = GUICtrlCreateButton("", 360, 10, 40, 27)

	Local $InputType = GUICtrlCreateInput("", 410, 10, 90, 30, $ES_READONLY)
	GUICtrlSetColor(-1, 0x000080)
	GUICtrlSetFont(-1, 8.5, 900)

	GUICtrlCreateLabel("Request Size:", 10, 53)
	Local $InputRequest = GUICtrlCreateInput($RequestSize, 120, 50, 90, 30)
	GUICtrlCreateUpdown(-1)
	GUICtrlSetLimit(-1, 512, 16)

	GUICtrlCreateLabel("Display Size:", 10, 93)
	Local $InputDisplay = GUICtrlCreateInput($DisplaySize, 120, 90, 90, 30)
	GUICtrlCreateUpdown(-1)
	GUICtrlSetLimit(-1, 512, 16)

	GUICtrlCreateLabel("Return Size:", 10, 133)
	Local $InputReturn = GUICtrlCreateInput("", 120, 130, 90, 30, $ES_READONLY)

	Local $ButtonPng = GUICtrlCreateButton("Save to .png file", 10, 180, 200, 30)
	Local $ButtonIco = GUICtrlCreateButton("Save to .ico file", 10, 220, 200, 30)
	Local $ButtonIcoVista = GUICtrlCreateButton("Save to .ico file for vista", 10, 260, 200, 30)
	Local $ButtonIcoMultiple = GUICtrlCreateButton("Save to .ico file with 6 size", 10, 300, 200, 30)

	Local $IconBorder = GUICtrlCreateLabel("", 239, 59, $DisplaySize + 2, $DisplaySize + 2, $SS_SUNKEN)
	Local $IconControl = GUICtrlCreateIcon("", "", 240, 60, $DisplaySize, $DisplaySize)

	; Create image button
	Local $PngData = Binary("0x89504E470D0A1A0A0000000D4948445200000020000000200806000000737A7AF40000000774494D4507DB09110A1E38A04C16500000001774455874536F66747761726500474C44504E472076657220332E347185A4E10000000874704E47474C4433000000004A80291F0000000467414D410000B18F0BFC61050000052A49444154789CBD567D4C535714EFD22C8430C287422163306C4A81D282514216C1022D6C6612E66620CCB1691808A271CB46661495F24D1C9B732A31C010D02963080A2DD00D418110686636704ED8F816F9DCC8B21092FD7376CEE33EF3A6684B68FCE3E4BE7BDFBBE7F73B9FEF8800E085678948247A519363B8AECD6D057CB631F7FD5A45640101FBC89C1628ED79006A5DB3D5493C41606B4AB193FA64539D3ACBF00F01AA752DB0EDB81EAA4CD350F8E32844645B97C42302A8541C96D9509456D603E5B74740FFEB3CB4DCFF0BAA06C6A0B27F0CAE0F2C4074C91D38D339695512A247E0C7EA4EE5D7F78301810BDB2621A57610526B87606FCD5DD88712FB4D3FBC5D3100BB2AEEC285DE39B0564E7004646F24BBA79675A3D50B90DE300C719706E1DD2BBFC39EAB247FACACDF0EC19E2B439058330C9A150F2851ECD64D80B23CECE8F715058DF7E054E72C24D68E425CF92FB0F38B5BB0F3CBDB109D6F44400324D58F4152C304683121F14E18CA2B56F1006579B8AE79E96CD743C8304E43ECD91ED0EA9A9602767F5AE4BD3DEE90C7D61DFBA3F25AE150EB3C68F38C5605E709B844A24B4F9BFE8603F5231099D5B884676A14058A37ADE199759DAF1775581D9C2720D1A05BF3FB9621A1F427502564EAF0CC8BE24BE1417162F10E5E2B38DD0F0C0C6C0E0A0A025C4DB45F9540548E61519BDB02E46A55FC9144617211A052E15F1D18E007412A259032954A554E95C3579052A94CA2735EF8F7145EDAE3BA8556B158ECB01A01B270338A062504C59D674A4A5059255D96C964056E6E6E090A8582B308418BE93BDCEF62A0D3AEAEAEF1010101DDB4C7F364320ECF27990728B42EAB5601232141D980622B78C959804A7BF13910456E6363B38DB794BE47C553F4ECE8E8184BA17276768E61EF9799310A163E5A9D56ED034F13B4584ECAE472790923684BF9212020133C4B8930AD8233BAE3409633239FCC013349E4CE9491FBB8F86128DE1100288560EC8EE4F133B39DF019045CA452A98E5788EE1E1026DBF32040A528DBB87163BCAFAF6F254A2986A39A25DDECF320206624A8FEFD2811FDFDFDF5A4DCCBCBEB30252511A1BD8787075592183D16C508CEAF9B00095A6C87251581C0C91882595662B7C83394237C8810B0DDC7C7670BADB4974A37E956CBFA357B0081630439F0A7A7A76706EB8C4ECC3B723E2C9C60B3F2D9E4554DE7C292E63B6364B69E1BEFB438E6712DC082106C60EE0F61F52CE34B8ABD7760EF8305DFF8B173B1505FC8C1F311A9157D90D736C1CF13F696CC84B68C848401DB0A15F32D97D53A7DE3AE2D68A35F78A3B0EEA9A5871EBDD65ED63D01278DE3A059F9ADBB9825B0162190F01C235CFD790EC7360EC021E2444388FA58ED07EA2CFD78DA451334FDB6080997EED150B3CC11B62638CD8A85371F4079DF0C44E4B74154513B249CEB82F74ABAE144E320D40C2CC291D64978EBBC09423F2EAFB39A07089C668A9CB629D85D751F4A7B67E074D7347C85720EE7C712D302E476CC40FA8D49D85B3BC24D59CEDEAAD72CCA014BC0E9377EBC7D0EF6D58D7200EF7F370CA938C2A5D68FC287D72660FF8D29486B7A0829F5E3F066710704277FFE35DE7BD96C155848804A514973E3C1E6590442B0C62938A09F8674C30C7C645C80349C25E34BEF4074C10F109276E632AB123BB37DC0520FB04E194624326E2EC2E1564CC2AC26ACF5664E34D9FAE5D04F2E76BE24793586FF2DF39564B51CE04944EA0C90D5F72FD08485FB1DA295418766CCCDEC1BFBFF95B135AB80271185E0DB3FBBDCC99A9684F508FB35CF03EB20C14F40F666EF5893808084C3D32C7E5CFE037886BAD0B3E83D2A0000000049454E44AE426082")
	Local $hIcon = _IconImage_ToHandle(_IconImage_FromImageData($PngData), 32, 32)
	Local $ImageList = _GUIImageList_Create(32, 32, 5)
	_GUIImageList_ReplaceIcon($ImageList, -1, $hIcon)
	_WinAPI_DestroyIcon($hIcon)

    Local $ButtonIL = DllStructCreate("handle;int[4];uint")
    DllStructSetData($ButtonIL, 1, $ImageList)
    DllStructSetData($ButtonIL, 2, 2, 1)
    DllStructSetData($ButtonIL, 2, 2, 2)
    DllStructSetData($ButtonIL, 2, 2, 3)
    DllStructSetData($ButtonIL, 2, 2, 4)
    DllStructSetData($ButtonIL, 3, $BUTTON_IMAGELIST_ALIGN_LEFT)
    GUICtrlSendMsg($ButtonRotate, $BCM_SETIMAGELIST, 0, DllStructGetPtr($ButtonIL))

	; Create test path
	FileCreateShortcut(@AutoItExe, @TempDir & "\test.lnk")
	GUICtrlSetData($ComboFile, @TempDir & "\test.lnk")
	GUICtrlSetData($ComboFile, @ComSpec)
	GUICtrlSetData($ComboFile, "Example.???")
	GUICtrlSetData($ComboFile, "Example.txt")
	GUICtrlSetData($ComboFile, "Example.ini")
	GUICtrlSetData($ComboFile, "C:\")
	GUICtrlSetData($ComboFile, "D:\")
	GUICtrlSetData($ComboFile, @ScriptDir)
	GUICtrlSetData($ComboFile, @ScriptFullPath)
	GUICtrlSetData($ComboFile, @DesktopDir)
	GUICtrlSetData($ComboFile, @FavoritesDir)
	GUICtrlSetData($ComboFile, @MyDocumentsDir)

	; Main loop
	Local $NextAction = $InputRequest
	Local $Path = "", $IconImage, $DisplayImage, $Rotate = 0
	While 1
		If $NextAction = 0 Then $NextAction = GUIGetMsg()
		Switch $NextAction
		Case $InputDisplay
			$DisplaySize = Int(GUICtrlRead($InputDisplay))

			Local $Width = $DisplaySize + 274
			Local $Height = $DisplaySize + 107
			If $Width < 535 Then $Width = 535
			If $Height < 380 Then $Height = 380
			WinMove($MainWin, "", Default, Default, $Width, $Height)
			ControlMove($MainWin, "", GUICtrlGetHandle($IconControl), Default, Default, $DisplaySize, $DisplaySize)
			ControlMove($MainWin, "", GUICtrlGetHandle($IconBorder), Default, Default, $DisplaySize + 2, $DisplaySize + 2)

			$DisplayImage = _IconImage_Scale($IconImage, $DisplaySize, $DisplaySize, $GDIP_ModeHighQuality)
			Local $hIcon = _IconImage_ToHandle($DisplayImage, $DisplaySize, $DisplaySize)
			_SetHIcon($IconControl, $hIcon)
			_WinAPI_DestroyIcon($hIcon)
			GUISetState()

		Case $InputRequest, $ComboFile
			If $Path <> GUICtrlRead($ComboFile) Or $RequestSize <> Int(GUICtrlRead($InputRequest)) Then
				$Path = GUICtrlRead($ComboFile)
				$RequestSize = Int(GUICtrlRead($InputRequest))

				Switch _WinAPI_PathFindExtension($Path)
				Case ".png", ".jpg", ".bmp", ".tif", ".gif", ".tiff"
					$IconImage = _IconImage_FromImageFile($Path)
				Case Else
					$IconImage = _IconImage_FromPath($Path, $RequestSize)
				EndSwitch

				Local $ReturnSize = _IconImage_GetSize($IconImage)
				$ReturnSize = $ReturnSize[0] & "x" & $ReturnSize[1]
				GUICtrlSetData($InputReturn, $ReturnSize)

				If _IconImage_IsPNGIcon($IconImage) Then
					GUICtrlSetData($InputType, "PNG Icon")
				Else
					GUICtrlSetData($InputType, "BMP Icon")
				EndIf

				$NextAction = $InputDisplay
				ContinueLoop
			EndIf

		Case $ButtonRotate
			$IconImage = _IconImage_RotateFlip($IconImage, $GDIP_Rotate90FlipNone)
			$NextAction = $InputDisplay
			ContinueLoop

		Case $ButtonFile
			Local $File = FileOpenDialog("", @ScriptDir, "All (*.*)", 3, "", $MainWin)
			If Not @Error Then GUICtrlSetData($ComboFile, $File, $File)

		Case $ButtonPng
			Local $PngFile = FileSaveDialog("Save As .png File", @ScriptDir, "PNG files (*.png)", 0, "", $MainWin)
			If Not @Error Then
				If _WinAPI_PathFindExtension($PngFile) <> ".png" Then $PngFile &= ".png"
				_IconImage_ToImageFile($DisplayImage, $PngFile, "PNG")
				GUICtrlSetData($ComboFile, $PngFile, $PngFile)
			EndIf

		Case $ButtonIco
			Local $IcoFile = FileSaveDialog("Save As .ico File", @ScriptDir, "Icon files (*.ico)", 0, "", $MainWin)
			If Not @Error Then
				If _WinAPI_PathFindExtension($IcoFile) <> ".ico" Then $IcoFile &= ".ico"
				Local $NewImage = _IconImage_ConvertToBMPIcon($DisplayImage)
				If Not @Error Then
					_IconImage_ToIcoFile($NewImage, $IcoFile)
					GUICtrlSetData($ComboFile, $IcoFile, $IcoFile)
				EndIf
			EndIf

		Case $ButtonIcoVista
			Local $IcoFile = FileSaveDialog("Save As .ico File", @ScriptDir, "Icon files (*.ico)", 0, "", $MainWin)
			If Not @Error Then
				If _WinAPI_PathFindExtension($IcoFile) <> ".ico" Then $IcoFile &= ".ico"
				Local $NewImage = _IconImage_ConvertToPNGIcon($DisplayImage)
				If Not @Error Then
					_IconImage_ToIcoFile($NewImage, $IcoFile)
					GUICtrlSetData($ComboFile, $IcoFile, $IcoFile)
				EndIf
			EndIf

		Case $ButtonIcoMultiple
			Local $IcoFile = FileSaveDialog("Save As Multiple .ico File", @ScriptDir, "Icon files (*.ico)", 0, "", $MainWin)
			If Not @Error Then
				If _WinAPI_PathFindExtension($IcoFile) <> ".ico" Then $IcoFile &= ".ico"
				Local $NewImage[6] = [ _
					_IconImage_Scale($IconImage, 16, 16, $GDIP_ModeHighQuality),	_
					_IconImage_Scale($IconImage, 24, 24, $GDIP_ModeHighQuality),	_
					_IconImage_Scale($IconImage, 48, 48, $GDIP_ModeHighQuality),	_
					_IconImage_Scale($IconImage, 64, 64, $GDIP_ModeHighQuality),	_
					_IconImage_Scale($IconImage, 128, 128, $GDIP_ModeHighQuality),	_
					_IconImage_ConvertToPNGIcon(_IconImage_Scale($IconImage, 256, 256, $GDIP_ModeHighQuality))	_
				]

				_IconImage_ToIcoFile($NewImage, $IcoFile)
				GUICtrlSetData($ComboFile, $IcoFile, $IcoFile)
			EndIf

		Case $GUI_EVENT_CLOSE
			ExitLoop

		Case Else
			If $Path <> GUICtrlRead($ComboFile) Then
				$NextAction = $ComboFile
				ContinueLoop
			EndIf
		EndSwitch
		$NextAction = 0
	WEnd

	GUIDelete()
	FileDelete(@TempDir & "\test.lnk")
EndFunc
