; PNG-to-ICO using _WinAPI_SaveHICONToFile
;   • keeps alpha   • makes every common size ≤ the source bitmap


; ----------------------------------------------------------
Global Const $g_sSrcPNG = @ScriptDir & "\source.png"
Global Const $g_sOutICO = @ScriptDir & "\result.ico"

