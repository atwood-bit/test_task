#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <IE.au3>

parseIE()

Func parseIE()
   Global $url = ClipGet();
Local $oIE = _IECreate($url,0,0,100,100)
Local $oLis = _IETagNameGetCollection($oIE, "span")
For $oLi In $oLis
 If StringInStr ($oLi.outerhtml,'class="pp_last_activity_text"')>0 Then
  Local $oLiHtml=_IEPropertyGet($oLi,"innerhtml")
  Global $Num=$oLiHtml
 EndIf
next
_IEQuit($oIE)
ClipPut($Num)
WinActivate("MainWindow")
   EndFunc