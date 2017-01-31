#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Christian Chroszewski

 Script Function:
	Will remove all users accepted needed administrativ accounts from
	the local Administrators Group

#ce ----------------------------------------------------------------------------
#RequireAdmin
; Script Start - Add your code below here
#Include <Array.au3>
$Members = _LocalGroupMembers("Administrators")
;_ArrayDisplay($Members)

For $element IN $Members
	  if $element = "Administrator" Or $element = "otoservice" Or $element = "td\Domain Admins" Or $element = "OBE GG wks_Admins" Then
		 ContinueLoop
	  Else
		 _removeuser($element)
	  EndIf

Next
Func _LocalGroupMembers($sGroup)
	;funkey 03.12.2009
	Local $line, $aMembers
	Local $cmd = "net localgroup "& $sGroup
	Local $Pid = Run(@ComSpec & " /c " & $cmd, "", @SW_HIDE, 2)
	While 1
		$line &= StdoutRead($Pid)
		If @error Then ExitLoop
	Wend
	$aMembers = StringSplit(StringTrimLeft($line, StringInStr($line, "-----", 0, -1) + 6), @CRLF, 3)
	ReDim $aMembers[UBound($aMembers) - 3]
	Return $aMembers
 EndFunc

 Func _removeuser($sUser)
	  Local $cmd = "net localgroup Administrators "&$sUser&" /delete"
	  Run(@ComSpec & " /c " & $cmd, "", @SW_HIDE, 2)
 EndFunc