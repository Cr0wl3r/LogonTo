#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Christian Chroszewski
			     IT-Engineer ICT Germany
 Version:		 1.170127

 Script Function:
	This Script is checking the NERD database for the owner of the computer
	where the script is running. If the user who is currently logon To
	the computer is not the owner, the user will automatically logged off
	because he is not authorized to use the computer. To work offline Then
	script will create a registry key in HKCU if it was checked once. If
	this key does not exist user will be logged off.
	If a local Admin is detected the script will exit.

#ce ----------------------------------------------------------------------------
#include <MsgBoxConstants.au3>
#cs -----------------------------------------------------------------------------------------------
   Variables to connect to OT Asset DB
#ce -----------------------------------------------------------------------------------------------
$DBServer="OBESR098"
$DBInstance="SAFEEXPERT"
$DB="ICTAssetDB"
$DBUser="Nerd"
$DBPW="start123"

;Check if user is admin
IsUserAdmin()
;Check if db server is available
$checksrv = Ping ($DBServer, 250)
;MsgBox(0,"",$checksrv )
if $checksrv > 0 Then
   CheckNerd()
Else
   ;MsgBox(0,"","Keine Verbindung zu DB Server" )
   CheckOffline()
EndIf




Func CheckNerd()
   ; ---------------------------------------------------------------------------------------------------
   ;   Connect to Database
   ; ---------------------------------------------------------------------------------------------------
   Local $oErrorHandler = ObjEvent("AutoIt.Error", "_ErrFunc")
   $constrim="DRIVER={SQL Server};SERVER="&$DBServer&"\"&$DBInstance&";DATABASE="&$DB&";uid="&$DBUser&";pwd="&$DBPW&";"
   $adCN = ObjCreate ("ADODB.Connection") ; <== Create SQL connection
   $adCN.Open ($constrim) ; <== Connect with required credentials


   if @error Then
	  ;MsgBox(0,"",$constrim )
	  MsgBox(0, "ERROR", "Failed to connect to the database")
	  Exit
   Else
	  ;MsgBox(0, "Success!", "Connection to database successful!")
   EndIf

   ; Query which will check the user who is member of the computer in Nerd database
   $sQuery = "select u.UserAccount from tAsset as a INNER JOIN tUser as u ON a.AssetUser = u.UserID where AssetOBE = '"& @ComputerName &"'"
   ConsoleWrite($sQuery)
   ConsoleWrite(@CRLF)

   $result = $adCN.Execute($sQuery)
   $DBPCOwner = $result.Fields( "UserAccount" ).Value
   ;MsgBox(0, "", $result.Fields( "UserAccount" ).Value)
   ;Write Registry Key to check offline
   RegWrite("HKEY_CURRENT_USER\Outotec", "OTLogon", "Reg_SZ", $DBPCOwner)
   $adCN.Close ; ==> Close the database
   LogonApproval($DBPCOwner)
EndFunc

Func CheckOffline()
   ;Check who is the owner of the PC in Registry
   $RegPCOwner = RegRead("HKEY_CURRENT_USER\Outotec", "OTLogon")
   LogonApproval($RegPCOwner)
EndFunc

Func LogonApproval($usr)
   If $usr <> @UserName Then
	  MsgBox($MB_ICONERROR,"","You are not authorized to use this Computer"&@CRLF&@CRLF&"Computername: "&@ComputerName&@CRLF&"Nerduser: "&$usr&@CRLF& "Current user: " & @UserName & @CRLF& @CRLF &"You will be logged off automatically in 10 seconds!", 10)
	  Shutdown(0)
   Else
	  MsgBox($MB_ICONINFORMATION,"","You are the Owner of this Computer"&@CRLF&"Computername: "&@ComputerName&@CRLF&"Nerduser: "&$usr& @CRLF & "Current user: " & @UserName, 10 )
   EndIf
EndFunc

Func IsUserAdmin()
   if IsAdmin() Then Exit
EndFunc

Func _ErrFunc($oError)
    ; This is an error handler like: https://www.autoitscript.com/autoit3/docs/functions/ObjEvent.htm
    ConsoleWrite(@ScriptName & " (" & $oError.scriptline & ") : ==> COM Error intercepted !" & @CRLF & _
            @TAB & "err.number is: " & @TAB & @TAB & "0x" & Hex($oError.number) & @CRLF & _
            @TAB & "err.windescription:" & @TAB & $oError.windescription & @CRLF & _
            @TAB & "err.description is: " & @TAB & $oError.description & @CRLF & _
            @TAB & "err.source is: " & @TAB & @TAB & $oError.source & @CRLF & _
            @TAB & "err.helpfile is: " & @TAB & $oError.helpfile & @CRLF & _
            @TAB & "err.helpcontext is: " & @TAB & $oError.helpcontext & @CRLF & _
            @TAB & "err.lastdllerror is: " & @TAB & $oError.lastdllerror & @CRLF & _
            @TAB & "err.scriptline is: " & @TAB & $oError.scriptline & @CRLF & _
            @TAB & "err.retcode is: " & @TAB & "0x" & Hex($oError.retcode) & @CRLF & @CRLF)
EndFunc   ;==>_ErrFunc