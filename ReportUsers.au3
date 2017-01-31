#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         myName

 Script Function:
	Template AutoIt script.

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

;Check if db server is available
$checksrv = Ping ($DBServer, 250)
if $checksrv > 0 Then
   Report()
Else
   Exit
EndIf

Func Report()
   ; ---------------------------------------------------------------------------------------------------
   ;   Connect to Database
   ; ---------------------------------------------------------------------------------------------------
   Local $oErrorHandler = ObjEvent("AutoIt.Error", "_ErrFunc")
   $constrim="DRIVER={SQL Server};SERVER="&$DBServer&"\"&$DBInstance&";DATABASE="&$DB&";uid="&$DBUser&";pwd="&$DBPW&";"
   $adCN = ObjCreate ("ADODB.Connection") ; <== Create SQL connection
   $adCN.Open ($constrim) ; <== Connect with required credentials


   if @error Then
	  ;MsgBox(0,"",$constrim )
	  ;MsgBox(0, "ERROR", "Failed to connect to the database")
	  Exit
   Else
	  ;MsgBox(0, "Success!", "Connection to database successful!")
   EndIf

   If IsAdmin() Then
	  $adm=1
   Else
	  $adm=0
   EndIf

   ; Query which will check the user who is member of the computer in Nerd database
   $sQuery = "INSERT INTO tClientReport (AssetID, LoggedOnUser, IsAdmin, Date) Values ((select AssetID from tAsset where AssetOBE = '"&@ComputerName&"'),(select UserID from tUser where UserAccount = '"&@UserName&"'),"&$adm&",CURRENT_TIMESTAMP)"
   ConsoleWrite($sQuery)
   ConsoleWrite(@CRLF)

   $result = $adCN.Execute($sQuery)
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