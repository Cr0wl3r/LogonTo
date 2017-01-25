#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Christian Chroszewski

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
;MsgBox(0,"", @ComputerName)
;MsgBox(0,"", @UserName)
;#RequireAdmin
#cs -----------------------------------------------------------------------------------------------
   Variables to connect to OT Asset DB
#ce -----------------------------------------------------------------------------------------------
$DBServer="OBESR098"
$DBInstance="SAFEEXPERT"
$DB="ICTAssetDB"
$DBUser="Nerd"
$DBPW="start123"

;Check if db server is available
IsUserAdmin()
$checksrv = Ping ($DBServer, 250)
;MsgBox(0,"",$checksrv )
if $checksrv > 0 Then
   CheckNerd()
Else
   MsgBox(0,"","Keine Verbindung zu DB Server" )
   CheckOffline()
EndIf




Func CheckNerd()
   ; ---------------------------------------------------------------------------------------------------
   ;   Connect to Database
   ; ---------------------------------------------------------------------------------------------------
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
   $sQuery = "select u.UserAccount from tAsset as a INNER JOIN tUser as u ON a.AssetUser = u.UserID where AssetOBE = 'OBELTF73'"

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
	  MsgBox(0,"","You are not the Owner of this Computer" )
   Else
	  MsgBox(0,"","You are the Owner of this Computer" )
   EndIf
EndFunc

Func IsUserAdmin()
   if IsAdmin() Then Exit
EndFunc
