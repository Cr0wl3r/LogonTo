#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         Christian Chroszewski

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
MsgBox(0,"", @ComputerName)
MsgBox(0,"", @UserName)

#cs -----------------------------------------------------------------------------------------------
   Variables to connect to OT Asset DB
#ce -----------------------------------------------------------------------------------------------
$DBServer="OBESR098\SAFEEXPERT"
$DB="ICTAssetDB"
$DBUser="Nerd"
$DBPW="start123"

$checksrv = Ping ("OBESR098", 250)
MsgBox(0,"",$checksrv )
if $checksrv > 0 Then
   CheckNerd()
Else
   MsgBox(0,"","Keine Verbindung zu DB Server" )
EndIf

; ---------------------------------------------------------------------------------------------------
;   Connect to Database
; ---------------------------------------------------------------------------------------------------
Func CheckNerd()
   $constrim="DRIVER={SQL Server};SERVER="&$DBServer&";DATABASE="&$DB&";uid="&$DBUser&";pwd="&$DBPW&";"
   $adCN = ObjCreate ("ADODB.Connection") ; <== Create SQL connection
   $adCN.Open ($constrim) ; <== Connect with required credentials


   if @error Then
	  MsgBox(0,"",$constrim )
	  MsgBox(0, "ERROR", "Failed to connect to the database")
	  Exit
   Else
	  MsgBox(0, "Success!", "Connection to database successful!")
   EndIf

   ; Query which will check the user who is member of the computer in Nerd database
   $sQuery = "select u.UserAccount from tAsset as a INNER JOIN tUser as u ON a.AssetUser = u.UserID where AssetOBE = 'OBELTF73'"

   $result = $adCN.Execute($sQuery)
   MsgBox(0, "", $result.Fields( "UserAccount" ).Value)
   $adCN.Close ; ==> Close the database
EndFunc
