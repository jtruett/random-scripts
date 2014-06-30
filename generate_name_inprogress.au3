#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=generate_name.exe
#AutoIt3Wrapper_Res_Fileversion=0.1.0.1
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <TRIMFunctions.au3>


#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=frmMain.kxf
$frmMain = GUICreate("Computer Name", 228, 209, 272, 192)
$txtName = GUICtrlCreateInput(BuildNewName(), 48, 104, 129, 21)
$lblHeading = GUICtrlCreateLabel("Generated Name", 48, 80, 85, 17)
$btnCopy = GUICtrlCreateButton("Copy to Clipboard", 48, 136, 131, 25)
$lblName = GUICtrlCreateLabel("Current Name", 48, 16, 69, 17)
$lblCurrName = GUICtrlCreateLabel(@ComputerName, 48, 48, 124, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$btnRename = GUICtrlCreateButton("Rename Computer", 48, 160, 131, 25)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

#cs ===============================================================================
 Function:      _RenameComputer( $iCompName , $iUserName = "" , $iPassword = "" )

 Description:   Renames the local computer

 Parameter(s):  $iCompName: The new computer name

                Required Only if PC is joined to Domain:
                    $iUserName: Username in DOMAIN\UserNamefFormat
                    $iPassword: Password of the specified account

 Returns:       1 - Succeeded (Reboot to take effect)

                0 - Invalid parameters
                    @error 2 - Computername contains invalid characters.
                    @error 3 - Current account does not have sufficient rights
                    @error 4 - Failed to create COM Object

                Returns error code returned by WMI
                    Sets @error 1

 Author(s):  Kenneth Morrissey (ken82m)
#ce ===============================================================================

Func _RenameComputer($iCompName, $iUserName = "", $iPassword = "")
    $Check = StringSplit($iCompName, "`~!@#$%^&*()=+_[]{}\|;:.'"",<>/? ")
    If $Check[0] > 1 Then
        SetError(2)
        Return 0
    EndIf
    If Not IsAdmin() Then
        SetError(3)
        Return 0
    EndIf

    $objWMIService = ObjGet("winmgmts:\root\cimv2")
    If @error Then
        SetError(4)
        Return 0
    EndIf

    For $objComputer In $objWMIService.InstancesOf("Win32_ComputerSystem")
        $oReturn = $objComputer.rename($iCompName,$iPassword,$iUserName)
        If $oReturn <> 0 Then
            SetError(1)
            Return $oReturn
        Else
            Return 1
        EndIf
    Next
EndFunc


Func BuildNewName()
	$objWMI = ObjGet("winmgmts:\\localhost\root\CIMV2")
	$objItems = $objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem", "WQL", 0x10 + 0x20)
	If IsObj($objItems) Then
		For $objItem In $objItems
			$Manufacturer = $objItem.Manufacturer
			$Model = $objItem.Model
		Next
	EndIf
	$objItems = $objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct", "WQL", 0x10 + 0x20)
	If IsObj($objItems) Then
		For $objItem In $objItems
			$Serial = $objItem.IdentifyingNumber
		Next
	EndIf

	$Serial = _LTRIM($Serial,'0') ;Trims leading 0's from serial#
	$Serial = StringStripWS($Serial,8) ;Strips whitespace

	Dim $NewName = ""
	Select
		Case StringInStr($Manufacturer,"Gateway",0)
			$NewName = "G"
		Case StringInStr($Manufacturer,"Fujitsu",0)
			$NewName = "F"
		Case StringInStr($Manufacturer,"Dell",0)
			$NewName = "D"
		Case StringInStr($Manufacturer,"Howard",0)
			$NewName = "H" ; Does Howard always set the name?  What about a system board replacement?
	EndSelect

	$NewName = $NewName & $Serial

	If $NewName == "" Then
		MsgBox(16,"Error","Serialized name could not be determined")
	EndIf

	Return $NewName
EndFunc

While 1
  $nMsg = GUIGetMsg()
  Switch $nMsg
	Case $GUI_EVENT_CLOSE
		Exit
	Case $btnCopy
		ClipPut(GUICtrlRead($txtName))
	Case $btnRename
		$rReturn = _RenameComputer(GUICtrlRead($txtName))
		If $rReturn == 1 Then
			$RestartMsg = MsgBox(49,"Restart required","Renamed computer to " & GUICtrlRead($txtName) & ". Restarting in 5 seconds",5)
			Switch $RestartMsg
				Case -1
					;message box timed out, restart
					Shutdown(6)
				Case 1
					;user clicked ok, restart
					Shutdown(6)
				Case Else
					;a 2 means they clicked the x or "cancel"
					Msgbox(0,"Restart Cancelled", "Restart was cancelled")
			EndSwitch
		Else
			MsgBox(0,"Problem","I think there was a problem. " & $rReturn)
		EndIf
  EndSwitch
WEnd



