#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

' This automatically generated script may need to be
' edited in order to work correctly.
iosConfig = InputBox(_
	"Please specify the filename:", _
	"Target File?", _
	"switch333.txt")

	
Sub Main
crt.Screen.Send chr(13)
		
		
	Do
		Dim nIndex
		nIndex = crt.screen.WaitForStrings( _
			"Would you like to terminate autoinstall?", _
			"Would you like to enter the initial configuration dialog?", _
			"Press RETURN to get started!", _
			"Username:", _
			"Password:", _
			"Switch>", _
			"yourname>", _
			"Switch#", _
			"yourname#", _
			"Router>", _
			"Router#")

		Select Case nIndex
			Case 1
				crt.Screen.Send "y" & chr(13)

			Case 2
				
				crt.Screen.Send "n" & chr(13)
			Case 3

				crt.Screen.Send chr(13)

			Case 4
				crt.Screen.Send "cisco" & chr(13)

			Case 5

				crt.Screen.Send "cisco" & chr(13)

			Case 6

				crt.Screen.Send "en" & chr(13)

			Case 7

				crt.Screen.Send "en" & chr(13)


			Case 8

				crt.Screen.Send chr(13)
				
				Exit Do
				
			Case 9

				crt.Screen.Send chr(13)	
				
				Exit Do

			Case 10

				crt.Screen.Send "en" & chr(13)

			Case 11

				crt.Screen.Send chr(13)
				
				Exit Do
				
		Case Else
			
			crt.Screen.Send chr(13)
	
			'crt.Screen.Send chr(13)
	
			'Exit Do

		End Select	
		
		
	Loop 

	crt.Screen.WaitForString "Switch#"
 	crt.Screen.Send "erase flash:" & chr(13)
 	crt.Screen.Send chr(13)
 	crt.Screen.Send chr(13)
 	crt.Screen.Send chr(13)
 	crt.Screen.WaitForString "Erase of flash: complete"
 	crt.Screen.Send chr(13)
	crt.Screen.Send "conf t" & chr(13)
	crt.Screen.Send "int vlan1" & chr(13)
	crt.Screen.Send "ip address dhcp" & chr(13)
	crt.Screen.Send "no shut" & chr(13)
	crt.Screen.Send "end" & chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "mask 255.255.255.0"
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "Switch#"
	crt.Screen.Send "copy tftp flash:" & chr(13)
	crt.Screen.Send "10.201.2.8" & chr(13)
	crt.Screen.Send "c2960x-universalk9-mz.150-2.EX4.bin" & chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "Source filename []? c2960x-universalk9-mz.150-2.EX4.bin"
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "Switch#"
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "Switch#"
	crt.Screen.Send "copy tftp flash:" & chr(13)
	crt.Screen.Send "10.201.2.8" & chr(13)
	crt.Screen.Send "Projects/" & iosConfig & chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "bytes copied"
	crt.Screen.Send "copy flash:" & iosConfig & " startup-config" & vbcr
	crt.Screen.WaitForString "Destination filename [startup-config]? "
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "Switch#"
	crt.Screen.Send "reload" & chr(13)
	crt.Screen.WaitForString "Proceed with reload? [confirm]"
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "NVRAM"
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString ">"
	crt.Screen.Send "en" & chr(13)
	crt.Screen.WaitForString "Password: "
	crt.Screen.Send "uS3rious" & chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send "!" & chr(13)
	crt.Screen.Send "conf t" & chr(13)
	crt.Screen.Send "!" & chr(13)
	crt.Screen.WaitForString "(config)#"
	crt.Screen.Send "vlan 2" & chr(13)
	crt.Screen.Send "name ADMIN" & chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send "vlan 3" & chr(13)
	crt.Screen.Send "name Security" & chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send "vlan 4" & chr(13)
	crt.Screen.Send "name POS" & chr (13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send "vlan 5" & chr(13)
	crt.Screen.Send "name EFT" & chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send "vlan 6" & chr(13)
	crt.Screen.Send "name Integration" & chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send "vlan 100" & chr(13)
	crt.Screen.Send "name Voice" & chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send "vlan 200" & chr(13)
	crt.Screen.Send "name SSW-MGMT" & chr(13)	
	crt.Screen.WaitForString "#"
	crt.Screen.Send "vlan 201" & chr(13)
	crt.Screen.Send "name SSW-EMPLOYEE" & chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send "vlan 202" & chr(13)
	crt.Screen.Send "name SSW-GUEST" & chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send "vlan 203" & chr(13)
	crt.Screen.Send "name SSW-TOOLS" & chr(13)
	crt.Screen.WaitForString "#"	
	crt.Screen.Send "vlan 204" & chr(13)
	crt.Screen.Send "name SSW-TOOLS-02" & chr(13)
	crt.Screen.WaitForString "#"	
	crt.Screen.Send "vlan 205" & chr(13)
	crt.Screen.Send "name SSW-TOOLS-03" & chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send "vlan 206" & chr(13)
	crt.Screen.Send "name SSW-TOOLS-04" & chr(13)
	crt.Screen.WaitForString "#"	
	crt.Screen.Send "vlan 207" & chr(13)
	crt.Screen.Send "name SSW-TOOLS-05" & chr(13)
	crt.Screen.WaitForString "#"	
	crt.Screen.Send "vlan 208" & chr(13)
	crt.Screen.Send "name SSW-TOOLS-06" & chr(13)
	crt.Screen.WaitForString "#"	
	crt.Screen.Send "vlan 209" & chr(13)
	crt.Screen.Send "name SSW-TOOLS-07" & chr(13)
	crt.Screen.WaitForString "#"	
	crt.Screen.Send "vlan 210" & chr(13)
	crt.Screen.Send "name SSW-TOOLS-08" & chr(13)
	crt.Screen.WaitForString "#"	
	crt.Screen.Send "vlan 211" & chr(13)
	crt.Screen.Send "name SSW-TOOLS-09" & chr(13)
	crt.Screen.WaitForString "#"	
	crt.Screen.Send "vlan 212" & chr(13)
	crt.Screen.Send "name SSW-TOOLS-10" & chr(13)	
	crt.Screen.WaitForString "#"	
	crt.Screen.Send "vlan 213" & chr(13)
	crt.Screen.Send "name SSW-TOOLS-11" & chr(13)
	crt.Screen.WaitForString "#"		
	crt.Screen.Send "vlan 214" & chr(13)
	crt.Screen.Send "name SSW-TOOLS-12" & chr(13)
	crt.Screen.WaitForString "#"	
	crt.Screen.Send "vlan 215" & chr(13)
	crt.Screen.Send "name SSW-TOOLS-13" & chr(13)	
	crt.Screen.WaitForString "#"	
	crt.Screen.Send "vlan 216" & chr(13)
	crt.Screen.Send "name SSW-TOOLS-14" & chr(13)
	crt.Screen.WaitForString "#"		
	crt.Screen.Send "vlan 217" & chr(13)
	crt.Screen.Send "name SSW-TOOLS-15" & chr(13)
	crt.Screen.WaitForString "#"		
	crt.Screen.Send "vlan 218" & chr(13)
	crt.Screen.Send "name SSW-TOOLS-16" & chr(13)	
	crt.Screen.WaitForString "#"	
	crt.Screen.Send "vlan 219" & chr(13)
	crt.Screen.Send "name SSW-TOOLS-17" & chr(13)
	crt.Screen.WaitForString "#"		
	crt.Screen.Send "!" & chr(13)
	crt.Screen.Send "end" & chr(13)
	crt.Screen.Send "!" & chr(13)
	crt.Screen.Send "wr" & chr(13)
	crt.Screen.Send "!" & chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send "term len 0" & chr(13)
	crt.Screen.WaitForString "#"
'' getting serial number
	
	Do
		crt.Screen.Send "sh ver | in Processor board" & vbcr

		nResult = Crt.Screen.WaitForString("Processor board ID", 3)

   		If nResult <> FALSE Then

			'msgbox "Got Serial Number"

			Exit Do

   		End If
	Loop

	strResult = crt.Screen.ReadString("#")
	'MsgBox strResult
	Set re = New RegExp
	re.Pattern = "[\S\s]+([0-9A-Z]{11})"
	Set matches = re.Execute(strResult)
	Set match = matches(0)
	strSerial = match.SubMatches(0)
	'MsgBox "Serial number extracted as: " & strSerial
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	
	
' getting model number

	Set re = New RegExp
	re.pattern = "^[\w]{5}\s(\S+)\s\(\S+\).*"
	re.Global = TRUE
	Do
		Crt.Screen.send "sh ver | in bytes of memory" & vbcr
		
		nResult = Crt.Screen.WaitForString("bytes of memory", 2)

		If nResult <> FALSE Then
			'msgbox "Got Model number at last"
			Exit Do
		End If
	Loop
	nRow = Crt.Screen.CurrentRow
	Crt.Screen.WaitForStrings "#", 2
	szLine = Crt.Screen.Get(nRow-1,0,nRow-1,100)
	Set Matches = re.Execute(szLine)
	Set Match = Matches(0)
	'MsgBox "Model number is = " & Match.Submatches(0) 
	strModelNum = Match.Submatches(0)
	'MsgBox strModelNum
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	
' let's start logging

	Dim objTab
	Set objTab = crt.GetScriptTab

	'objTab.Screen.Send "term len 0" & vbcr


	LogFile = "C:\Logs\%Y-%M-%D\(ModelNum+SerialNum).txt"
	'LogFile = "C:\Logs\%Y-%M-%D\(ModelNum)-(SerialNum).txt"

	If objTab.Session.Logging Then
		' Turn off logging before setting up our script's
		' logging...
		objTab.Session.Log False
	End If
			
	NLogFile = Replace(LogFile, "ModelNum+SerialNum", (strSerial) & "-" & (strModelNum))
		
	objTab.Screen.Send "sh ver" & vbcr
	objTab.Screen.Send "dir" & vbcr
	objTab.Screen.Send "sh inv" & vbcr
	objTab.Screen.Send "sh run" & vbcr
	objTab.Screen.Send "sh vlan" & vbcr
	'crt.sleep (20000)
	'objTab.Screen.Send "!Serial Number is:" & strSerial
	'objTab.Screen.Send chr(13)
	'objTab.Screen.Send "!Model Number is:" & strModelNum
	'objTab.Screen.Send chr(13)
	objTab.Screen.Send chr(13)
		
	objTab.Session.LogFileName = NLogFile
	objTab.Session.Log True

	Do
        	bCursorMoved = objTab.Screen.WaitForCursor(1)
    	Loop until bCursorMoved = False
	'crt.sleep (20000)
	objTab.Screen.WaitForString vbcr, 1
    	objTab.Screen.WaitForString vblf, 1
	objTab.Session.Log False
	objTab.Screen.Send "term len 48" & vbcr
	
End Sub