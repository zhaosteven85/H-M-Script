#$language = "VBScript"
#$interface = "1.0"

crt.Screen.Synchronous = True

iosConfig = InputBox(_
	"Please specify the filename:", _
	"Target File?", _
	"router333.txt")

' This automatically generated script may need to be
' edited in order to work correctly.

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

	crt.Screen.WaitForString "#"
	crt.Screen.Send chr(13)	
	crt.Screen.WaitForString "#"
	crt.Screen.Send "format flash:" & chr(13)
	crt.Screen.WaitForString "Format operation may take a while. Continue? [confirm]"
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "Format operation will destroy all data in " & chr(34) & "flash:" & chr(34) & ".  Continue? [confirm]"
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "#"


	crt.Screen.Send chr(13)
	crt.Screen.Send "conf t " & chr(13) 
	crt.Screen.WaitForString "(config)#"
	crt.Screen.Send "int g0/0" & chr(13)
	crt.Screen.WaitForString "(config-if)#"
	crt.Screen.Send "ip address dhcp" & chr(13)
	crt.Screen.WaitForString "(config-if)#"
	crt.Screen.Send "no shut" & chr(13)
	crt.Screen.WaitForString "(config-if)#"
	crt.Screen.Send "end" & chr(13)
	crt.Screen.WaitForString "mask 255.255.255.0"
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send "copy tftp flash:" & chr(13)
	crt.Screen.Send "10.201.2.8" & chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send "c2900-universalk9-mz.SPA.151-4.M5.bin" & chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	
	
	
	crt.Screen.WaitForString "bytes copied"	
'	crt.Screen.Send "copy tftp flash:" & chr(13)
'	crt.Screen.Send "10.201.2.8" & chr(13)
'	crt.Screen.Send "IOS-S636-CLI.pkg" & chr(13)
'	crt.Screen.Send chr(13)
'	crt.Screen.Send chr(13)
'	crt.Screen.Send chr(13)




'	crt.Screen.WaitForString "bytes copied"
	crt.Screen.WaitForString "#"
	crt.Screen.Send "write erase" & chr(13)
	crt.Screen.WaitForString "Erasing the nvram filesystem will remove all configuration files! Continue? [confirm]"
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "Erase of nvram: complete"
	crt.Screen.Send chr(13)
	
	
	
	
	crt.Screen.Send "copy tftp flash:" & chr(13)
	crt.Screen.Send "10.201.2.8" & chr(13)
	crt.Screen.Send "music-on-hold.au" & chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "bytes copied"



	crt.Screen.Send "copy tftp flash:" & chr(13)
	crt.Screen.Send "10.201.2.8" & chr(13)
	crt.Screen.Send "Projects/" & iosConfig & chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "bytes copied"
		crt.Screen.Send "copy flash:" & iosConfig & " running-config" & vbcr
		crt.Screen.Send chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.WaitForString "bytes copied"
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "#"

	crt.Screen.Send "conf t" & chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "#"
	'crt.Screen.Send "lic accept end user agreement" & chr(13)
	'crt.Screen.WaitForString "ACCEPT? [yes/no]: "
	'crt.Screen.Send "yes" & chr(13)	
	crt.Screen.WaitForString "#"
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "#"
	crt.Screen.Send "exit" & chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)

	crt.Screen.WaitForString "#"
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "#"	
	crt.Screen.Send "wr" & chr(13)
	crt.Screen.WaitForString "[OK]"
	crt.Screen.Send chr(13)	
	crt.Screen.WaitForString "#"
	
	
	
	crt.Screen.Send "reload" & chr(13)
	If crt.Screen.WaitForString("System configuration has been", 3) <> TRUE Then
			crt.Screen.Send chr(13)
			Else
			crt.Screen.Send "no" & vbcr
		End If
		crt.Screen.Send chr(13)
		
	crt.Screen.WaitForString "Press RETURN to get started!"
	crt.Screen.Send chr(13)
	crt.Screen.WaitForString "Password: "
	crt.Screen.Send "R3ally!!" & chr(13)
	crt.Screen.WaitForString ">"
	crt.Screen.Send "en" & chr(13)
	crt.Screen.WaitForString "Password: "
	crt.Screen.Send "uS3rious" & chr(13)

	
		crt.Screen.Send chr(13)
	crt.Screen.Send "term len 0" & vbcr
	
	' getting serial number
	
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

	Do
		Crt.Screen.send "sh ver | in bytes of memo" & vbcr
		'Crt.Screen.send "sh ver | in bytes" & vbcr
		nResult = Crt.Screen.WaitForString("bytes of memory", 2)

		If nResult <> FALSE Then
			'msgbox "Got Model number at last"
			Exit Do
		End If
	Loop
	nRow = Crt.Screen.CurrentRow
	Crt.Screen.WaitForStrings "#", 2
	szLine = Crt.Screen.Get(nRow-1,0,nRow-1,100)	

	Set re = New RegExp
	re.Pattern = "^[\w]{5}\s(\S+)\/.+\s.+|^[\w]{5}\s(\S+)\s.+"
	Set matches = re.Execute(szLine)
	If matches.Count > 0 Then
		Set match = matches(0)
		If match.SubMatches.Count > 0 Then
			For I = 0 To match.SubMatches.Count-1
				If match.SubMatches(I) <> "" Then
					strModelNum = Match.Submatches(I)
        		End If
			Next
		End If
		'MsgBox strModelNum	
	End If
	crt.Screen.Send chr(13)
	crt.Screen.Send chr(13)
	
' let's start logging

	Dim objTab
	Set objTab = crt.GetScriptTab

	LogFile = "C:\Logs\%Y-%M-%D\(ModelNum+SerialNum).txt"
	
	If objTab.Session.Logging Then
		' Turn off logging before setting up our script's
		' logging...
		objTab.Session.Log False
	End If
			
	NLogFile = Replace(LogFile, "ModelNum+SerialNum", (strSerial) & "-" & (strModelNum))
		
	objTab.Screen.Send "sh ver" & vbcr
	objTab.Screen.Send "dir" & vbcr
	objTab.Screen.Send "sh inv" & vbcr
	objTab.Screen.Send "sh run" & chr(13)
	objTab.Screen.Send "sh lic" & chr(13)
	'objTab.Screen.Send "sh ip ips signa count" & chr(13)
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
	objTab.Screen.WaitForString vbcr, 1
    objTab.Screen.WaitForString vblf, 1
	objTab.Session.Log False
	objTab.Screen.Send "term len 48" & vbcr
	

End Sub




