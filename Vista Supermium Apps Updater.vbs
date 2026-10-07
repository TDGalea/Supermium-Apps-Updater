' Supermium Apps Updater by TDGalea.
' Not affiliated with Supermium in any way. Though I wouldn't mind if this was implemented into the browser itself in some way...
' This script will only work on Vista (and up I guess, if Supermium ever breaks for 7 onward...).
' There is a separate version for XP.

On Error Resume Next

Set oShell   = WScript.CreateObject("WScript.Shell")
Set fso      = CreateObject("Scripting.FileSystemObject")

localappdata = oShell.ExpandEnvironmentStrings("%localappdata%")
appdata      = oShell.ExpandEnvironmentStrings("%appdata%")

Set userData = fso.GetFolder(localappdata + "\Supermium\User Data\")

Function CreateShortcut
	If profile = "Default" Then
		sLinkFile = appdata + "\Microsoft\Windows\Start Menu\Programs\Supermium Apps\" + name + ".lnk"
	Else
		sLinkFile = appdata + "\Microsoft\Windows\Start Menu\Programs\Supermium Apps\" + name + " (" + profile + ").lnk"
	End If
	Set oLink = oShell.CreateShortcut(sLinkFile)
		oLink.TargetPath   = "C:\Program Files\Supermium\chrome_proxy.exe"
		oLink.Arguments    = "--profile-directory=""" + profile + """ --app-id=" + appid
		oLink.IconLocation = localappdata + "\Supermium\User Data\" + profile + "\Web Applications\_crx_" + appid + "\" + name + ".ico"
	oLink.Save
End Function

appList     = ""
prevProfile = ""

For Each profFolder in userData.SubFolders
	profile = profFolder.Name
	Set manifestFolder = fso.GetFolder(localappdata + "\Supermium\User Data\" + profile + "\Web Applications\Manifest Resources")
	
	If fso.FolderExists(manifestFolder) Then
		
		For Each manFolder in manifestFolder.SubFolders
			appid = manFolder.Name
			iconFolderPath = localappdata + "\Supermium\User Data\" + profile + "\Web Applications\_crx_" + appid
			
			If fso.FolderExists(iconFolderPath) Then
				Set iconFolder = fso.GetFolder(iconFolderPath)
				For Each file In iconFolder.Files
					If LCase(fso.GetExtensionName(file.Name)) = "ico" Then
						name = fso.GetBaseName(file.Name)
						CreateShortcut
						If profile = prevProfile Then
							appList = appList + ", " + name
						Else
							appList = appList + vbCrLf + vbCrLf + "Profile: " + profile + vbCrLf + "Apps: " + name
						End If
						prevProfile = profile
						Exit For
					End If
				Next
			End If
		Next
	End If
Next

x = msgbox("The following WebApp shortcuts have been created:" + appList,64,"Supermium Apps Updater")
