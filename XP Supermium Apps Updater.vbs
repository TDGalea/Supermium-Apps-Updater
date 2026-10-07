' Supermium Apps Updater by TDGalea.
' Not affiliated with Supermium in any way. Though I wouldn't mind if this was implemented into the browser itself in some way...
' This script will only work on XP (maybe W2K I guess, if Supermium ever works there...).
' There is a separate version for Vista.

On Error Resume Next

Set oShell   = WScript.CreateObject("WScript.Shell")
Set fso      = CreateObject("Scripting.FileSystemObject")

appdata      = oShell.ExpandEnvironmentStrings("%appdata%")
userprofile	 = oShell.ExpandEnvironmentStrings("%userprofile%")

Set userData = fso.GetFolder(appdata + "\Supermium\User Data\")

Function CreateShortcut
	If profile = "Default" Then
		sLinkFile = userprofile + "\Start Menu\Programs\Supermium Apps\" + name + ".lnk"
	Else
		sLinkFile = userprofile + "\Start Menu\Programs\Supermium Apps\" + name + " (" + profile + ").lnk"
	End If
	Set oLink = oShell.CreateShortcut(sLinkFile)
		oLink.TargetPath   = "C:\Program Files\Supermium\chrome_proxy.exe"
		oLink.Arguments    = "--profile-directory=""" + profile + """ --app-id=" + appid
		oLink.IconLocation = appdata + "\Supermium\User Data\" + profile + "\Web Applications\_crx_" + appid + "\" + name + ".ico"
	oLink.Save
End Function

appList     = ""
prevProfile = ""

For Each profFolder in userData.SubFolders
	profile = profFolder.Name
	Set manifestFolder = fso.GetFolder(appdata + "\Supermium\User Data\" + profile + "\Web Applications\Manifest Resources")
	
	If fso.FolderExists(manifestFolder) Then
		
		For Each manFolder in manifestFolder.SubFolders
			appid = manFolder.Name
			iconFolderPath = appdata + "\Supermium\User Data\" + profile + "\Web Applications\_crx_" + appid
			
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
