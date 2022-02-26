'============================================='
' Impedindo a exibição de error para o usuário'
'============================================='

On error Resume Next
Err.clear 0


'=============================================='
'Mapear pastas de acordo com o grupo do usuário'
'=============================================='

set objNetwork= CreateObject("WScript.Network")
strDom = objNetwork.UserDomain
strUser = objNetWork.UserName
Set objUser = GetObject("WinNT://" & strDom & "/" & strUser & ",user")

For Each objGroup In objUser.Groups

	Select Case objGroup.Name
		Case "Grupo_1"
			If Not FSODrive.DriveExists("X:") then
				objNetwork.MapNetworkDrive "X:", "\\NOMESERVIDOR\COMPARTILHAMENTO01","true"
			End If

	Select Case objGroup.Name
		Case "Grupo_2"
			If Not FSODrive.DriveExists("X:") then
				objNetwork.MapNetworkDrive "X:", "\\NOMESERVIDOR\COMPARTILHAMENTO02","true"
			End If

	Select Case objGroup.Name
		Case "Grupo_2"
			If Not FSODrive.DriveExists("X:") then
				objNetwork.MapNetworkDrive "X:", "\\NOMESERVIDOR\COMPARTILHAMENTO02","true"
			End If
	End Select
Next

'============================='
'Mapeando impressoras e pastas'
'============================='

Set WshNetwork = Wscript.CreateObject("Wscript.Network")
WshNetwork.AddWindowsPrinterConnection "\\NOMESERVIDOR\NOMEIMPRESSORA", "NOMECOMPARTILHAMENTO"
WshNetwork.AddWindowsPrinterConnection "\\NOMESERVIDOR\NOMEIMPRESSORA2", "NOMECOMPARTILHAMENTO2"
WshNetwork.SetDefaultPrinter "\\NOMESERVIDOR\NOMEIMPRESSORAPADRAO", "NOMECOMPARTILHAMENTO"

WshNetwork.MapNetworkDrive "P:","\\NOMESERVIDOR\PUBLICA","true"

'================================='
'Criar atalho para site no Desktop'
'================================='

set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")

set oUrlLink = WshShell.CreateShortcut(strDesktop & "\NOMEDOLINK.lnk")

oUrlLink.TargetPath = "ENDERECODOSITE"

oUrlLink.IconLocation = "CAMINHODOICONE.ico"

oUrlLink.Save

'=========================================='
'Cria atalho do compartilhamento no desktop'
'=========================================='

strAppPath = "X:\"
Set wshShell = CreateObject("WScript.Shell")
objDesktop = wshShell.SpecialFolders("Desktop")
set oShellLink = WshShell.CreateShortcut(ObjDesktop & "\NOMEDOLINK.lnk")
oShellLink.TargetPath = strAppPath
oShellLink.WindowStyle = "1"
oShellLink.Description = "NOMEPASTA"
oShellLink.Save

strAppPath = "X:\"
Set wshShell = CreateObject("WScript.Shell")
objDesktop = wshShell.SpecialFolders("Desktop")
set oShellLink = WshShell.CreateShortcut(ObjDesktop & "\NOMEDOLINK.lnk")
oShellLink.TargetPath = strAppPath
oShellLink.WindowStyle = "1"
oShellLink.Description = "NOMEPASTA"
oShellLink.Save

'Envia o comando para apertar F5 para atualizar os ícones no desktop
WshShell.SendKeys "{F5}"


'================='
'Mensagem no Logon'
'================='

MsgBox ("ATENÇÃO: Digite aqui sua mensagem" & vbcrlf & "continue aqui na linha de baixo")

wscript.quit