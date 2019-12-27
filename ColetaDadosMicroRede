' Open notepad 
Set WshShell = WScript.CreateObject("WScript.Shell")

    Dim GetUserN
    Dim ObjNetwork
    Dim nomeUsuarioRede
    Dim nomeGabinte
    Dim nomeDominio
    Dim nomeMAC

Set ObjNetwork = CreateObject("WScript.Network")
	GetUserN = ObjNetwork.UserName
	nomeUsuarioRede = GetUserN


Set ObjNetwork = CreateObject("WScript.Network")
	GetGabi = ObjNetwork.COMPUTERNAME
    nomeGabinte = GetGabi

Set ObjNetwork = CreateObject("WScript.Network")
	GetDomi = ObjNetwork.USERDOMAIN
    nomeDominio = GetDomi   


intCount = 0
 strMAC   = ""
 strQuery = "SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionID > ''"

 Set objWMIService = GetObject( "winmgmts://./root/CIMV2" )
 Set colItems      = objWMIService.ExecQuery( strQuery, "WQL", 48 )

 For Each objItem In colItems
     If InStr( strMAC, objItem.MACAddress ) = 0 Then
         strMAC   = strMAC & "," & objItem.MACAddress
         intCount = intCount + 1
     End If
 Next

 If intCount > 0 Then strMAC = Mid( strMAC, 2 )

 Select Case intCount
     Case 0
         nomeMAC = "No MAC Addresses were found"

     Case Else
         nomeMAC = strMAC
 End Select

WshShell.Run "notepad.exe"
WScript.Sleep 500 
WshShell.SendKeys "Usuario de Rede: " & nomeUsuarioRede
WshShell.SendKeys "{ENTER}"

WshShell.SendKeys "Acesso pelo Micro: " & nomeGabinte
WshShell.SendKeys "{ENTER}"

WshShell.SendKeys "Dominio Registrado: " & nomeDominio
WshShell.SendKeys "{ENTER}"

WshShell.SendKeys "MAC Adress: " & nomeMAC
WshShell.SendKeys "{ENTER}"

WshShell.SendKeys "Acesso realizado em: "
WshShell.SendKeys "{F5}"
