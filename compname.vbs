'Instantiation
Set WshNetwork= WScript.CreateObject("WScript.Network")

WScript.Echo "Nom ordinateur = " & WshNetwork.ComputerName & VBCRLF & _
"Nom utilisateur = " & WshNetwork.UserName & VBCRLF & _
"Nom du domaine =  "& WshNetwork.UserDomain 
