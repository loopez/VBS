'
'get_user.vbs
'
'This script displays a username, the OU name and the domain name.
'The user, OU and domain have to exist in the server.

'We will user this name for our object:
Dim objLDAP

'
'Get domain name automatically so we can use the script in a different machine  
Set objRootDSE = getObject("LDAP://rootDSE")
domain = objRootDSE.Get("DefaultNamingContext")

'We use a variable for username
user = "freddie lopez"


'Instantiation of an object containing an AD user. 
'We use " " for strings, variables go outside " "
' & is used to concatenate strings and variables 
'CN= complete user name
'OU= Organisational Unit
'DC= domain name
' info will allow us t
Set objLDAP = GetObject("LDAP://CN=" & user & ",OU=scriptage," & domain)
info=""

'The WITH statement executes a series of statements on a single object.
with objLDAP
	info = info & "Full name : " & .fullname & VBCRLF
	info = info & "Login name : " &.sAMAccountName & VBCRLF
	info = info & "Boot script : " &.scriptPath & VBCRLF
end with 
wscript.echo info 
