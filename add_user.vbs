'
'add_user.vbs
'

' connexion string
' OU for Organisational Unit or CN for a container (eg. Users)
strUser="OU=scriptage,DC=freddie,DC=NET"

' Connexion to active directory  
set user = GetObject("LDAP://" & strUser)

' Creation of an objet user
set objUser = user.Create("user","cn=user01")

' Writing in active directory. Infos to create the user: user name, login name, script path, 
objUser.put "SamAccountName","user01"
objUser.put "givenname","user01"
objUser.put "scriptPath","usager01.vbs"

'Testing for erros in the creation: we display error number an description 
on error resume next 
objUser.setinfo
if err.number <> 0 then
	wscript.echo err.number & vbTab & err.description
end if

'True to create a disabled account
objUser.AccountDisabled = false
objUser.setinfo

'User password 
objUser.setpassword "Crosemont1"
objUser.setinfo






