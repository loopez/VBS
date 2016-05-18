'
'add_ou.vbs
'

' connexion string
domain="DC=freddie,DC=NET"

' Connexion to active directory  
set AD = GetObject("LDAP://" & domain)

' Creation of an objet Organizational Unit 
set objOu = AD.Create("OrganizationalUnit","ou=TESTOU")

'Écrire dans AD
objOu.setinfo 



