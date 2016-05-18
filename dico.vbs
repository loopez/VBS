'
'dico.vbs
'
'Instantiation of dico with Dictionary
Set dico = createobject("Scripting.Dictionary")

'Adding keys and values with add method
dico.add "abc", "Contenu abc"
'Displaying those values
wscript.echo dico("abc")

'Adding keys and values without method: we use an array 
dico ("xyz")="Contenu xyz"
'Displaying those values
wscript.echo dico("xyz")

'Cstr converts number to string
dico("administrator")="name:lastname:" & Cstr(30)
wscript.echo dico("administrator")

'Split using a defined separator                        
info=split(dico("administrator"),":")
for each e in info
	wscript.echo e 
next 
