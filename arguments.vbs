'
' This script repeats the command line arguments
'
Set objArgs = WScript.Arguments  		'Instantiation (objArgs) of the WScript objet to get the arguments
if objArgs.Count <> 0 then		'Test for the existence of arguments  The operator <> means different
	wscript.echo VBCrLf & VBTab & "Given arguments list: " & VBCrLf
	For I = 0 to objArgs.Count - 1		'For each argument 
		WScript.Echo VBTab & "Argument " & i & " = " & objArgs(I)   'We show the arguments 
	Next
else
	wscript.echo VBCrLf & VBTab & "No argument has been passed" & VBCrLf
end if
