'
'copy_file.vbs
'
'This script makes a copy of a file 


'Variables for the iomode of Opentextfile 
ForReading = 1
ForWriting = 2


'Variable

mess = "	Syntax : " & wscript.scriptname	& "/if:source	/of:destination"
	

'Instantiation of args (arguments) and fso(File system) objects
set args = Wscript.Arguments.named
set fso = createobject("Scripting.FileSystemObject")

'Verification of arguments existence. If no arguments, exit. 
if args.count<>2 then
	Wscript.echo " "
	Wscript.echo mess
	wscript.quit
end if

'Verification of "if" (input file) argument: if no if argument, exit. 
if not args.Exists("if") then 
	Wscript.echo " "
	Wscript.echo mess
	wscript.quit
end if

'Verification of "of" (output file) argument: if no of argument, exit. 
if not args.Exists("if") then 
	Wscript.echo " "
	Wscript.echo mess
	wscript.quit
end if

'Get input and output files. We say tell the script the names of these files
file_in=args.item("if")
file_out=args.item("of")

'We verify if the files are empty 
if IsEmpty(file_in) OR IsEmpty(file_out) then 
	Wscript.echo " "
	Wscript.echo mess
	wscript.quit
end if

'We verify if the first file exists
if not fso.fileexists(file_in) then 
	Wscript.echo " "
	Wscript.echo "	File " & file_in & " doesn't exist"
end if


'Verify if the two files are identical
if file_in = file_out then	
	Wscript.echo " "
	Wscript.echo "	The files " & file_in & " and " & file_out & " are identical."
	Wscript.echo "	Why would you want to copy one in another?"
	wscript.quit
end if


'If output file exists, ask permission to overwrite. 
if fso.fileexists(file_out) then
	Wscript.echo " "
	Wscript.stdout.write "	Fle " & file_out & " exists. Dou you want to overwrite it? (y/n): "
	if wscript.stdin.readline <> "y" then
		wscript.quit
	end if
end if 

'Variables: input file opens in reading mode, output file opens in writing mode.
'True gets the program to write even the file downs't exist.  
Set in_file = fso.OpenTextFile(file_in, ForReading)
Set out_file = fso.OpenTextFile(file_out, ForWriting, True)

'While is not the end of the file, copy in_file in out_file 
While not in_file.AtEndOfStream
	out_file.writeline in_file.Readline
wend	
in_file.close 
out_file.close 
