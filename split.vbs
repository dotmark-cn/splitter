'///////////////////////////////////////////////////////////////////////
'
'description:  	simple text file splitter
'     			takes or gathers file path, output folder, lines/size
'     			splits by either MB or line spec (line is faster)
'     			output is files with original file name + _NUM
'
'author:		ryan rogers | rrogers@-------.com
'
'initial rel:	2/8/2017
'
'rev history:	2/10/2017
'					force run in cscript
'					account for user opening file in UI
'					move new file name creation into function		
'					improve file naming scheme
'					remove vars from console args case
'					modify functions to use function args
'					input validation, incl no args or invalid args
'					handle multiple argument lengths
'					error handling for file types and directory spec /
'					verify file path, file exists, before split
'					cleanup unused vars
'					rewrite console input to account for new vars
'				2/9/2017
'					declare variables and indicate "type"
'					move logic into functions/subs
'					provide usage help
'
'pending:		
'				verify if input is appropriate for splitting (is text)
'				create output folder if needed
'				improve performance of size-based splits
'				recursion to handle full folders
'				mode: point to directory, find files of x size, split
'				allow for size specs other than Mb
'				move all path checks to independent function
'				error handling for file size < size input
'				make output dir arg optional; if omitted, use current (-r)
'
'///////////////////////////////////////////////////////////////////////

'%%%%%%%% SET DEBUG FLAG %%%%%%%%
'0 = suppress debug messages
'1 = display debug messages

		const DEBUG_MODE = 1

'%%%%%%%%% CHECK ENGINE %%%%%%%%%
'before we do anything, verify we're running in cscript via command line

		call checkengine()

'%%%%%%%%% MAIN ROUTINE %%%%%%%%%

'***** VARIABLE DEFINITIONS *****

const BYTEMULT = 1048576	'constant int: multiplier for bytes->mb
const READ = 1				'constant int: value for textstream iomode
const OVERWRITE = true		'constant bool: value for texstream file overwriting
const SUCCESS = 1			'constant int: function return value success indicator
const FAIL = 0				'constant int: function return value failure indicator
const ERR_NOFILE = 0		'constant int: error for file not exist
const ERR_NODIR = 1			'constant int: error for dir not exist
const ERR_INVARG = 2		'constant int: error for invalid initial argument
const ERR_INVARGNUM = 3		'constant int: error for invalid argument number
dim args 					'array: arg array
dim ofso					'object: file system object
dim shell					'object: scripting shell
dim oreadfile				'object: file object to be read
dim ilinecounter			'int: counter for lines written to file
dim ifilecounter 			'int: counter for files output to folder
dim ibytecounter			'int: counter for bytes written to file
dim line 					'string: line being read from file
dim owritefile				'object: file object to be written
dim filex					'string: input file extension
dim filename				'string: input file name
dim inputarr(4)				'string array: console input string array

'***** INITIALIZE OBJECTS, SET COUNTERS, GATHER ARGUMENTS *****

set args = wscript.arguments 'mode, filepath, outputpath, size

set ofso = createobject("scripting.filesystemobject")
set shell = createobject("wscript.shell")

ilinecounter = 0
ifilecounter = 0
ibytecounter = 0

if (args.length > 0) then

	select case args(0)

		case "/?"

			call runhelp()

		case "-l"

			'check paths, if exist, call processing, else show error
			
			if(not(checkpath(args(1), "file"))) then

				'file doesn't exist
				call errmsg(ERR_NOFILE, args(1))

			else
			
				if(not(checkpath(args(2), "dir"))) then

					'dir doesn't exist
					call errmsg(ERR_NODIR, args(2))
				
				else

					call processline(args(1), fixdirpath(args(2)), args(3))

				end if

			end if

		case "-s"

			'check paths, if exist, call processing, else show error
			
			if(not(checkpath(args(1), "file"))) then

				'file doesn't exist
				call errmsg(ERR_NOFILE, args(1))

			else
			
				if(not(checkpath(args(2), "dir"))) then

					'dir doesn't exist
					call errmsg(ERR_NODIR, args(2))
				
				else

					call processsize(args(1), fixdirpath(args(2)), args(3))

				end if

			end if

		case "-p"

			'gather input via stdin

			getinput()

			select case inputarr(0)

				case "l"

					call processline(inputarr(1), fixdirpath(inputarr(2)), inputarr(3))

				case "s"

					call processsize(inputarr(1), fixdirpath(inputarr(2)), inputarr(3))

			end select

			'call appropriate function

		case else

			if ((args.length < 4) OR (args.length > 4)) then

				'if too few or too many args provided, raise error, show usage
				call errmsg(ERR_INVARGNUM, args.length)
			
			elseif (args.length = 4) then
			
				'if 4 args provided but first arg is invalid, raise error, show usage
				call errmsg(ERR_INVARG, args(0))
			
			end if

	end select

else

	'if no args provided, call runusage()
	call runusage()

end if

'%%%%%% FUNCTION/SUB DEFINITONS %%%%%%

function checkengine()

	engine = ucase(mid(wscript.fullname, instrrev(wscript.fullname, "\")+1))

	select case engine

		case "WSCRIPT.EXE"

			set args= wscript.arguments

			if args.length > 0 then

				for i = 0 to (args.length-1)

					sargs = sargs & " " & args(i)

				next

				sdiag = "This script is intended to be run from the command line using cscript" & vbcrlf & vbcrlf &_
						"e.g.: cscript " & wscript.scriptname & sargs & vbcrlf & vbcrlf &_
						"e.g.: cscript " & wscript.scriptname & " /? (for help)"

			else

				sdiag = "This script is intended to be run from the command line using cscript" & vbcrlf & vbcrlf &_
						"e.g.: cscript " & wscript.scriptname & vbcrlf & vbcrlf &_
						"e.g.: cscript " & wscript.scriptname & " /? (for help)"

			end if

			result = msgbox(sdiag, 0+48, "Oops!")

			wscript.quit
		
		case else

			'doing nothing if not wscript engine

	end select

end function

sub writemsg(smessage)
	
	wscript.echo(smessage)

end sub

sub debugmsg(smessage)

	if (DEBUG_MODE) then
		writemsg(smessage)
	end if

end sub

sub runusage()

	writemsg("-- split.vbs --")
	writemsg("to view help and usage examples, use the /? flag")
	writemsg("")
	writemsg("ex.:	cscript split.vbs /?")

end sub

sub runhelp()
	
	writemsg("-- split.vbs help -- ")
	writemsg("")
	writemsg("when run as a single-line command, split takes 4 arguments:")
	writemsg("")
	writemsg("	mode:  either -l for 'line mode' or -s for 'size mode'")
	writemsg("	input file: full filesystem path to file being split, including file extension")
	writemsg("	output directory:  file filesystem path to directory for output files")
	writemsg("	size:  a number that represents either lines or megabytes, depending on mode")
	writemsg("")
	writemsg("ex.:  cscript split.vbs -l c:\myinputfile.txt c:\mydirectory 10000")
	writemsg("ex.:  cscript split.vbs -s c:\myinputfile.txt c:\mydirectory 50")
	writemsg("")
	writemsg("split can also be run with the -p argument, in which case the script will prompt for necessary values")
	writemsg("")
	writemsg("ex.:	cscript split.vbs -p")
	writemsg("")

end sub

sub errmsg(errtype, data)

	'takes err type, output err message to window

	select case errtype

		case ERR_NOFILE

			writemsg("ERR: File does not exist.")
			writemsg("File path provided: " & data)

		case ERR_NODIR

			writemsg("ERR: Directory does not exist.")
			writemsg("Directory path provided: " & data)

		case ERR_INVARG

			writemsg("ERR: Initial argument is invalid.")
			writemsg("Initial argument provided: " & data)
			writemsg("")
			call runusage()

		case ERR_INVARGNUM

			writemsg("ERR: Wrong number of arguments provided.")
			writemsg("Expected 4 arguments; received: " & data)
			writemsg("")
			call runusage()

		case else

	end select 

end sub

function checkpath(path, mode)

	'returns true/false

	select case mode

		case "file"

			checkpath = ofso.fileexists(path)

		case "dir"

			if(ofso.folderexists(path)) then
				checkpath = ofso.folderexists(path)
			else
				'call directory create subroutine
				'ask if want to create directory
				'if y then create and return true
				'if n then return false
				'leaving original return value until implemented
				checkpath = ofso.folderexists(path)
			end if

	end select

end function

function fixdirpath(path)
	
	'check path, if missing / add it and return, else return original value

	if (right(trim(path),1) = "\") then

		fixdirpath = path

	else

		fixdirpath = path & "\"

	end if

end function

function getfileext(path)

	'return file extension (could add err handling for bad paths here too)
	getfileext = mid(path, instrrev(path, "."))

end function

function getfilename(path)

	getfilename = ofso.getbasename(path)

end function

function getinput()

	'gather input, place into inputarr, return inputarr
	'do filepath verification here

	wscript.stdout.write("Enter a mode option [l for line mode, s for size/Mb mode]>")
	inputarr(0) = wscript.stdin.readline()

	if(NOT inputarr(0) = "l") then

		if(NOT inputarr(0) = "s") then

			debugmsg("evaluated false")
			call errmsg(ERR_INVARG, inputarr(0))
			wscript.quit

		end if

	end if

	wscript.stdout.write("Enter the full path to the file to be split>")
	inputarr(1) = wscript.stdin.readline()

	if(NOT checkpath(inputarr(1), "file")) then

		call errmsg(ERR_NOFILE, inputarr(1))
		wscript.quit

	end if

	wscript.stdout.write("Enter the full path to the output directory [should already exist]>")
	inputarr(2) = wscript.stdin.readline()

	if(NOT checkpath(inputarr(2), "dir")) then

		call errmsg(ERR_NODIR, inputarr(2))
		wscript.quit

	end if

	select case inputarr(0)
		case "l"
			wscript.stdout.write("Enter number of lines per file [e.g. 100000]>")
		case "s"
			wscript.stdout.write("Enter Mb size per file [e.g. 50]>")
	end select

	inputarr(3) = wscript.stdin.readline()

	getinput = inputarr

end function

function getnewfilename(path, counter)

	getnewfilename = getfilename(path) & "_" & counter & getfileext(path)

end function

function processline(source, output, numlines)
	
	writemsg("Beginning file processing...")

	set oreadfile = ofso.opentextfile(source, READ)

	do until oreadfile.atendofstream

		line = oreadfile.readline

		if ilinecounter = numlines-1 then

			'reset line counter
			ilinecounter = 0
			
			'write last line and close
			owritefile.writeline(line)
			owritefile.close

		else

			if ilinecounter = 0 then

				writemsg(numlines*ifilecounter & " lines written [line counter evaluated = 0], generating new file '" & getnewfilename(source, ifilecounter) & "'")

				'generate new file
				set owritefile = ofso.createtextfile(output & getnewfilename(source, ifilecounter), OVERWRITE)

				'increment file counter
				ifilecounter = ifilecounter + 1

				owritefile.writeline(line)

				'increment line counter
				ilinecounter = ilinecounter + 1

			else
			
				'write to existing file
				owritefile.writeline(line)

				'increment line counter
				ilinecounter = ilinecounter + 1
				
			end if

		end if

	loop

	writemsg(ifilecounter & " files written, script complete.")

	owritefile.close
	oreadfile.close

end function

function processsize(source, output, chunksize)

	writemsg("Beginning file processing...")

	set oreadfile = ofso.opentextfile(source, READ)

	do until oreadfile.atendofstream

		line = oreadfile.readline

		'reset bytecounter if bytes > MB*1048576
		'also the time to create a new owritefile

		if ibytecounter >= chunksize*BYTEMULT then

			'reset byte
			ibytecounter = 0
			
			'write last line and close
			owritefile.writeline(line)
			owritefile.close

		else

			if ibytecounter = 0 then

				writemsg(chunksize*ifilecounter & " Mb written [bytes evaluated = 0], generating new file '" & getnewfilename(source, ifilecounter) & "'")

				'generate new file
				set owritefile = ofso.createtextfile(output & getnewfilename(source, ifilecounter), OVERWRITE)

				owritefile.writeline(line)

				'set bytcounter to current file bytes
				set getfile = ofso.getfile(output & getnewfilename(source, ifilecounter))
				ibytecounter = getfile.size

				'increment file counter
				ifilecounter = ifilecounter + 1

			else
			
				'write to existing file
				owritefile.writeline(line)

				'set bytecounter to current file bytes
				set getfile = ofso.getfile(output & getnewfilename(source, ifilecounter-1))
				ibytecounter = getfile.size

			end if

		end if

	loop

	writemsg(ifilecounter & " files written, script complete.")

	owritefile.close
	oreadfile.close

end function


