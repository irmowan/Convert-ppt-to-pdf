on run {input, parameters}
	set theOutput to {}
	tell application "Microsoft PowerPoint" -- work on version 15.15 or newer
		launch
		set theDial to start up dialog
		set start up dialog to false
		repeat with i in input
			open i
			set pdfPath to my makeNewPath(i)
			save active presentation in pdfPath as save as PDF -- save in same folder
			close active presentation saving no
			set end of theOutput to pdfPath as alias
		end repeat
		set start up dialog to theDial
		quit
	end tell
	return theOutput
end run

on makeNewPath(f)
	set t to f as string
	if t ends with ".pptx" then
		return (text 1 thru -5 of t) & "pdf"
	else
		return (text 1 thru -4 of t) & "pdf"
	end if
end makeNewPath
