'Run a batch file without a console window.'
CreateObject("Wscript.Shell").Run "" & WScript.Arguments(0) & "", 0, False
