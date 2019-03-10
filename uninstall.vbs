' Module: Global Uninstaller.
'    Author: Sergey Trishkin
'    Version: 1.0.0.0
'    Description: 
'		Module help uninstall program or many old version of programs with use register.
'
' Example:
' 	dim Branch, BranchNode, NodeValue, NodeForLook, NodeForLookConstValue
'	
'		Branch				  - string - Branch of reg with programs uninstall data (64 or 32);
'		BranchNode			  - string - key of reg with program name of some other special value;
'		NodeValue			  - string - value that BranchNode should be;
'		NodeForLook			  - string - node than script will compare;
'		NodeForLookConstValue - string - value then node mustn't equ.
'
'	Branch           	   = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
' 	BranchNode       	   = "DisplayName"
' 	NodeValue              = "Mozilla Maintenance Service"
' 	NodeForLook            = "DisplayVersion"
' 	NodeForLookConstValue  = "66.00"
'
' 	Main Branch, BranchNode, NodeValue, NodeForLook, NodeForLookConstValue
'   all version of Mozilla NEQ 66.00 will uninstall.
'
' Sorry my English and good luck!

dim reg_path, HKLM, oReg, WshShell, CurentNode, comand, isEmpty, remotely

HKLM       = 2147483650			'HKey Local Machine.
CurentNode = "UninstallString"  'RegKey to call uninstall. 
remotely   = 0					'How many version of program was remoted.
SilentKey  = " /S"				'StartKey for silent uninstall.

Set WShell = CreateObject("Wscript.Shell") 							'For Run.
Set oReg   = GetObject("winmgmts:{impersonationLevel=impers" & _
					   "onate}!root\default:StdRegProv") 			'For RegRead.
errBook    = Array( _
	"Parameter is not exists. [", _
	"Uninstall string not found, please change value.", _
	"The script worked, but no one condition was not met.", _
	"The script worked, programs were uninstalled: ", _
	"when we tried to call the file something went wrong.", _
	"Main function have 5 args, but called with: ")         'Array with application error.


function Main(fBranch, fBranchNode, fNodeValue, fNodeForLook, fNodeForLookConstValue)
	' Function Main.
	'
	' 	Parameters:
	'		fBranch      		   - string - RegBranch   - val with system uninstall reg path; 
	'		fBranchNode  		   - string - RegKey with - key than we need check for find current key; 
	'		fNodeValue   		   - string - RegKeyValue - current value of fBranchNode;
	'		fNodeForLook           - string - RegKey      - key than we need look that check condition; 
	'		fNodeForLookConstValue - string - RegKeyValue - if script get from reg this value script skip this.
	'
	'	All parameters is require.
	
	dim keysArr
	
	if oReg.EnumKey(HKLM, fBranch, keysArr) <> 0 then ErrorCather Array(0, "[Branch]")
	
	for n = 0 to UBound(keysArr)
		dim n, val, fullPath, response, path
		fullPath = fBranch & "\" & keysArr(n)
		
		if oReg.GetStringValue(HKLM, fullPath, fBranchNode, val) = 0 then 
			if val = fNodeValue then LookCondition fullPath, fNodeForLook, fNodeForLookConstValue
		end if
	next
	
	if remotely = 0 then 
		ErrorCather 2
	else 
		ErrorCather Array(3, remotely)
	end if 
	
end function


function LookCondition(path, forlook, condition)
	' Function LookCondition.
	'
	' 	Parameters:
	' 		path      - string - RegBranch - key with reg branch; 
	'		forlook   - string - RegKey	   - key for read; 
	'		condition - string - Value 	   - if forlook EQU condition will return nul.
	'
	' 	All parameters is require.
	
	dim result, resp, curent
	
	oReg.GetStringValue HKLM, path, forlook, result
	if TypeName(result) = "Null" then ErrorCather Array(0, "[NodeForLook]")
	
	if result = condition then 
		exit function
		
	else
		oReg.GetStringValue HKLM, path, CurentNode, curent
		if TypeName(curent) = "Null" then ErrorCather Array(0, "[CurentNode]")
		if WShell.Run(curent & SilentKey, 6, true) <> 0 then ErrorCather Array(4, "[" & curent & "]")
		remotely = remotely + 1
		
	end if
	
end function


function ErrorCather(args)
	' Function ErrorCather.
	'
	'	 Parameters:
	'		args - Array  - Array, where Arr(0) is error number from errBook Array and Arr(1) some sting.
	'			   Number - is error number from errBook Array.
	'							
	' 	 All parameters is require.   
	
	dim msg
	
	if TypeName(args) = "Variant()" then
		msg = errBook(args(0)) & args(1)	
	else
	    msg = errBook(args)
	end if
	
	MsgBox msg, 0, "Attention"
	WScript.quit
	
end function


if WScript.Arguments.Count <> 0 then 
	dim argv
	set argv = WScript.Arguments
	
	if argv.Count < 5 then ErrorCather(Array(5, "'" + cstr(argv.Count) + "'"))
	Main argv(0), argv(1), argv(2), argv(3), argv(4)
	
end if