option explicit

'Set Variables
Dim args, argCount, base, merge, output, working, word

'Capture args and get arg count'
Set args = WScript.Arguments
argCount = args.Count


'If args is greater than two, get out'
if argCount > 2 then
	WScript.echo "To many aguments"
	WScript.Quit 1
end if

'If all is good and dandy set files to vars
base = args(0)
merge = args(1)

'MsgBox "merging " + merge + " into " + base 

Set word = createobject("Word.Application")

'make word visible and open the base document
word.visible = True
set working = word.Documents.Open(base)


'Merge it and close the base document
working.Merge(merge)
working.Close

'Success!
