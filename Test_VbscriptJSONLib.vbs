includeRelFile "./VbscriptJSONLib.vbs"

Dim data
Dim jsLib
Dim sErr
Dim jRoot
Dim jPlugin
Dim jRequire
Dim jPlugins

Call Test_ParseErrors()

Sub Test_ParseErrors()
	data = createobject("Scripting.FileSystemObject").OpenTextFile("data.json", 1).ReadAll()

	Set jsLib = New VbscriptJSONLib
	on error resume next
	Set jRoot = jsLib.parse( data )
	on error goto 0
	
	sErr = jsLib.GetParserErrors()
	If sErr <> "" then
		Wscript.Echo "Error while rendering the supplied Json." & vbcrlf & vbcrlf & sErr
		Wscript.Quit
	End If
	
	Set jPlugin = jRoot("plugin")
	Set jRequire = jPlugin("require")
	Set jPlugins = jRequire("plugins")
	
	Wscript.Echo jPlugin("id")
	
	'Object Loop Method 1
	Dim i
	For i = 0 to jPlugins.count - 1
		Wscript.Echo jPlugins(i)
	Next

	'Object Loop Method 2
	Dim jPlug
	For each jPlug in jPlugins
		Wscript.Echo jPlug
	Next	
End Sub

Sub includeRelFile(fSpec)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub