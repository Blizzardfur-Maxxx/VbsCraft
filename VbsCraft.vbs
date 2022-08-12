Set WshShell = WScript.CreateObject("WScript.Shell")
Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.3.0")
Set oFSO = CreateObject("Scripting.FileSystemObject")
x=msgbox("a friendly minecraft launcher in vbs" ,vbYesNo, "VBSCraft-Luancher")


intAnswer = _
    Msgbox("do you need to download files? (not needed if lib folder exists)",vbYesNo, "do you want to run Minecraft?")
If intAnswer = vbYes Then
    WScript.sleep 200
	'download libs and game
	oFSO.CreateFolder "lib"
	oFSO.CreateFolder "lib/native"
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://piston-data.mojang.com/v1/objects/3a799f179b6dcac5f3a46846d687ebbd95856984/client.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/minecraft.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/jinput.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/jinput.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/jopt-simple-4.5.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/jopt-simple-4.5.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/lwjgl_test.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lwjgl_test.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/lwjgl_util.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lwjgl_util.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/lwjgl_util_applet.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lwjgl_util_applet.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/lwjgl.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lwjgl.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/lzma.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lzma.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/lwjgl-debug.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lwjgl-debug.jar"
    oStream.Close
	'download natives
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/jinput-dx8.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/jinput-dx8.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/jinput-dx8_64.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/jinput-dx8_64.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/jinput-raw.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/jinput-raw.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/jinput-raw_64.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/jinput-raw_64.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/lwjgl.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/lwjgl.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/lwjgl64.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/lwjgl64.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/OpenAL32.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/OpenAL32.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://raw.githubusercontent.com/Blizzardfur-Maxxx/Assets-For-Projetcs/main/OpenAL64.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/OpenAL64.dll"
    oStream.Close
	'run minecraft after download
	intAnswer = _
    Msgbox("Do you want to run minecraft?",vbYesNo, "quit")
	If intAnswer = vbYes Then
	wshshell.Run "java -Djava.library.path=lib\native -Xmx1024M -cp lib\lwjgl.jar;lib\lwjgl_util.jar;lib\jinput.jar;lib\minecraft.jar com.mojang.minecraft.Minecraft"
	Else
    Msgbox "bye"
	End If

'run minecraft plain
Else
intAnswer = _
    Msgbox("Do you want to run minecraft?",vbYesNo, "quit")
If intAnswer = vbYes Then
	wshshell.Run "java -Djava.library.path=lib\native -Xmx1024M -cp lib\lwjgl.jar;lib\lwjgl_util.jar;lib\jinput.jar;lib\minecraft.jar com.mojang.minecraft.Minecraft"
Else
    Msgbox "bye"
End If
End If
