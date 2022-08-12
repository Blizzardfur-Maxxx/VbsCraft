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
	oXMLHTTP.Open "GET", "https://cdn.discordapp.com/attachments/941767122642669634/1007452371636465684/minecraft.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/minecraft.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://cdn.discordapp.com/attachments/854426771621150741/854449585166352405/jinput.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/jinput.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://cdn.discordapp.com/attachments/854426771621150741/854449586194219009/jopt-simple-4.5.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/jopt-simple-4.5.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://cdn.discordapp.com/attachments/854426771621150741/854449589884157962/lwjgl_test.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lwjgl_test.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://cdn.discordapp.com/attachments/854426771621150741/854449593538183208/lwjgl_util.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lwjgl_util.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://cdn.discordapp.com/attachments/854426771621150741/854449594666844180/lwjgl_util_applet.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lwjgl_util_applet.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://cdn.discordapp.com/attachments/854426771621150741/854449590702964796/lwjgl.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lwjgl.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://cdn.discordapp.com/attachments/854426771621150741/854449599226052608/lzma.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lzma.jar"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://cdn.discordapp.com/attachments/854426771621150741/854449599379669042/lwjgl-debug.jar", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/lwjgl-debug.jar"
    oStream.Close
	'download natives
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://cdn.discordapp.com/attachments/854426771621150741/854449775750938636/jinput-dx8.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/jinput-dx8.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://cdn.discordapp.com/attachments/854426771621150741/854449777286053888/jinput-dx8_64.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/jinput-dx8_64.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://cdn.discordapp.com/attachments/854426771621150741/854449777936433152/jinput-raw.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/jinput-raw.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://cdn.discordapp.com/attachments/854426771621150741/854449779257770014/jinput-raw_64.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/jinput-raw_64.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://cdn.discordapp.com/attachments/854426771621150741/854449780728791050/lwjgl.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/lwjgl.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://cdn.discordapp.com/attachments/854426771621150741/854449782096134164/lwjgl64.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/lwjgl64.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://cdn.discordapp.com/attachments/854426771621150741/854449784114774016/OpenAL32.dll", False
	oXMLHTTP.Send
    oStream.Open
    oStream.Type = 1
    oStream.Write = oXMLHTTP.responseBody
    oStream.SaveToFile "./lib/native/OpenAL32.dll"
    oStream.Close
	Set oStream = CreateObject("ADODB.Stream")
	oXMLHTTP.Open "GET", "https://cdn.discordapp.com/attachments/854426771621150741/854449786228441108/OpenAL64.dll", False
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