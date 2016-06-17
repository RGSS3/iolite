var dl = new ActiveXObject("msxml2.XMLHTTP");
dl.open("GET", "https://github.com/aria2/aria2/releases/download/release-1.24.0/aria2-1.24.0-win-64bit-build1.zip", 0)
dl.send()
WScript.Echo("Downloading");
var write = new ActiveXObject("ADODB.stream");
write.Mode = 3
write.Type = 1
write.Open()
write.Write(dl.responseBody);
write.SaveToFile("aria2c.zip", 2);
WScript.Echo("Decompressing");

var shell = new ActiveXObject("Shell.Application")
var zip = shell.NameSpace("aria2c.zip");
var dir = shell.NameSpace(".");
dir.copyhere(zip.Items(), 256)

