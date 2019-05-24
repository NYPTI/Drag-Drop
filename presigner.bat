signtool sign /tr http://timestamp.digicert.com /f "../Code Signing Certificate.pfx" ./BlankWindow/HelloWorld/bin/Release/HelloWorld.exe
signtool sign /tr http://timestamp.digicert.com /f "../Code Signing Certificate.pfx" ./NYPTIOutlookDragDrop/bin/Release/EasyHook.dll ./NYPTIOutlookDragDrop/bin/Release/log4net.dll "./NYPTIOutlookDragDrop/bin/Release/NYPTI Outlook Drag Drop.dll"
PAUSE