Dim objuft
Set objuft=CreateObject("QuickTest.Application")
objuft.visible=True
objuft.launch
objuft.open("‪C:\Keyword driven folder\Driver\Driver")
objuft.Test.run
objuft.Test.Close
objuft.quit
set objuft=nothing