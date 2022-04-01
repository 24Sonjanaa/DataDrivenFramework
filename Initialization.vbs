Dim objuft

Set objuft=CreateObject("QuickTest.Application")
objuft.visible=True
objuft.launch
objuft.open("C:\Users\Sanjana\Desktop\Lesson-13_test3")
objuft.Test.Run
objuft.Test.Close
objuft.quit
set objuft=nothing