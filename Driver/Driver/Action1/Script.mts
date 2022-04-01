
'Datatable.ImportSheet "C:\Keyword driven folder\Organizer\Organizer.xlsx",1,"Module"
Services.StartTransaction "Transaction"
'mrowcount=datatable.GetSheet("Module").GetRowCount
'msgbox mrowcount

'For  i = 1 To mrowcount Step 1
	
'Datatable.SetCurrentRow(i)

'Modexe=Datatable("ModuleExecution","Module")

'msgbox Modexe
'If modexe="Y" Then
	
	'Modid=Datatable("ModuleID","Module")
	
	'msgbox Modid
'End If

'Next

'Datatable.AddSheet "Module"
'Datatable.ImportSheet "E:\capgemini\KeywordDrivenFramework\Organizer\organizer.xlsx",1,"Module"
mrowcount=datatable.GetSheet("Action1").GetRowCount
msgbox mrowcount

For  i = 1 To mrowcount Step 1
	
Datatable.SetCurrentRow(i)

Modexe=Datatable("ModuleExecution","Action1")

msgbox Modexe
If modexe="Y" Then
	
	Modid=Datatable("ModuleID","Action1")
	
	msgbox Modid
	
	trowcount=datatable.GetSheet("Action2").GetRowCount
	
	msgbox trowcount
	
	For j=1  To trowcount Step 1
	Datatable.SetCurrentRow(j)
	If Modid=Datatable("ModuleID","Action2") and Datatable("TestCaseExecution","Action2")="Y" Then
	testcaseid=Datatable("TestCase_ID","Action2")
	msgbox testcaseid	
	tsrowcount=Datatable.GetSheet("Action3").GetRowCount
	msgbox tsrowcount
	
	For k = 1 to tsrowcount Step 1
	
	datatable.SetCurrentRow(k)
	
	If testcaseid=Datatable("TestCase_ID","Action3") Then
		
	keyword=Datatable("Keywords","Action3")
    msgbox keyword    
    
    select case(keyword)
    	
    Case "In"
    Call Login()
    
    Case "ca"
    Call Closeapp()
    
    Case "oo"
    Call OpenOrder()
    
    Case "uo"
    Call UpdateOrder()
    
    End Select
    		
End  If
Next
End  If
Next
End  If
Next

Services.EndTransaction "Transaction"


