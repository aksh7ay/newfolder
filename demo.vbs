call copy_sheets
Sub copy_sheets()
    Dim eapp
    Dim wkbk_from
    Dim wkbk_to
    Dim wksh
    Dim wbfound
	dim oShell 
	
dim sPath 
    wbfound = False

    Set eapp = GetObject(, "Excel.Application")
    
	Set oShell = CreateObject("WScript.Shell")
	sPath = Replace(oShell.SpecialFolders("MYDocuments"), "Documents", "Desktop") & "/Add Member SSN's.xlsx"	
    

    For Each wksh In eapp.Workbooks
    MsgBox wksh.Name
       If wksh.Name = "Add Member SSN's.xlsx" Then
        wksh.Activate
        
    End If
       
    MsgBox "hi"
    Next

If wbfound = False Then
    eapp.Workbooks.Open (sPath)
    eapp.Visible = True

End If



	myvalue = ""

    	i = 2


    	Do Until Trim(eapp.Worksheets("Sheet1").Cells(i, 5)) = ""

    				For j = 1 To 24

        				myvalue = myvalue & eapp.Worksheets("Sheet1").Cells(i, j).Value & ","
					MsgBox myvalue

    				Next

    				i = i + 1

    	Loop


    	'MsgBox myvalue
    	'wscript.echo myvalue






    
    
        
   MsgBox "hello"
    
End Sub

