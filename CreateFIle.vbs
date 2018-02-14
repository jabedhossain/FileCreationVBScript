Set doShell = CreateObject("Wscript.Shell")
'Find the  current path 
dsPath = doShell.CurrentDirectory
'Get Current Date 
ddCurrentDate = Date
'Adjusting Padding in the Date Field  
if(Len(DatePart("d",ddCurrentDate)) = 1) Then
diFirstSeparatorPositon = InStr(ddCurrentDate,"/")
ddCurrentDate = Left(ddCurrentDate,diFirstSeparatorPositon)+ "0" +  Right(ddCurrentDate,Len(ddCurrentDate)-diFirstSeparatorPositon) 	
End If 	
'Create File Name 
dsFileName = dsPath + "\Excel_" + Replace(ddCurrentDate,"/","-") +".xlsx"
Set doFileSystem = CreateObject("Scripting.FileSystemObject")
'Check if the file already exists 
if (doFileSystem.FileExists(dsFileName)) Then
dsMsg = msgbox(dsFileName & "Already Exists",0,"Message" )
Wscript.Quit
end if 
'Object to create Excel based application
Set doExcel = CreateObject("Excel.Application")
doExcel.Visible = false
'Object to create excel file from excel application
Set doWorkBook = doExcel.Workbooks.Add()
'Save the excel file with the created file name
doWorkBook.SaveAs(dsFileName)
doExcel.Quit
'x = msgbox(ddCurrentDate,0,"Check")
