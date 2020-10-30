
Dim Fso,fpwoExt,genfilePath,oExcelFileRead,outJsonValue,jsonKeyVal
Dim iterIndx,iDictr

set fso = CreateObject("Scripting.FileSystemObject")
scriptdir = fso.GetParentFolderName(WScript.ScriptFullName) ' Get parent direcrory
Set oFldr=fso.GetFolder(scriptdir) 
Set oFiles=oFldr.Files ' Get count of files in folder

On Error Resume Next


iterIndx=0
Dim oExcel ,objWorkbook, objSheet
Set oExcel = CreateObject("Excel.Application")
	'Set oExcel = GetObject(, "Excel.Application")
	oExcel.Visible = True
	oExcel.DisplayAlerts = False
	oExcel.Workbooks.Add
	Set objSheet = oExcel.ActiveWorkbook.Worksheets(1)
	objSheet.Name = "Automation"

Set oDict = CreateObject("Scripting.Dictionary")



For Each item In oFiles
iterIndx=iterIndx+1
'msgbox LCase(item.type)
If LCase(item.type) = "json file" Then
jsonfilePath=scriptdir+"\"+item.name
fpwoExt=Split(jsonfilePath,".")(0) ' changing the file extension
genfilePath=fpwoExt+".txt"
'============================================================
'Rename json to text file 
'============================================================
 fso.MoveFile jsonfilePath, genfilePath
'============================================================
'Reading file
'============================================================
elseif (LCase(item.type) = "text document") Then
genfilePath=scriptdir+"\"+item.name 'assuming text file
elseif (LCase(item.type) = "vbscript script file") Then
genfilePath="NA" 'assuming vbscript file
End if

if (genfilePath <> "NA") Then
	'msgbox genfilePath
	
	Set oFileRead = fso.OpenTextFile(genfilePath, 1) 
	outJsonValue=oFileRead.ReadAll     
	'msgbox outJsonValue  
	oDict.RemoveAll	
'============================================================
'Process JSON out to Dictionary
'============================================================
'Split(Replace(Replace(outJsonValue,"{",""),"}",""),",")
	jsonKeyVal=Split( Replace(Replace(Replace(Replace(outJsonValue,"{",""),"}",""),"[",""),"]","") ,",")
		For i=0 to ubound(jsonKeyVal)
		oDict.Add i+1,jsonKeyVal(i)
		Next
	'Wscript.Echo +oDict.Count
	
'============================================================
'Split dictionary value and out in excel
'============================================================
'Call WriteDictValuesToExcel(iterIndx,oExcel,scriptdir+"\CombinedExcel.xlsx",oDict)

iDictr=1
For each val in oDict.keys
    'Msgbox "Key: " & obj & " Value: " & oDict(obj)
	objSheet.Cells(iDictr, iterIndx).Value = oDict(val)
	iDictr=iDictr+1
Next

 
	
End if 
 
 
 
 
Next

	oExcel.ActiveWorkbook.SaveAs scriptdir+"\JsonToExcelMapper.xlsx"
	oExcel.ActiveWorkbook.Close
	oExcel.Application.Quit
	Set objSheet=nothing
	Set oExcel=nothing 
	'oDict.RemoveAll
	Set oDict = nothing

Wscript.Echo "Test Finished"
'Script End 

'============================================================
'Split dictionary value and out in excel, Create function
'============================================================

Function WriteDictValuesToExcel(iterIndx,excelObject,excelPath,jsonDictionary)
'WScript.Echo "Test Function"
Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = True
Set objWorkbook = oExcel.Workbooks.Add
iDictr=1
For Each item in jsonDictionary.keys
    'Msgbox "Key: " & obj & " Value: " & oDic(obj)
	objWorkbook.Cells(iDictr, iterIndx).Value = jsonDictionary(item)
	iDictr=iDictr+1
	'msgbox iDictr
Next

objWorkbook.SaveAs excelPath
objWorkbook.Close
oExcel.Quit
WScript.Echo "Finished."
'WScript.Quit
End Function