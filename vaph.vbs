'On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set Folder = fso.GetFolder(fso.GetAbsolutePathName("."))
Set Files = Folder.Files
Dim app
Set app = createobject("Excel.Application")
set Workbook = app.workbooks.add()
set WorkSheet = Workbook.Worksheets(1)
currentRow = 1
xlTop = -4160

For Each File in Files
	If Instr(fso.GetExtensionName(File), "xls") Then
		Wscript.Echo File.name
		Opleg fso.GetAbsolutePathName(File)
		Wscript.Echo "Done"
		Wscript.echo
	End If
Next
outputwb = fso.GetAbsolutePathName(".") & "\Overzicht.xlsx"
if fso.FileExists(outputwb) Then  
    fso.DeleteFile outputwb
end if 
WorkSheet.Cells.VerticalAlignment = xlTop
WorkSheet.Cells.WrapText = True
WorkSheet.Cells.AutoFilter
Worksheet.Columns(1).NumberFormat = "0"
Worksheet.Columns(2).ColumnWidth = 40
Worksheet.Columns(2).Hidden = True
Worksheet.Columns(15).ColumnWidth = 120
Worksheet.Columns(4).NumberFormat = "0"
WorkSheet.Rows.AutoFit

Workbook.Saveas(outputwb)
WScript.StdOut.WriteLine "Output:" & outputwb

app.quit

Sub Opleg(file)
	Set wb = app.Workbooks.Open(file,, True)
	Dim Fiche
	Set Fiche = Nothing
	For Each ws In wb.Worksheets
		If ws.Name = "fiche" Then
			Set Fiche = ws
			Exit For
		End If
	Next
	
	If Fiche is Nothing Then
		WScript.StdOut.WriteLine "Geen fiche gevonden."
		wb.close(False)
		Exit Sub
	End If
	
	row = 0
	Dim productcel
	Set productcel = Nothing
	Dim oplegcel
	Set oplegcel = Nothing
	Dim indelingcel
	Set indelingcel = Nothing
	Dim voorstelcel
	Set voorstelcel = Nothing
	For Each rw In Fiche.Rows
		Set oplegcel = rw.Find("opleg")
		If Not oplegcel is Nothing Then
		    Set productcel = rw.Find("productnr")
			Set indelingcel = rw.Find("indeling")
			Set voorstelcel = rw.Find("voorstel")
			Exit For
		End If
		If row > 10 Then
			Exit For
		End If
		row = row + 1
	Next 
	For i = 2 to 15
    	WorkSheet.cells(1, i).value = Fiche.Cells(oplegcel.row, i)
	Next
				
	If productcel is Nothing Then
		WScript.StdOut.WriteLine "Geen informatie over product gevonden."
		wb.close(False)
		Exit Sub
	End If
	If oplegcel is Nothing Then
		WScript.StdOut.WriteLine "Geen informatie over gevraagde opleg gevonden."
		wb.close(False)
		Exit Sub
	End If
	If indelingcel is Nothing Then
		WScript.StdOut.WriteLine "Geen informatie over indeling gevonden."
		wb.close(False)
		Exit Sub
	End If
	If voorstelcel is Nothing Then
		WScript.StdOut.WriteLine "Geen informatie over voorstel gevonden."
		wb.close(False)
		Exit Sub
	End If
	' Basisproduct opzoeken
	basisproduct = "?"
	For Each cel in productcel.entirecolumn.cells
	    If Instr(Fiche.Cells(cel.row, 1).value, "Productomschrijving") Then
		    basisproduct = cel.value
			Exit For
		End If
		If cel.row > 10 Then
		    Exit For
		End If	
	Next

	' Rijen met opleg kopieren naar resultaat spreadsheet
	Dim oplegcol
	Set oplegcol = oplegcel.entirecolumn
	Dim emptycells 
	emptycells = 0
	For Each cel in oplegcol.cells
		If cel.value = "" Then
			emptycells = emptycells + 1
		Else 
			emptycells = 0
		End If 
		If emptycells > 100 Then
			Exit For
		End If
		If IsNumeric(cel.value) Then
			If cel.value > 0 and cel.Interior.ColorIndex = 2 Then
				currentRow = currentRow + 1
				WorkSheet.cells(currentRow, 1).value = basisproduct
				For i = 2 to 15
    				WorkSheet.cells(currentRow, i).value = Fiche.Cells(cel.row, i)
				Next
				WorkSheet.cells(currentRow, 16).value = file
			End If
		End If
	Next
	wb.close(False)
End Sub



