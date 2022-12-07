'On Error Resume Next
Function Convert(sPath)
	'遍历一个文件夹下的所有文件夹文件夹
	Const wdExportFormatPDF = 17
	Set oWord = WScript.CreateObject("Word.Application")
	Set oFso = CreateObject("Scripting.FileSystemObject")
	Set oFolder = oFso.GetFolder(sPath)
	Set oSubFolders = oFolder.SubFolders

	For Each oSubFolder In oSubFolders
		'WScript.Echo oSubFolder.Path
		'oSubFolder.Delete
		'MsgBox oSubFolder.Path
		Convert(oSubFolder.Path)'递归
	Next

	Set oFiles = oFolder.Files
	For Each oFile In oFiles
		'MsgBox oFile.Path
		'WScript.Echo oFile.Path
		If (LCase(Right(oFile.Name,4))=".doc" Or LCase(Right(oFile.Name,4))="docx" ) And Left(oFile.Name,1)<>"~" Then
			Set oDoc=oWord.Documents.Open(oFile.Path)
			odoc.ExportAsFixedFormat Left(oFile.Path,InStrRev(oFile.Path,"."))&"pdf",wdExportFormatPDF
			If Err.Number Then
			MsgBox Err.Description
			End If
		End If
	Next

	oword.Quit 0
	Set oDoc=Nothing
	Set oWord =Nothing

	Set oFolder = Nothing
	Set oSubFolders = Nothing
	Set oFso = Nothing
End Function

Convert(".") '遍历
MsgBox "Word文件已全部轩换为PDF格式!"