Rem Drag'n'Drop MS Office to PDF Converter
Rem Copyright (c) 2024 Petros Kyladitis
Dim app, doc, args, ext
Set args = WScript.Arguments
If(args.Count > 0) Then
	ext = Mid (args(0), InStrRev(args(0),".") + 1)
	Select Case Left(ext, 3)
	Case "doc"
		Set app = CreateObject("Word.Application")
		Set doc = app.Documents.Open(args(0))
		doc.SaveAs Left(args(0), Len(args(0)) - Len(ext)) & "pdf", 17
		doc.Close
		app.Quit
	Case "xls"
		Set app = CreateObject("Excel.Application")
		Set doc = app.Workbooks.Open(args(0))
		doc.ExportAsFixedFormat 0, Left(args(0), Len(args(0)) - Len(ext)) & "pdf"
		doc.Close
		app.Quit
	Case "ppt"
		Set app = CreateObject("PowerPoint.Application")
		Set doc = app.Presentations.Open(args(0))
		doc.SaveAs Left(args(0), Len(args(0)) - Len(ext)) & "pdf", 32
		doc.Close
		app.Quit
	Case Else
		MsgBox "The dropped file is not supported." & vbNewLine & "doc, docx, xls, xlsx, ppt & pptx files are acceptable." , vbExclamation, "Not supported file intput"
	End Select 
	Set doc = Nothing
	Set app = Nothing
Else
	MsgBox "Drag'n'Drop the the MS Office file you want to make PDF on this script", vbInformation, "No file input"
End If