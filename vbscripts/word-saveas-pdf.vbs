Option Explicit

Sub WriteLine ( strLine )
    WScript.Stdout.WriteLine strLine
End Sub

Const wdFormatPDF = 17
' http://msdn.microsoft.com/en-us/library/office/ff744228.aspx

Dim inputFile
Dim outputFile
Dim objWORD
Dim objDocument
Dim objFso
Dim pptf

If WScript.Arguments.Count <> 2 Then
    WriteLine "You need to specify input and output files."
    WScript.Quit
End If

inputFile = WScript.Arguments(0)
outputFile = WScript.Arguments(1)

Set objFso = CreateObject("Scripting.FileSystemObject")

If Not objFso.FileExists( inputFile ) Then
    WriteLine "Unable to find your input file " & inputFile
    WScript.Quit
End If

If objFso.FileExists( outputFile ) Then
    WriteLine "Your output file (' & outputFile & ') already exists!"
    WScript.Quit
End If

WriteLine "Input File:  " & inputFile
WriteLine "Output File: " & outputFile

Set objWORD = CreateObject( "Word.Application" )

objWORD.Visible = True
objWORD.Documents.Open inputFile

Set objDocument = objWORD.ActiveDocument

objDocument.SaveAs outputFile, wdFormatPDF

objDocument.Close
objWORD.Quit
