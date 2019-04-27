Option Explicit

Sub WriteLine ( strLine )
    WScript.Stdout.WriteLine strLine
End Sub

' http://msdn.microsoft.com/en-us/library/office/aa432714(v=office.12).aspx
Const msoFalse = 0   ' False.
Const msoTrue = -1   ' True.

Const xlTypePDF = 0
'
' This is the actual script
'

Dim inputFile
Dim outputFile
Dim objEXCEL
Dim objWorkBook
Dim objPrintOptions
Dim objFso

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

Set objEXCEL = CreateObject( "Excel.Application" )

objEXCEL.Visible = True
objEXCEL.WorkBooks.Open inputFile

Set objWorkBook = objEXCEL.ActiveWorkBook

' Reference for this at http://msdn.microsoft.com/en-us/library/office/ff746080.aspx
objWorkBook.ExportAsFixedFormat xlTypePDF, outputFile

objWorkBook.Close
objEXCEL.Quit