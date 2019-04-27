Option Explicit

Sub WriteLine ( strLine )
    WScript.Stdout.WriteLine strLine
End Sub

Const ppSaveAsPDF = 32
' http://msdn.microsoft.com/en-us/library/office/ff744228.aspx
Const ppShowAll = 1

Dim inputFile
Dim outputFile
Dim objPPT
Dim objPresentation
Dim objPrintOptions
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

Set objPPT = CreateObject( "PowerPoint.Application" )

objPPT.Visible = True
objPPT.Presentations.Open inputFile

Set objPresentation = objPPT.ActivePresentation
Set objPrintOptions = objPresentation.PrintOptions

objPrintOptions.Ranges.Add 1,objPresentation.Slides.Count
objPrintOptions.RangeType = ppShowAll

objPresentation.SaveAs outputFile, ppSaveAsPDF, True

objPresentation.Close
objPPT.Quit
