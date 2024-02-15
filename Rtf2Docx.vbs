Const wdFormatDocumentDefault = 16
Dim wrdApp, wrdDoc, strFil
Set wrdApp = CreateObject("Word.Application")
For Each varArg in WScript.Arguments
    strFil = Replace(LCase(varArg), ".rtf", ".docx")
    Set wrdDoc = wrdApp.Documents.Add
    wrdDoc.SaveAs2 strFil, wdFormatDocumentDefault
    wrdDoc.Close False
    Set wrdDoc = Nothing 
Next
Set wrdApp = Nothing
