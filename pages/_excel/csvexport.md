# Speichern von Dateien im CSV-Format mit "" und , 

Links: 
https://www.extendoffice.com/de/documents/excel/3620-excel-save-worksheet-data-as-csv-file-with-without-double-quotes.html
http://www.excel-ist-sexy.de/csv-export-mit-anfuehrungsstrichen/

## Anleitung
Alt F11 Drücken, 

dann Einfügen Modul

Folgenden Code

```
Option Explicit

Sub CSV_mit_Anfuehrungszeichen()
   Dim wks As Worksheet, Ze As Long, Sp As Long, ZeTmp As String
   Dim lCol As Long, lRow As Long, Frf As Long
   Const csvExport = "d:\inbox\csvExportSpezial.csv"   'Anpassen
   Const Trenner As String = ","    'Trenner für Spalten, kann angepasst werden
   Const Anf As String = """"
   
   Frf = FreeFile
   Set wks = ThisWorkbook.Worksheets("Tabelle1")  'Anpassen: Register-Name
   lCol = wks.Cells(1, Columns.Count).End(xlToLeft).Column
   lRow = wks.Cells(Rows.Count, 1).End(xlUp).Row
   
   Open csvExport For Output As #Frf
   For Ze = 1 To lRow
       For Sp = 1 To lCol
           ZeTmp = ZeTmp & Anf & CStr(wks.Cells(Ze, Sp).Text) & Anf & Trenner
       Next Sp
       ZeTmp = Left(ZeTmp, Len(ZeTmp) - 1)   'Letztes Trennzeichen löschen
       Print #Frf, ZeTmp
       ZeTmp = ""
   Next Ze
   Close #Frf
End Sub
```

```
Option Explicit

Sub CSV_mit_Anfuehrungszeichen()
   Dim wks As Worksheet, Ze As Long, Sp As Long, ZeTmp As String
   Dim lCol As Long, lRow As Long, Frf As Long
   Const csvExport = "d:\inbox\csvExportSpezial.csv"   'Anpassen
   Const Trenner As String = ","    'Trenner für Spalten, kann angepasst werden
   Const Anf As String = """"
   
   Frf = FreeFile
   Set wks = ThisWorkbook.Worksheets("Tabelle1")  'Anpassen: Register-Name
   lCol = wks.Cells(1, Columns.Count).End(xlToLeft).Column
   lRow = wks.Cells(Rows.Count, 1).End(xlUp).Row
   
   Open csvExport For Output As #Frf
   For Ze = 1 To lRow
       For Sp = 1 To lCol
           ZeTmp = ZeTmp & Anf & CStr(wks.Cells(Ze, Sp).Text) & Anf & Trenner
       Next Sp
       ZeTmp = Left(ZeTmp, Len(ZeTmp) - 1)   'Letztes Trennzeichen löschen
       Print #Frf, ZeTmp
       ZeTmp = ""
   Next Ze
   Close #Frf
End Sub
```


Dann F5 zum Ausführen


## FIX for utf-8
```

Option Explicit

Sub CSV_mit_Anfuehrungszeichen()
   Dim wks As Worksheet, Ze As Long, Sp As Long, ZeTmp As String
   Dim lCol As Long, lRow As Long, Frf As Long
   Const csvExport = "d:\inbox\csvExportSpezial.csv"   'Anpassen
   Const Trenner As String = ","    'Trenner für Spalten, kann angepasst werden
   Const Anf As String = """"
   
   Set wks = ThisWorkbook.Worksheets("Tabelle1")  'Anpassen: Register-Name
   lCol = wks.Cells(1, Columns.Count).End(xlToLeft).Column
   lRow = wks.Cells(Rows.Count, 1).End(xlUp).Row
   
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set out = fso.CreateTextFile(csvExport, True, True)  'Das hintere True bedeutet Unicode '
   
   For Ze = 1 To lRow
       For Sp = 1 To lCol
           ZeTmp = ZeTmp & Anf & CStr(wks.Cells(Ze, Sp).Text) & Anf & Trenner
       Next Sp
       ZeTmp = Left(ZeTmp, Len(ZeTmp) - 1)   'Letztes Trennzeichen löschen
       out.WriteLine(ZeTmp)
       ZeTmp = ""
   Next Ze
   out.close
End Sub
```


``` 