Attribute VB_Name = "save"
Option Explicit

Sub BPIbutton_Click()
Call get_data
Call datafixing

End Sub

Sub get_data()


    Dim rngCopyTo As Range
    Dim rngCopyFrom As Range
    Dim wbkProtected As Workbook
    
    Set rngCopyTo = Worksheets("Raw Data").Range("A1:AM10000")
    
    Set wbkProtected = Application.Workbooks.Open(Filename:="\\192.168.15.252\admins\ACTIVE\DATABASE\PSB DATABASE 2017.xlsm", Password:="")
    
    'Set wbkProtected = Application.Workbooks.Open(Filename:="F:\DATABASE\BPI DATABASE 2017.xlsb", Password:="BPIMADRID")
   
    Set rngCopyFrom = wbkProtected.Worksheets("DATABASE").Range("A1:AM10000")
    
    rngCopyTo.Value = rngCopyFrom.Value
    
    wbkProtected.Close savechanges:=xlDoNotSaveChanges

End Sub

Sub save()
Dim fpath, fname As String
Dim newbook As Workbook
Dim lastrow As Long

With ThisWorkbook.ActiveSheet
    lastrow = .Cells(.Rows.Count, "B").End(xlUp).Row - 1
End With


fpath = "\\192.168.15.252\admins\ACTIVE\BONNA\BPI\BPI DPD"
fname = "BPI DPD " & Format(Now(), "mmddyy") & ".xls"




ThisWorkbook.ActiveSheet.Range("A4:Q" & lastrow).Copy
Set newbook = Workbooks.Add
newbook.Worksheets("Sheet1").Range("A1").PasteSpecial (xlPasteValuesAndNumberFormats)
newbook.SaveAs Filename:=fpath & "\" & fname
newbook.Close
MsgBox ("Your file was save on " & fpath)
End Sub

Sub datafixing()
Dim lastrow, cell, activecell, i, mydata As Long

With Worksheets("Raw Data")
    lastrow = .Cells(.Rows.Count, "B").End(xlUp).Row - 1
End With

cell = 1
activecell = 1
For i = 1 To lastrow

cell = cell + 1
activecell = activecell + 1
Worksheets("Data").Range("A" & activecell).Value = Worksheets("Raw Data").Range("B" & cell).Value
Worksheets("Data").Range("B" & activecell).Value = "' 0" & Worksheets("Raw Data").Range("A" & cell).Value
Worksheets("Data").Range("C" & activecell).Value = Worksheets("Raw Data").Range("H" & cell).Value
Worksheets("Data").Range("D" & activecell).Value = Worksheets("Raw Data").Range("T" & cell).Value
Worksheets("Data").Range("E" & activecell).Value = Format(Worksheets("Raw Data").Range("S" & cell).Value, "#,##0.00")
Worksheets("Data").Range("F" & activecell).Value = Format(Worksheets("Raw Data").Range("L" & cell).Value, "#,##0.00")
Worksheets("Data").Range("G" & activecell).Value = Format(Worksheets("Raw Data").Range("E" & cell).Value, "d/m/yyyy")
Worksheets("Data").Range("H" & activecell).Value = Format(Worksheets("Raw Data").Range("AK" & cell).Value, "d/m/yyyy")
'Worksheets("Data").Range("I" & activecell).Value = ""
Worksheets("Data").Range("J" & activecell).Value = Worksheets("Raw Data").Range("F" & cell).Value
Worksheets("Data").Range("K" & activecell).Value = Worksheets("Raw Data").Range("K" & cell).Value
Worksheets("Data").Range("L" & activecell).Value = Worksheets("Raw Data").Range("I" & cell).Value
Worksheets("Data").Range("M" & activecell).Value = Worksheets("Raw Data").Range("AA" & cell).Value
Worksheets("Data").Range("N" & activecell).Value = Worksheets("Raw Data").Range("AB" & cell).Value
Worksheets("Data").Range("O" & activecell).Value = Worksheets("Raw Data").Range("C" & cell).Value
'Worksheets("Data").Range("P" & activecell).Value = ""
Worksheets("Data").Range("Q" & activecell).Value = "PSB"

Next i
End Sub



