Attribute VB_Name = "Module4"
Option Compare Database

Private Sub Ny_eksport_rutine()

' Denne rutinen åpner excel og putter dataene inn der.
' Den tar hensyn til filtere som er satt i comboboksene.

Dim strWhere As String

Dim filter_streng As String
Dim lngLen As Long
Dim dato_for_ukestart As Date

Dim db As DAO.Database
Dim rs2 As DAO.Recordset
Dim rs As DAO.Recordset

Set db = CurrentDb
'Set rs2 = db.OpenRecordset("Query_ringelister", dbOpenSnapshot)
'Set rs2 = Me.Kunderegister.Form.RecordsetClone

'Lage en ny excel workbook
Dim oApp As New Excel.Application
Dim oBook As Excel.Workbook
Dim oSheet As Excel.Worksheet

Set oBook = oApp.Workbooks.Add
Set oSheet = oBook.Worksheets(1)

If lngLen > 0 Then
    strWhere = Left$(filter_streng, lngLen)
    rs2.Filter = strWhere
Else
    rs2.Filter = ""
End If

Set rs = rs2.OpenRecordset ' kopierer recordsettet, ellers fungerer ikke filteret

'Putt inn feltnavn i rad 1
Dim i As Integer
Dim iNumCols As Integer
iNumCols = rs.Fields.Count

For i = 1 To iNumCols
    oSheet.Cells(1, i).Value = rs.Fields(i - 1).Name
Next

'Putt inn data fra og med A2
oSheet.Range("A2").CopyFromRecordset rs

'Header bold og juster alle kolonner
With oSheet.Range("a1").Resize(1, iNumCols)
    .Font.Bold = True
    .EntireColumn.AutoFit
End With

oApp.Visible = True
oApp.UserControl = True

'Lukke alt
rs.Close
rs2.Close
db.Close
    

End Sub

