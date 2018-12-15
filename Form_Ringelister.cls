VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Ringelister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database




Private Sub Combo28_AfterUpdate()
Oppdatere_driten
End Sub

Private Sub Combo3_Click()

Dim filter_streng As String
Dim lngLen As Long
Dim dato_for_ukestart As Date

dato_for_ukestart = WeekStart(Text9.Value, Text15.Value)

If (Text9.Value Mod 2 = 0) = True Then
    Combo11.Value = "Partallsuke"
Else
    Combo11.Value = "Oddetallsuke"
End If

If Not Combo3.Value = "" Then
    filter_streng = filter_streng & "Område = '" & Combo3.Value & "' AND "
End If

If Not List50.Value = "" Then
    filter_streng = filter_streng & "Teleselger = '" & List50.Value & "' AND "
End If

If Not Combo7.Value = "Alle" Then
    If Combo7.Value = "Mandag" Then
        filter_streng = filter_streng & "[Bestillingsdag: Mandag] = TRUE AND "
    ElseIf Combo7.Value = "Tirsdag" Then
        filter_streng = filter_streng & "[Bestillingsdag: Tirsdag] = TRUE AND "
        dato_for_ukestart = dato_for_ukestart + 1
    ElseIf Combo7.Value = "Onsdag" Then
        filter_streng = filter_streng & "[Bestillingsdag: Onsdag] = TRUE AND "
        dato_for_ukestart = dato_for_ukestart + 2
    ElseIf Combo7.Value = "Torsdag" Then
        filter_streng = filter_streng & "[Bestillingsdag: Torsdag] = TRUE AND "
        dato_for_ukestart = dato_for_ukestart + 3
    ElseIf Combo7.Value = "Fredag" Then
        filter_streng = filter_streng & "[Bestillingsdag: Fredag] = TRUE AND "
        dato_for_ukestart = dato_for_ukestart + 4
    ElseIf Combo7.Value = "Lørdag" Then
        filter_streng = filter_streng & "[Bestillingsdag: Lørdag] = TRUE AND "
        dato_for_ukestart = dato_for_ukestart + 5
    ElseIf Combo7.Value = "Søndag" Then
        filter_streng = filter_streng & "[Bestillingsdag: Søndag] = TRUE AND "
        dato_for_ukestart = dato_for_ukestart + 6
    End If
End If

Text13.Value = dato_for_ukestart

If Combo11.Value = "Partallsuke" Then
    filter_streng = filter_streng & "([Bestillingsuker] = 'Partallsuker' OR [Bestillingsuker] = 'Alle uker' OR [Bestillingsuker] = 'Annet (spesifiser)') AND "
ElseIf Combo11.Value = "Oddetallsuke" Then
    filter_streng = filter_streng & "([Bestillingsuker] = 'Oddetallsuker' OR [Bestillingsuker] = 'Alle uker' OR [Bestillingsuker] = 'Annet (spesifiser)') AND "
End If

lngLen = Len(filter_streng) - 5 'Without trailing " AND ".

'Call sjekke_calls(Left$(filter_streng, lngLen), Text13.Value)
Call sjekke_calls(filter_streng & "([Bestillingsuker (annet)] Like '*" & Text9.Value & ",*' OR [Bestillingsuker (annet)] = 'N/A')", Text13.Value)

If Not Text13.Value = "" Then
    filter_streng = filter_streng & "[Ringedag] = '" & Text13.Value & "' AND "
End If

If Not Text30.Value = "" Then
    filter_streng = filter_streng & "[Kundenummer] = '" & Text30.Value & "' AND "
End If

If Not Text37.Value = "" Then
    filter_streng = filter_streng & "[Telefonnummer / SMS nummer] = '" & Text37.Value & "' AND "
End If

If Not Combo28.Value = "" Then
    filter_streng = filter_streng & "[Resultat] = '" & Combo28.Value & "' AND "
End If

filter_streng = filter_streng & "([Bestillingsuker (annet)] Like '*" & Text9.Value & ",*' OR [Bestillingsuker (annet)] = 'N/A') AND "
lngLen = Len(filter_streng) - 5 'Without trailing " AND ".

If lngLen > 0 Then
    filter_streng = Left$(filter_streng, lngLen)
    Me.Kunderegister.Form.Filter = filter_streng
    Me.Kunderegister.Form.FilterOn = True
Else
    Me.Kunderegister.Form.Filter = ""
    Me.Kunderegister.Form.FilterOn = False
End If

End Sub



Private Sub Command36_Click()

' Denne rutinen åpner excel og putter dataene inn der.
' Den tar hensyn til filtere som er satt i comboboksene.

'Dim strWhere As String

'Dim filter_streng As String
'Dim lngLen As Long
'Dim dato_for_ukestart As Date

Dim db As DAO.Database
Dim rs2 As DAO.Recordset
Dim rs As DAO.Recordset

Set db = CurrentDb
'Set rs2 = db.OpenRecordset("Query_ringelister", dbOpenSnapshot)
Set rs2 = Me.Kunderegister.Form.RecordsetClone

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





Private Sub Command49_Click()
Dim k_nr As String
Dim funnet_dato As Date
'Her søker vi etter kunde. Putter den funnede leveringsdatoen inn i ringedagfeltet. Putter også kundenummeret inn i kundenummeret feltet.

k_nr = finne_kunde
funnet_dato = finne_bestillingsdag(k_nr, DateValue(Now))

If Not funnet_dato = "01.01.1900" Then
    Text13.Value = funnet_dato
    Text30.Value = k_nr
    Oppdatere_driten
End If

End Sub

Private Sub Form_Load()

Dim i As Integer


For i = 1 To Text13.ListCount
    'Remove an item from the ListBox.
    Text13.RemoveItem 0

Next i


For i = -10 To 20
        If Not WeekDay(Now() + i, vbMonday) = 6 And Not WeekDay(Now() + i, vbMonday) = 7 Then
            ' oppdaterer listeboksen med datoer før og datoer etter dagens dato.
            Text13.AddItem (DateValue(Now) + i)
        End If
Next i


'Text13.SetFocus
Command17_Click
End Sub



Private Sub List50_Click()
Combo3_Click
End Sub

Private Sub Text13_AfterUpdate()
Oppdatere_driten
End Sub

Private Sub Oppdatere_driten()
'Her er koden for ringedag tekstboksen

Select Case WeekDay(Text13.Value, vbMonday)
    Case 1
        Combo7.Value = "Mandag"
    Case 2
        Combo7.Value = "Tirsdag"
    Case 3
        Combo7.Value = "Onsdag"
    Case 4
        Combo7.Value = "Torsdag"
    Case 5
        Combo7.Value = "Fredag"
    Case 6
        Combo7.Value = "Lørdag"
    Case 7
        Combo7.Value = "Søndag"
End Select
If Not Text13.Value = "" Then
    Text9.Value = Format(ISOWeekNum(Text13.Value), "00")
    Text15.Value = Year(Text13.Value)
Else
    Text9.Value = Format(ISOWeekNum(Now()), "00")
    Text15.Value = Year(Now())
End If

Combo3_Click
End Sub



Private Sub Command17_Click()
Text13.Value = Now()
Oppdatere_driten
End Sub

Private Sub Command18_Click()
On Error GoTo Err_Command18_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Registrere_kunde"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command18_Click:
    Exit Sub

Err_Command18_Click:
    MsgBox Err.Description
    Resume Exit_Command18_Click
    
End Sub
Private Sub Command19_Click()
On Error GoTo Err_Command19_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Kunderegister_alle_redigere"
    DoCmd.OpenForm stDocName, 3 ', , stLinkCriteria

Exit_Command19_Click:
    Exit Sub

Err_Command19_Click:
    MsgBox Err.Description
    Resume Exit_Command19_Click
    
End Sub



Private Sub Text30_Exit(Cancel As Integer)
Combo3_Click
End Sub

Private Sub Text37_Exit(Cancel As Integer)
Combo3_Click
End Sub

Private Sub ta_ut_hele_uka()

' Denne rutinen åpner excel og putter dataene inn der.
' Den tar hensyn til filtere som er satt i comboboksene.

Dim strWhere As String

Dim filter_streng As String
Dim lngLen As Long
Dim dato_for_ukestart As Date
Dim dato_for_ukeslutt As Date

dato_for_ukestart = WeekStart(Text9.Value, Text15.Value)
dato_for_ukeslutt = WeekStart(Text9.Value, Text15.Value) + 7


If (Text9.Value Mod 2 = 0) = True Then
    Combo11.Value = "Partallsuke"
Else
    Combo11.Value = "Oddetallsuke"
End If

If Not Combo3.Value = "" Then
    filter_streng = filter_streng & "Område = '" & Combo3.Value & "' AND "
End If

If Not List50.Value = "" Then
    filter_streng = filter_streng & "Teleselger = '" & List50.Value & "' AND "
End If

If Not Combo7.Value = "Alle" Then
    If Combo7.Value = "Mandag" Then
        filter_streng = filter_streng & "[Bestillingsdag: Mandag] = TRUE AND "
    ElseIf Combo7.Value = "Tirsdag" Then
        filter_streng = filter_streng & "[Bestillingsdag: Tirsdag] = TRUE AND "
        dato_for_ukestart = dato_for_ukestart + 1
    ElseIf Combo7.Value = "Onsdag" Then
        filter_streng = filter_streng & "[Bestillingsdag: Onsdag] = TRUE AND "
        dato_for_ukestart = dato_for_ukestart + 2
    ElseIf Combo7.Value = "Torsdag" Then
        filter_streng = filter_streng & "[Bestillingsdag: Torsdag] = TRUE AND "
        dato_for_ukestart = dato_for_ukestart + 3
    ElseIf Combo7.Value = "Fredag" Then
        filter_streng = filter_streng & "[Bestillingsdag: Fredag] = TRUE AND "
        dato_for_ukestart = dato_for_ukestart + 4
    ElseIf Combo7.Value = "Lørdag" Then
        filter_streng = filter_streng & "[Bestillingsdag: Lørdag] = TRUE AND "
        dato_for_ukestart = dato_for_ukestart + 5
    ElseIf Combo7.Value = "Søndag" Then
        filter_streng = filter_streng & "[Bestillingsdag: Søndag] = TRUE AND "
        dato_for_ukestart = dato_for_ukestart + 6
    End If
End If

Text13.Value = dato_for_ukestart

If Combo11.Value = "Partallsuke" Then
    filter_streng = filter_streng & "([Bestillingsuker] = 'Partallsuker' OR [Bestillingsuker] = 'Alle uker' OR [Bestillingsuker] = 'Annet (spesifiser)') AND "
ElseIf Combo11.Value = "Oddetallsuke" Then
    filter_streng = filter_streng & "([Bestillingsuker] = 'Oddetallsuker' OR [Bestillingsuker] = 'Alle uker' OR [Bestillingsuker] = 'Annet (spesifiser)') AND "
End If

lngLen = Len(filter_streng) - 5 'Without trailing " AND ".

Call sjekke_calls(Left$(filter_streng, lngLen), Text13.Value)

If Not Text13.Value = "" Then
    filter_streng = filter_streng & "[Ringedag] = '" & Text13.Value & "' AND "
End If

If Not Text30.Value = "" Then
    filter_streng = filter_streng & "[Kundenummer] = '" & Text30.Value & "' AND "
End If

If Not Text37.Value = "" Then
    filter_streng = filter_streng & "[Telefonnummer / SMS nummer] = '" & Text37.Value & "' AND "
End If

If Not Combo28.Value = "" Then
    filter_streng = filter_streng & "[Resultat] = '" & Combo28.Value & "' AND "
End If

filter_streng = filter_streng & "([Bestillingsuker (annet)] Like '*" & Text9.Value & ",*' OR [Bestillingsuker (annet)] = 'N/A') AND "
lngLen = Len(filter_streng) - 5 'Without trailing " AND ".

'If Check39.Value = True Then
'    filter_streng = ""
'    If Not Text30.Value = "" Then
'        filter_streng = filter_streng & "[Kundenummer] = '" & Text30.Value & "' AND "
'    End If
'
'    If Not Text37.Value = "" Then
'        filter_streng = filter_streng & "[Telefonnummer / SMS nummer] = '" & Text37.Value & "' AND "
'    End If
    
'    lngLen = Len(filter_streng) - 5 'Without trailing " AND ".

'End If

Dim db As DAO.Database
Dim rs2 As DAO.Recordset
Dim rs As DAO.Recordset

Set db = CurrentDb
Set rs2 = db.OpenRecordset("Query_ringelister", dbOpenSnapshot)

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


