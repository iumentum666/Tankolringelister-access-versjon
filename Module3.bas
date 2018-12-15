Attribute VB_Name = "Module3"
Option Compare Database

Public Function finne_kunde() As String

    Dim henvendelse As Long
    On Error GoTo ErrHandler:
    ' Open the form in dialog mode.
    ' This "halts" execution until the called form is closed or hidden
    DoCmd.OpenForm "Frm_Søke_på_kunder", WindowMode:=acDialog
    
    ' Since the form was opened as a dialog, we will not reach this line until it is hidden.
    ' Here, we will retrieve the value in the password text box.
    henvendelse = Forms!Frm_Søke_på_kunder.Kundenr & vbNullString
    
    ' Now, we will actually close the password form
    DoCmd.Close acForm, "Frm_Søke_på_kunder"
    
    ' And finally, we return the value we retrieved
    finne_kunde = henvendelse
ErrHandler:
End Function

Public Function finne_bestillingsdag(k_nr As String, dato As Date) As Date

' Denne funksjonen finner bestillingsdagen på kunden. Den er en kopi av finne leveringsdagsfunksjonen, derfor har alle variablene navnene derfra.

Dim rsKundedata As DAO.Recordset
Dim mandag As Boolean
Dim tirsdag As Boolean
Dim onsdag As Boolean
Dim torsdag As Boolean
Dim fredag As Boolean
Dim best_dag As Integer
Dim lev_dag As Integer
Dim lev_dag2 As Integer
Dim funnet_leveringsdag As Integer
Dim lev_dato As Date
Dim lev_dato2 As Date
Dim Bestillingsuker As String
Dim Bestillingsuker_annet As String
Dim i As Integer, f As Integer

Set rsKundedata = CurrentDb.OpenRecordset("Select * FROM kunderegister Where Kundenummer ='" & k_nr & "'")

mandag = rsKundedata![Bestillingsdag: Mandag]
tirsdag = rsKundedata![Bestillingsdag: Tirsdag]
onsdag = rsKundedata![Bestillingsdag: Onsdag]
torsdag = rsKundedata![Bestillingsdag: Torsdag]
fredag = rsKundedata![Bestillingsdag: Fredag]

Bestillingsuker = rsKundedata![Bestillingsuker]
Bestillingsuker_annet = rsKundedata![Bestillingsuker (annet)]

funnet_leveringsdag = 0

If mandag = True Then
    lev_dag = 1
    funnet_leveringsdag = funnet_leveringsdag + 1
End If

If tirsdag = True Then
    If funnet_leveringsdag > 0 Then
        lev_dag2 = 2
    Else
        lev_dag = 2
    End If
    funnet_leveringsdag = funnet_leveringsdag + 1
End If

If onsdag = True Then
    If funnet_leveringsdag > 0 Then
        lev_dag2 = 3
    Else
        lev_dag = 3
    End If
    funnet_leveringsdag = funnet_leveringsdag + 1
End If

If torsdag = True Then
    If funnet_leveringsdag > 0 Then
        lev_dag2 = 4
    Else
        lev_dag = 4
    End If
    funnet_leveringsdag = funnet_leveringsdag + 1
End If

If fredag = True Then
    If funnet_leveringsdag > 0 Then
        lev_dag2 = 5
    Else
        lev_dag = 5
    End If
    funnet_leveringsdag = funnet_leveringsdag + 1
End If

rsKundedata.Close

If funnet_leveringsdag = 0 Then
    'Sjekker om bestillingsdag er funnet
    MsgBox ("Ingen bestillingsdag er funnet")
    finne_bestillingsdag = DateValue("01.01.1900")
    Exit Function
Else
    best_dag = WeekDay(dato, vbMonday)
End If

If best_dag <= lev_dag Then
    lev_dato = dato + (lev_dag - best_dag)
ElseIf lev_dag <= best_dag Then
    lev_dato = dato + ((lev_dag - best_dag) + 7)
End If

If funnet_leveringsdag > 1 Then
    If best_dag <= lev_dag2 Then
        lev_dato2 = dato + (lev_dag2 - best_dag)
    ElseIf lev_dag2 <= best_dag Then
        lev_dato2 = dato + ((lev_dag2 - best_dag) + 7)
    End If

    If lev_dato2 < lev_dato Then
        lev_dato = lev_dato2
    End If
End If



' lev_dato er kundens leveringsdato
' rutinen fungerer ikke hvis kunden har mer enn 2 leveringsdatoer

' Under her så sjekker vi om den funnede datoen er innenfor de angitte bestillingsukene, hvis ikke så finner vi riktig dato.

Select Case Bestillingsuker
    Case "Partallsuker"
        'Sjekker om datoen er oddetallsuke, hvis så legger vi på sju dager for å få partall.
        If (ISOWeekNum(lev_dato) Mod 2 = 0) = False Then
            lev_dato = lev_dato + 7
        End If
    Case "Oddetallsuker"
        'Sjekker om datoen er partallsuke, hvis så legger vi på sju dager for å få oddetall.
        If (ISOWeekNum(lev_dato) Mod 2 = 0) = True Then
            lev_dato = lev_dato + 7
        End If
    Case "Annet (spesifiser)"
        For i = 0 To 52
            If ISOWeekNum(lev_dato) + i > 52 Then
                f = i - 52
            Else
                f = i
            End If
            If InStr(Bestillingsuker_annet, Format(ISOWeekNum(lev_dato) + f, "00")) <> 0 Then
                Exit For
            End If
        Next i
        lev_dato = lev_dato + (7 * i)
    Case Else
        'Hvis ikke kunden har spesielle forhold rundt bestilling, så går vi ut i fra at dagens dato er et fint utgangspunkt for neste bestillinsdag
        ' så vi gjør ingenting
        
End Select


finne_bestillingsdag = lev_dato


End Function

Public Sub test_2()

'MsgBox (finne_bestillingsdag("3008627", DateValue(Now)))
MsgBox (finne_bestillingsdag("3048400", DateValue("17.01.2013")))

End Sub

