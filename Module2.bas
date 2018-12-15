Attribute VB_Name = "Module2"
Option Compare Database

Public Sub lage_calls(k_nr As String, nydato As String)

Dim rsData As DAO.Recordset
Dim intSokestreng As String

intSokestreng = "Select * From Ringehistorikk Where Kundenummer = '" & k_nr & "' And Ringedag = '" & nydato & "';"

Set rsData = CurrentDb.OpenRecordset(intSokestreng, dbOpenDynaset, dbSeeChanges)

If rsData.EOF = True Then
    With rsData
        .AddNew
        !Kundenummer = k_nr
        !Ringedag = nydato
        !Resultat = "Ikke ringt enda"
        .Update
    End With
End If

rsData.Close

End Sub


Public Sub sjekke_calls(sokestreng As String, dato As String)

Dim stNysokestreng As String
Dim rsData As DAO.Recordset
Dim rsRingeliste As DAO.Recordset
Dim aapnet_ringehistorikk As Boolean

stNysokestreng = "Select * From Kunderegister Where " & sokestreng
aapnet_ringehistorikk = False

Set rsData = CurrentDb.OpenRecordset(stNysokestreng)

Do While rsData.EOF = False
    aapnet_ringehistorikk = True
    Set rsRingeliste = CurrentDb.OpenRecordset("SELECT * FROM Ringehistorikk Where Kundenummer = '" & rsData!Kundenummer & "' AND Ringedag = '" & dato & "';", dbOpenDynaset, dbSeeChanges)
    If rsRingeliste.EOF = True Then
        Call lage_calls(rsData!Kundenummer, dato)
    End If
    rsData.MoveNext
Loop

'Call lage_calls("310", "30.11.2010")

rsData.Close

If aapnet_ringehistorikk = True Then
    rsRingeliste.Close
End If

End Sub

Public Sub sjekke_avvikende_uker()

Dim rsData As DAO.Recordset
Dim rsKundedata As DAO.Recordset
Dim ukestreng As String

Set rsKundedata = CurrentDb.OpenRecordset("Select * FROM kunderegister")
    With rsKundedata
    Do While .EOF = False
        .Edit
        ![Bestillingsuker (annet)] = "N/A"
        .Update
        .MoveNext
    Loop
    End With
rsKundedata.Close

Set rsData = CurrentDb.OpenRecordset("Select * FROM avvikende_uker")
With rsData
    Do While rsData.EOF = False
        ukestreng = ""
        If ![Uke 1] = True Then
            ukestreng = ukestreng & "01, "
        End If
        If ![Uke 2] = True Then
            ukestreng = ukestreng & "02, "
        End If
        If ![Uke 3] = True Then
            ukestreng = ukestreng & "03, "
        End If
        If ![Uke 4] = True Then
            ukestreng = ukestreng & "04, "
        End If
        If ![Uke 5] = True Then
            ukestreng = ukestreng & "05, "
        End If
        If ![Uke 6] = True Then
            ukestreng = ukestreng & "06, "
        End If
        If ![Uke 7] = True Then
            ukestreng = ukestreng & "07, "
        End If
        If ![Uke 8] = True Then
            ukestreng = ukestreng & "08, "
        End If
        If ![Uke 9] = True Then
            ukestreng = ukestreng & "09, "
        End If
        If ![Uke 10] = True Then
            ukestreng = ukestreng & "10, "
        End If
        If ![Uke 11] = True Then
            ukestreng = ukestreng & "11, "
        End If
        If ![Uke 12] = True Then
            ukestreng = ukestreng & "12, "
        End If
        If ![Uke 13] = True Then
            ukestreng = ukestreng & "13, "
        End If
        If ![Uke 14] = True Then
            ukestreng = ukestreng & "14, "
        End If
        If ![Uke 15] = True Then
            ukestreng = ukestreng & "15, "
        End If
        If ![Uke 16] = True Then
            ukestreng = ukestreng & "16, "
        End If
        If ![Uke 17] = True Then
            ukestreng = ukestreng & "17, "
        End If
        If ![Uke 18] = True Then
            ukestreng = ukestreng & "18, "
        End If
        If ![Uke 19] = True Then
            ukestreng = ukestreng & "19, "
        End If
        If ![Uke 20] = True Then
            ukestreng = ukestreng & "20, "
        End If
        If ![Uke 21] = True Then
            ukestreng = ukestreng & "21, "
        End If
        If ![Uke 22] = True Then
            ukestreng = ukestreng & "22, "
        End If
        If ![Uke 23] = True Then
            ukestreng = ukestreng & "23, "
        End If
        If ![Uke 24] = True Then
            ukestreng = ukestreng & "24, "
        End If
        If ![Uke 25] = True Then
            ukestreng = ukestreng & "25, "
        End If
        If ![Uke 26] = True Then
            ukestreng = ukestreng & "26, "
        End If
        If ![Uke 27] = True Then
            ukestreng = ukestreng & "27, "
        End If
        If ![Uke 28] = True Then
            ukestreng = ukestreng & "28, "
        End If
        If ![Uke 29] = True Then
            ukestreng = ukestreng & "29, "
        End If
        If ![Uke 30] = True Then
            ukestreng = ukestreng & "30, "
        End If
        If ![Uke 31] = True Then
            ukestreng = ukestreng & "31, "
        End If
        If ![Uke 32] = True Then
            ukestreng = ukestreng & "32, "
        End If
        If ![Uke 33] = True Then
            ukestreng = ukestreng & "33, "
        End If
        If ![Uke 34] = True Then
            ukestreng = ukestreng & "34, "
        End If
        If ![Uke 35] = True Then
            ukestreng = ukestreng & "35, "
        End If
        If ![Uke 36] = True Then
            ukestreng = ukestreng & "36, "
        End If
        If ![Uke 37] = True Then
            ukestreng = ukestreng & "37, "
        End If
        If ![Uke 38] = True Then
            ukestreng = ukestreng & "38, "
        End If
        If ![Uke 39] = True Then
            ukestreng = ukestreng & "39, "
        End If
        If ![Uke 40] = True Then
            ukestreng = ukestreng & "40, "
        End If
        If ![Uke 41] = True Then
            ukestreng = ukestreng & "41, "
        End If
        If ![Uke 42] = True Then
            ukestreng = ukestreng & "42, "
        End If
        If ![Uke 43] = True Then
            ukestreng = ukestreng & "43, "
        End If
        If ![Uke 44] = True Then
            ukestreng = ukestreng & "44, "
        End If
        If ![Uke 45] = True Then
            ukestreng = ukestreng & "45, "
        End If
        If ![Uke 46] = True Then
            ukestreng = ukestreng & "46, "
        End If
        If ![Uke 47] = True Then
            ukestreng = ukestreng & "47, "
        End If
        If ![Uke 48] = True Then
            ukestreng = ukestreng & "48, "
        End If
        If ![Uke 49] = True Then
            ukestreng = ukestreng & "49, "
        End If
        If ![Uke 50] = True Then
            ukestreng = ukestreng & "50, "
        End If
        If ![Uke 51] = True Then
            ukestreng = ukestreng & "51, "
        End If
        If ![Uke 52] = True Then
            ukestreng = ukestreng & "52, "
        End If
        If ![Uke 53] = True Then
            ukestreng = ukestreng & "53, "
        End If
        
        lngLen = Len(ukestreng) 'uten ,
        Debug.Print lngLen & "lnglen"
        If lngLen > 0 Then
            ukestreng = Left$(ukestreng, lngLen)
            Set rsKundedata = CurrentDb.OpenRecordset("Select * FROM kunderegister Where Kundenummer ='" & !Kundenummer & "'")
            If Not rsKundedata.EOF = True Then
                rsKundedata.Edit
                rsKundedata![Bestillingsuker (annet)] = ukestreng
                rsKundedata.Update
                Debug.Print !Kundenummer & "**"
            End If
            rsKundedata.Close
        Else
            Set rsKundedata = CurrentDb.OpenRecordset("Select * FROM kunderegister Where Kundenummer ='" & !Kundenummer & "'")
            If Not rsKundedata.EOF = True Then
                rsKundedata.Edit
                rsKundedata![Bestillingsuker (annet)] = "N/A"
                rsKundedata.Update
                Debug.Print !Kundenummer
            End If
            rsKundedata.Close
        End If
        rsData.MoveNext
    Loop
End With
rsData.Close

End Sub

Public Function finne_leveringsdag(k_nr As String, dato As Date) As Date

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
Dim frekvens As Integer

Set rsKundedata = CurrentDb.OpenRecordset("Select * FROM kunderegister Where Kundenummer ='" & k_nr & "'")

mandag = rsKundedata![Leveringsdag: Mandag]
tirsdag = rsKundedata![Leveringsdag: Tirsdag]
onsdag = rsKundedata![Leveringsdag: Onsdag]
torsdag = rsKundedata![Leveringsdag: Torsdag]
fredag = rsKundedata![Leveringsdag: Fredag]
frekvens = rsKundedata![Bestillingsfrekvens]

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

best_dag = WeekDay(dato, vbMonday)

If best_dag + 1 + frekvens <= lev_dag Then
    lev_dato = dato + (lev_dag - best_dag)
ElseIf lev_dag <= best_dag + 1 Then
    lev_dato = dato + ((lev_dag - best_dag) + 7)
End If

If funnet_leveringsdag > 1 Then
    If best_dag + 1 + frekvens <= lev_dag2 Then
        lev_dato2 = dato + (lev_dag2 - best_dag)
    ElseIf lev_dag2 <= best_dag + 1 + frekvens Then
        lev_dato2 = dato + ((lev_dag2 - best_dag) + 7)
    End If

    If lev_dato2 < lev_dato Then
        lev_dato = lev_dato2
    End If
End If

' lev_dato er kundens leveringsdato
' rutinen fungerer ikke hvis kunden har mer enn 2 leveringsdatoer

finne_leveringsdag = lev_dato

End Function
