VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Kunderegister_alle_datasheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Bestillingsuker__annet__Click()

Form_avvikende_uker_form.StartForm Kundenummer.Value


End Sub


Private Sub Bestillingsuker_Click()
Command207_Click
End Sub

Private Sub Kundenavn_Click()
Command207_Click
End Sub

Private Sub Kundenummer_Click()
Command207_Click
End Sub


Private Sub Kundenummer_DblClick(Cancel As Integer)
Form_Fm_Ringehistorikk.StartForm Kundenummer.Value, Ringedag.Value

End Sub

Private Sub Leveringsdag_DblClick(Cancel As Integer)
Form_Fm_Ringehistorikk.StartForm Kundenummer.Value, Ringedag.Value
End Sub

Private Sub Resultat_Click()

'Form_Fm_Ringehistorikk.StartForm Kundenummer.Value, Ringedag.Value

End Sub
Private Sub Command207_Click()
On Error GoTo Err_Command207_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Registrere_kunde"
    
    stLinkCriteria = "[Kundenummer]=" & "'" & Me![Kundenummer] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command207_Click:
    Exit Sub

Err_Command207_Click:
    MsgBox Err.Description
    Resume Exit_Command207_Click
    
End Sub

Private Sub Resultat_DblClick(Cancel As Integer)
Form_Fm_Ringehistorikk.StartForm Kundenummer.Value, Ringedag.Value
End Sub

Private Sub Ringedag_DblClick(Cancel As Integer)
Form_Fm_Ringehistorikk.StartForm Kundenummer.Value, Ringedag.Value
End Sub

Private Sub Solgt_volum_DblClick(Cancel As Integer)
Form_Fm_Ringehistorikk.StartForm Kundenummer.Value, Ringedag.Value
End Sub

Private Sub Solgt_volum_Enter()
'Form_Fm_Ringehistorikk.StartForm Kundenummer.Value, Ringedag.Value
End Sub
