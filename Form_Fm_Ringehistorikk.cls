VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Fm_Ringehistorikk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public Sub StartForm(k_nr As String, dato As String)

Dim filter_streng As String
Dim dato2 As Date
dato2 = CDate(dato)

filter_streng = "[Kundenummer] = '" & k_nr & "' AND [Ringedag] = '" & dato & "'"



Me.Filter = filter_streng
Me.FilterOn = True

Me.Visible = True

If IsNull(Leveringsdag.Value) Then
   Leveringsdag.Value = finne_leveringsdag(k_nr, dato2)
End If
End Sub
Private Sub Command8_Click()
On Error GoTo Err_Command8_Click


    DoCmd.Close
    Form_Kunderegister_alle_datasheet.Refresh

Exit_Command8_Click:
    Exit Sub

Err_Command8_Click:
    MsgBox Err.Description
    Resume Exit_Command8_Click
    
End Sub

Private Sub Form_Close()
'Form_Kunderegister_alle_datasheet.Refresh
End Sub
