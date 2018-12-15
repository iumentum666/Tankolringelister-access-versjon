VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_avvikende_uker_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Public Sub StartForm(k_nr As String)

Dim filter_streng As String
Dim rsAvvikende As DAO.Recordset

Set rsAvvikende = CurrentDb.OpenRecordset("SELECT * FROM avvikende_uker where kundenummer = '" & k_nr & "'")

If rsAvvikende.EOF = True Then
    rsAvvikende.AddNew
    rsAvvikende!Kundenummer = k_nr
    rsAvvikende.Update
End If

rsAvvikende.Close

filter_streng = "[Kundenummer] = '" & k_nr & "'"

Me.Filter = filter_streng
Me.FilterOn = True

Me.Visible = True

End Sub
Private Sub Command109_Click()
On Error GoTo Err_Command109_Click


    DoCmd.Close

Exit_Command109_Click:
    Exit Sub

Err_Command109_Click:
    MsgBox Err.Description
    Resume Exit_Command109_Click
    
End Sub

Private Sub Form_Close()
sjekke_avvikende_uker
End Sub

