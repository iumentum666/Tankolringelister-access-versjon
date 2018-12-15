VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Registrere_kunde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Bestillingsuker__annet__Click()
Kundenummer.Locked = True
Form_avvikende_uker_form.StartForm Kundenummer.Value
End Sub

Private Sub Command42_Click()
On Error GoTo Err_Command42_Click
Kundenummer.Locked = True
Kundenummer.SetFocus

    Screen.PreviousControl.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70

Exit_Command42_Click:
    Exit Sub

Err_Command42_Click:
    MsgBox Err.Description
    Resume Exit_Command42_Click
    
End Sub
Private Sub Command43_Click()
On Error GoTo Err_Command43_Click
Kundenummer.Locked = True

    DoCmd.Close

Exit_Command43_Click:
    Exit Sub

Err_Command43_Click:
    MsgBox Err.Description
    Resume Exit_Command43_Click
    
End Sub
Private Sub Command45_Click()
On Error GoTo Err_Command45_Click
Kundenummer.Locked = True

    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70

Exit_Command45_Click:
    Exit Sub

Err_Command45_Click:
    MsgBox Err.Description
    Resume Exit_Command45_Click
    
End Sub
Private Sub Command46_Click()
On Error GoTo Err_Command46_Click
Kundenummer.Locked = False

    DoCmd.GoToRecord , , acNewRec

Exit_Command46_Click:
    Exit Sub

Err_Command46_Click:
    MsgBox Err.Description
    Resume Exit_Command46_Click
    
End Sub

