Attribute VB_Name = "Module1"


Sub Protect()

    shClientIntakeRecord.Protect DrawingObjects:=True, _
    Contents:=True, Scenarios:=True
    
End Sub

Sub Unprotect()

    shClientIntakeRecord.Unprotect
    
End Sub

Sub btnMakeEdits_Click()

shClientIntakeRecord.Unprotect

End Sub

