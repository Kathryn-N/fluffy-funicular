Attribute VB_Name = "modClientIntakeForm"
Option Explicit

' Declarations
Dim strClientName As String
Dim strContactName As String
Dim strPhone As String
Dim strEmail As String
Dim strReferral As String
Dim dateDate As Date
Dim strServices As String
Dim strSummary As String
Dim strQuestions As String


' Shows Client Intake Form


Sub frmClientIntakeForm_Initialize()

    frmClientIntakeForm.Controls("cboServices").List = Array("Accounting Services", _
    "Compliance Evaluation", "Custom Technology Solutions", "Employee Handbook Creation", _
    "Labor Standards Compliance", "Management Consulting", "Payroll Services", _
    "Process Evaluation and Documentation", "Project Management")
    
    frmClientIntakeForm.Show
    
End Sub


' Stores data in input fields as module-level variables and prints them to shClientIntakeRecord
Sub RecordData()

    With frmClientIntakeForm
        strClientName = .txtClientName.Value
        strContactName = .txtContactName.Value
        strPhone = .txtPhone.Value
        strEmail = .txtEmail.Value
        strReferral = .txtReferral.Value
        dateDate = .txtDate.Value
        strServices = .cboServices.Value
        strSummary = .txtSummary.Value
        strQuestions = .txtQuestions.Value
    End With
    
    With frmClientIntakeForm.txtDate.Value
        If Not .IsDate Then
        MsgBox "The date you entered is not valid." & vbNewLine & "Please try again,"
            
    End If
        
    End With
    
' Copies data to Sheet2
shClientIntakeRecord.Activate
ActiveSheet.Unprotect

Cells(5, 1).Activate ' activates first cell in data sheet header

    If ActiveCell.Offset(1, 0) = "" Then ' checks for next empty row
        ActiveCell.Offset(1, 0).Activate
    Else
        ActiveCell.End(xlDown).Offset(1, 0).Activate
    End If

ActiveCell.Value = strClientName    ' records form data into first empty row
ActiveCell.Offset(0, 1).Value = strContactName
ActiveCell.Offset(0, 2).Value = strPhone
ActiveCell.Offset(0, 3).Value = strEmail
ActiveCell.Offset(0, 4).Value = strReferral
ActiveCell.Offset(0, 5).Value = dateDate
ActiveCell.Offset(0, 6).Value = strServices
ActiveCell.Offset(0, 7).Value = strSummary
ActiveCell.Offset(0, 8).Value = strQuestions

' Protects sheet (without password)

shClientIntakeRecord.Protect DrawingObjects:=True, _
    Contents:=True, Scenarios:=True

' Clears variables

strClientName = ""
strContactName = ""
strPhone = ""
strEmail = ""
strReferral = ""
dateDate = ""
strServices = ""
strSummary = ""
strQuestions = ""
   
' Clears form

With frmClientIntakeForm
    .txtClientName.Value = ""
    .txtContactName.Value = ""
    .txtPhone.Value = ""
    .txtEmail.Value = ""
    .txtReferral.Value = ""
    .txtDate.Value = ""
    .txtQuestions = ""
End With
    
    
End Sub
