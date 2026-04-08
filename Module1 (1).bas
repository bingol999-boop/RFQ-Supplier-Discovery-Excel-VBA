Attribute VB_Name = "Module1"
Option Explicit

Public Sub RunRFQProcess()

    On Error GoTo ErrorHandler

    MsgBox "RFQ process started.", vbInformation

    ' Placeholder logic
    ' TODO: Add Outlook scanning and AI extraction logic

    MsgBox "RFQ process completed successfully.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical

End Sub
