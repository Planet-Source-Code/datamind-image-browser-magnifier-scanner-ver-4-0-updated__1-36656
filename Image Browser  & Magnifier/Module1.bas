Attribute VB_Name = "Module1"
Sub DH()
ErrorNum = Err.Number
ErrorDesc = Err.Description
Beep
MsgBox "An Error occured while executing command. Error Number :" & ErrorNum & "Error Description :" & ErrorDesc, vbCritical, "Error"
End Sub
