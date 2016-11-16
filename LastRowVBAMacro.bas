Attribute VB_Name = "LastRowVBAMacro"
Sub LastRow()
Attribute LastRow.VB_ProcData.VB_Invoke_Func = " \n14"
' LastRow macro

' Declare variables.
' As rows numbers can sometimes be very large,
' declare nLastRow As Long, NOT As Integer...
Dim sColumnLetter As String
Dim nLastRow As Long

' In this example, let's pretend we are trying to find the last row in column B,
' but any column will do...
sColumnLetter = "B"

' Set nLastRow as the first row containing a value,
' starting from the bottom of the worksheet...
nLastRow = Range(sColumnLetter & ActiveSheet.Rows.Count).End(xlUp).Row

' In this example, let's print the last row number in cell A1 to test the macro.
' In practice, we may not bother printing the variable value.
Range("A1").Value2 = nLastRow

End Sub
