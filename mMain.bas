Attribute VB_Name = "mMain"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'
Dim answer As Integer
answer = MsgBox("Do you want to Continue?", vbQuestion + vbYesNo, "Generate Portfolio Scenarios Presentation")
  If answer = vbYes Then
    Call validateSheet
    If isValidWorkSheet Then
        Call GeneratePPT
        Call CoverPage
    Else
        Debug.Print ("Validation FAILED !")
    End If
  Else
  End If
End Sub

