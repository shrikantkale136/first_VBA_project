Attribute VB_Name = "common"
Public isValidWorkSheet As Boolean
Public Const outPutFileName = "E:\VBA_Project\Output\My_VBA.pptx"
Public PowerPointApp As PowerPoint.Application

Public Function validateSheet()
Dim ws As Worksheet
    Dim counter As Integer, portFolioCount As Integer, benchmarkCount As Integer, flg As Boolean
    Dim port1, port2, bench1, bench2 As String
 
    counter = 0
    portFolioCount = 0
    benchmarkCount = 0
    'Get Sheet count in Workbook
    sheetsCount = ActiveWorkbook.Sheets.Count
    Debug.Print ("Sheet Count : " & ActiveWorkbook.Sheets.Count)
    If sheetsCount > 1 And sheetsCount <= 3 Then
        For Each ws In ActiveWorkbook.Worksheets
            If LCase(ws.Name) Like "*port*" Then
                portFolioCount = portFolioCount + 1
            ElseIf LCase(ws.Name) Like "*bench*" Then
                benchmarkCount = benchmarkCount + 1
            End If
        Next ws
        combinator = portFolioCount & "V" & benchmarkCount
        If portFolioCount <> 0 And benchmarkCount <> 0 Then
            Debug.Print ("Our Combination is : " & combinator)
            Debug.Print ("portFolioCount : " & portFolioCount)
            Debug.Print ("benchmarkCount : " & benchmarkCount)
            isValidWorkSheet = True
            MsgBox "Got Combination " & combinator & " !", vbInformation, "Success : Portfolio Scenarios"
        Else
            isValidWorkSheet = False
            MsgBox "Should have alteast one Portfolio And one Benchmark !", vbCritical, "Error : Portfolio Scenarios"
        End If
    Else
        isValidWorkSheet = False
        MsgBox "It should have atleast 2 or 3 sheets OR " & vbNewLine & "It should not have more than 3 or less than 2 sheets", vbCritical, "Error : Portfolio Scenarios"
    End If
End Function

Public Function GeneratePPT()
    Dim rng As Range
    
    Dim myPresentation As Object
    Dim ws As Worksheet
    'Create an Instance of PowerPoint
    On Error Resume Next
    
    'Is PowerPoint already opened?
    Set PowerPointApp = GetObject(class:="PowerPoint.Application")
    
    'Clear the error between errors
    Err.Clear
    
    'If PowerPoint is not already open then open PowerPoint
    If PowerPointApp Is Nothing Then Set PowerPointApp = CreateObject(class:="PowerPoint.Application")
    'Handle if the PowerPoint Application is not found
    If Err.Number = 429 Then
        MsgBox "PowerPoint could not be found, aborting."
    End If
    
    On Error GoTo 0
    'Optimize Code
    Application.ScreenUpdating = True
    
    'Create a New Presentation
    Set myPresentation = PowerPointApp.Presentations.Add

    'Debug.Print (ThisWorkbook.Worksheets)
    
    'Make PowerPoint Visible and Active
    'PowerPointApp.Visible = True
    'PowerPointApp.Activate
    
    'Clear The Clipboard
    Application.CutCopyMode = False
    'Add 1 slide to existing PPT
    'PowerPointApp.ActivePresentation.Slides.Add Index:=PowerPointApp.ActivePresentation.Slides.Count + 1, Layout:=ppLayoutCustom
    Call savePPT
    'Close running PPT
    'PowerPointApp.Windows(1).Close
    'Call openPPT(outPutFileName)
    
End Function

Public Function openPPT(path)
    Set openPPT = PowerPointApp.Presentations.Open(path)
    Debug.Print "Presentation open !"
End Function

Public Function savePPT()
    'Save File as PPT
    PowerPointApp.ActivePresentation.SaveAs Filename:=outPutFileName
    Debug.Print "Presentation Saved !"
End Function



