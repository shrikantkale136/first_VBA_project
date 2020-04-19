Attribute VB_Name = "mCoverPage"
Sub CoverPage()
    Dim value1 As String
    Dim value2 As String
    value1 = ThisWorkbook.Sheets(1).Range("A1").Value 'value from sheet1
    value2 = ThisWorkbook.Sheets(1).Range("A2").Value 'value from sheet1
    value3 = Date
    PowerPointApp.ActivePresentation.Slides.InsertFromFile Filename:="E:\VBA_Project\Template\template1.pptx", Index:=0, SlideStart:=1, SlideEnd:=1
    Debug.Print "CoverPage Slide copied from source !"
    Set FirstSlide = PowerPointApp.ActivePresentation.Slides(1)
    'FirstSlide.Shapes.AddShape msoShapeRectangle, 5, 25, 100, 50
    FirstSlide.Shapes(1).TextFrame.TextRange.Text = value1
    FirstSlide.Shapes(2).TextFrame.TextRange.Text = value2
    FirstSlide.Shapes(3).TextFrame.TextRange.Text = value3
    PowerPointApp.ActivePresentation.Slides.InsertFromFile Filename:="E:\VBA_Project\Template\template1.pptx", Index:=1, SlideStart:=3, SlideEnd:=3
    Debug.Print "DisclosurePage slide copied from source !"
    'Set DisclosureSlide = PowerPointApp.ActivePresentation.Slides(2)
    Call savePPT
End Sub

