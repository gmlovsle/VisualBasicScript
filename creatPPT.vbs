Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True
Set objPresentation = objPPT.Presentations.Add
objPresentation.ApplyTemplate("C:\Office2013\Templates\2052\ContemporaryPhotoAlbum.potx")
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("Select * From Win32_Process")

With Application.ActivePresentation.Slides(2)
With .Shapes.AddTextbox(msoTextOrientationHorizontal, 300, 300, 500, 400)
.TextFrame.TextRange.Text = "µ½´ËÒ»ÓÎ"
End With
End With

objPresentation.SaveAs("C:\Users\Administrator\Desktop\Process.ppt")
objPresentation.Close
objPPT.Quit
