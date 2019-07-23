Attribute VB_Name = "WordAddin_PAGE"
Sub BgColor(r, g, B)
   Set objUndo = Application.UndoRecord
   objUndo.StartCustomRecord ("BgColor()")
   ActiveDocument.Background.Fill.Visible = msoTrue
   If ActiveDocument.Background.Fill.ForeColor.RGB <> RGB(r, g, B) Then
      ActiveDocument.Background.Fill.ForeColor.RGB = RGB(r, g, B)
   End If
   ActiveDocument.Background.Fill.Solid
   objUndo.EndCustomRecord
End Sub
Sub BgInvert()
   Set objUndo = Application.UndoRecord
   objUndo.StartCustomRecord ("BgInvert()")
   ActiveDocument.Background.Fill.ForeColor.ObjectThemeColor = _
    wdThemeColorText1
   ActiveDocument.Background.Fill.ForeColor.TintAndShade = 0#
   ActiveDocument.Background.Fill.Visible = msoTrue
   ActiveDocument.Background.Fill.Solid
   objUndo.EndCustomRecord
End Sub

