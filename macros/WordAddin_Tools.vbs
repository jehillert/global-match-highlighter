Attribute VB_Name = "WordAddin_Tools"
'ERASE ALL HIGHLIGHTS
Sub EraseMarkup()
'Original code - Word Macro Backup files: 'remove_colors.vb'
   SetSearchSettings ("EraseMarkup")
   Application.ScreenUpdating = False
   'Options.DefaultHighlightColorIndex = color
   Options.DefaultHighlightColorIndex = wdNoHighlight
   If Selection.Type = wdSelectionIP Then
      EraseMarkupWholeDocument
   Else
      EraseMarkupSelectedText
   End If
   Application.ScreenUpdating = True
   RestoreSettings
   End Sub
Sub EraseMarkupWholeDocument()
   With ActiveDocument.Range
      'REMOVE FONT
      .Font.ColorIndex = wdAuto
      
      'REMOVE HIGHLIGHTING
      Options.DefaultHighlightColorIndex = wdNoHighlight
      .HighlightColorIndex = wdNoHighlight
         
      'REMOVE BACKGROUND SHADING, CHAR
      .Font.Shading.Texture = wdTextureNone
      .Font.Shading.ForegroundPatternColor = wdColorAutomatic
      .Font.Shading.BackgroundPatternColor = wdColorAutomatic

      'REMOVE BACKGROUND SHADING, PARA
      .ParagraphFormat.Shading.Texture = wdTextureNone
      .ParagraphFormat.Shading.ForegroundPatternColor = wdColorAutomatic
      .ParagraphFormat.Shading.BackgroundPatternColor = wdColorAutomatic
         
      If Selection.Type = wdSelectionColumn Or Selection.Type = wdSelectionRow Then
         .Cells.Shading.Texture = wdTextureNone
         .Cells.Shading.ForegroundPatternColor = wdColorAutomatic
         .Cells.Shading.BackgroundPatternColor = wdColorAutomatic
      End If
   End With
   End Sub
Sub EraseMarkupSelectedText()
   UnHighlight
   unshade
   End Sub
'TOGGLE HIGHLIGHTED TEXT
Sub ShowHideHighlighting()
   ActiveWindow.View.ShowHighlight = Not ActiveWindow.View.ShowHighlight
   End Sub



