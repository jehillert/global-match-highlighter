Attribute VB_Name = "WordAddin_Shade"
'TO DO - figure out how to handle a region selected by holding down Alt and draging the mouse (wdSelectionBlock).  This could be useful (or not).  Maybe text pattern applications...
'**********************************************************************************************
' SHADING VARS
'**********************************************************************************************
'Dim obj_undo As UndoRecord
Private FormattingWasTracked As Boolean
Private LastBM As String
Private LastShadingOld
Private LastShadingNew

'**********************************************************************************************
' SHADE FUNCTIONS
'**********************************************************************************************
Sub Shade(TargetShade)
   On Error GoTo ErrorHandler
   Set obj_undo = Application.UndoRecord
   obj_undo.StartCustomRecord ("Shade()")
   'Application.ScreenUpdating = False
   If (ActiveDocument.TrackFormatting = True) Then
      FormattingWasTracked = True
      ActiveDocument.TrackFormatting = False
   End If
   With Selection
   If .Information(wdWithInTable) <> True Then
      'UNSHADE PARAGRAPH WHEN UNSHADED TEXT SELECTED
      If .Type <> wdSelectionIP Then
         'MsgBox "hi0"
         If Not FullParagraphsSelected And _
            .PARAGRAPHS.Count < 2 And _
            .Font.Shading.BackgroundPatternColorIndex = 0 And _
            .ParagraphFormat.Shading.BackgroundPatternColorIndex <> 0 And _
             TargetShade = wdColorAutomatic Then
               .ParagraphFormat.Shading.BackgroundPatternColor = wdColorAutomatic
               .ParagraphFormat.Shading.BackgroundPatternColorIndex = wdNoHighlight
         'TEXT SELECTED, BUT NOT PARA
         ElseIf Not FullParagraphsSelected And .PARAGRAPHS.Count < 2 Then
            With .Font.Shading
               If .BackgroundPatternColor = TargetShade Then
                  TargetShade = wdColorAutomatic
                  .BackgroundPatternColorIndex = wdNoHighlight
               End If
               .BackgroundPatternColor = TargetShade
            End With
         'ONE OR MORE FULL PARAGRAPHS SELECTED
         ElseIf FullParagraphsSelected = True Or .PARAGRAPHS.Count >= 2 Then
            Call ShadeParagraphs(TargetShade)
         End If
      ElseIf .Type = wdSelectionIP Then
         'NO SELECTION AND IP LOCATED IN SHADED NON-PARAGRAPH SELECTION
         'MsgBox "hi0"
         If .Font.Shading.BackgroundPatternColorIndex <> 0 Then
            ReshadeContiguouslyShadedText (TargetShade)
         'NO SELECTION AND IP LOCATED IN UNSHADED REGION OF PARAGRAPH
         Else
         'MsgBox "hi3"
            Call ShadeParagraphs(TargetShade)
         End If
      End If
   'TABLES
   ElseIf .Information(wdWithInTable) = True Then
      If .Type = wdSelectionColumn Or .Type = wdSelectionRow Then
         With .Cells.Shading
         If .BackgroundPatternColor = TargetShade Then
            TargetShade = wdColorAutomatic
            .BackgroundPatternColorIndex = wdNoHighlight
         End If
         .BackgroundPatternColor = TargetShade
         End With
      Else
         .Font.Shading.BackgroundPatternColor = TargetShade
         With .Font.Shading
         If .BackgroundPatternColor = TargetShade Then
            TargetShade = wdColorAutomatic
            .BackgroundPatternColorIndex = wdNoHighlight
         End If
         .BackgroundPatternColor = TargetShade
         End With
      End If
   Else
      MsgBox "Error in execution of sub shade(TargetShade). Invalid Selection.Type."
   End If
   End With
ErrorHandler:
   If Err <> 0 Then MsgBox "Error executing Shade() subroutine."
   If FormattingWasTracked Then ActiveDocument.TrackFormatting = True
   Application.ScreenUpdating = True
   obj_undo.EndCustomRecord
   End Sub
Sub MoveIPAfterHighlight(TargetShade, LastShadingOld)
   If Selection.Type <> wdSelectionIP Then Exit Sub
   Dim sRng, oRng As Range
   Set oRng = Selection.PARAGRAPHS(1).Range
   Set sRng = Selection.Range
   'LastShadingOld, LastShadingNew
   'EXIT IF JUST TRIED TO UNSHADE SOMETHING THAT WAS ALREADY UNSHADED
   'ALSO EXIT IF WILL BE TRYING TO UNSHADE A PARAGRAPH THAT'S ALREADY UNSHADED
   If TargetShade = wdColorAutomatic Then
      If (LastShadingOld = wdColorAutomatic) Or _
         Selection.PARAGRAPHS(1).Next.Range.HighlightColorIndex = wdColorAutomatic Then
            Exit Sub
      End If
   End If
   If Selection.Type = wdSelectionIP Then
      sRng.MoveEnd Unit:=wdParagraph, Count:=1
      If sRng.Characters.Count = 1 Or sRng.Characters.Count = oRng.Characters.Count Then
            MoveToEndOfNextParagraph
      Else
         Selection.collapse Direction:=wdCollapseStart
      End If
   End If
   End Sub
Sub ShadeParagraphs(TargetShade)
   LastShadingOld = Selection.ParagraphFormat.Shading.BackgroundPatternColor
   'EXIT IF TRYING TO UNSHADE A PARAGRAPH THAT'S ALREADY UNSHADED
   If TargetShade = wdColorAutomatic And Selection.ParagraphFormat.Shading.BackgroundPatternColor = wdColorAutomatic Then
      Exit Sub
   End If
   'IF TARGET AND DESTINATION SHADES ARE THE SAME, THEN CLEAR SHADING INSTEAD
   If Selection.ParagraphFormat.Shading.BackgroundPatternColor = TargetShade Then
      If Selection.Type = wdSelectionIP Then
         'IF IP IN A SINGLE COLOR MULTI-PARAGRAPH BLOCK, JUST CLEAR THE CURRENT PARAGRAPH
         Selection.PARAGRAPHS(1).Shading.BackgroundPatternColor = wdColorAutomatic
      Else
         'OTHERWISE CLEAR ALL PARAGRAPHS OF MULTI-PARAGRAPH SELECTION
         Selection.ParagraphFormat.Shading.BackgroundPatternColor = wdColorAutomatic
      End If
   'IF NOT THE SAME, THEN SHADE ACCORDING TO "TargetShade"
   ElseIf Selection.Type = wdSelectionIP Then
      'IF IP AN UNSHADED OR DIFFERENT-COLORED PARAGRAPH, RESHADE TO 'TargetShade'
      Selection.PARAGRAPHS(1).Shading.BackgroundPatternColor = TargetShade
   'OTHERWISE RESHADE ALL PARAGRAPHS OF MULTI-PARAGRAPH SELECTION TO 'TargetShade'
   Else
      Selection.ParagraphFormat.Shading.BackgroundPatternColor = TargetShade
   End If
   Call MoveIPAfterHighlight(TargetShade, LastShadingOld)
   End Sub
Sub ShadeSelection(TargetShade)
   Selection.Shading.BackgroundPatternColor = TargetShade
   End Sub
Sub ScrollToUpperLeftCorner()
   Application.ScreenUpdating = False
   ActiveWindow.ScrollIntoView Selection.Range, True
   ActiveDocument.ActiveWindow.SmallScroll down:=14
   Application.ScreenUpdating = True
   End Sub
'**********************************************************************************************
'* AUXILIARY FUNCTIONS
'**********************************************************************************************
Sub SetShadingSettings(SubName As String)
   With Selection.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Font.Shading.BackgroundPatternColor = wdColorAutomatic
      .Font.Shading.BackgroundPatternColorIndex = wdNoHighlight
      .Text = ""
      .Replacement.Text = ""
      .Forward = True
      .Wrap = wdFindContinue
      .FORMAT = True
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
   End With
   End Sub
Sub RestoreShadingSettings()
   With Selection.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Forward = True
      .FORMAT = False
   End With
   End Sub
Sub unshade()
   Shade (wdColorAutomatic)
   End Sub
'**********************************************************************************************
'* BOOKMARK FUNCTIONS
'**********************************************************************************************
Function DeleteLastBM()
   ActiveDocument.Bookmarks(LastBM).Delete
   End Function
Function GoToLastBM()
   Selection.GoTo what:=wdGoToBookmark, Name:=LastBM
   End Function
Function NewBM() As String
   NewBM = "BM" & CStr(ActiveDocument.Bookmarks.Count + 1)
   ActiveDocument.Bookmarks.Add Name:=NewBM
   LastBM = NewBM
   End Function
'   Call SetShadingSettings("unshade")
'   If Selection.Information(wdWithInTable) = True Then
'      If Selection.Type = wdSelectionColumn Or Selection.Type = wdSelectionRow Then
'         Call ShadeCells(wdColorAutomatic)
'      ElseIf Selection.Type = wdSelectionIP Then
'         Call ShadeCells(wdColorAutomatic)
'         Call ShadeText(wdColorAutomatic)
'      Else
'         Call ShadeText(wdColorAutomatic)
'      End If
'   ElseIf Selection.Words.Count = ActiveDocument.Words.Count Then
'      Call ShadeText(wdColorAutomatic)
'      Call ShadeParagraphs(wdColorAutomatic)
'      If Selection.Type = wdSelectionColumn Or Selection.Type = wdSelectionRow Then
'         ShadeCells (wdColorAutomatic)
'      End If
'   ElseIf Selection.Type = wdSelectionIP Or FullParagraphsSelected Then
'      'if the cursor is at a spot that has only paragraph shading, then remove the paragraph shading
'      Call ShadeParagraphs(wdColorAutomatic)
'      If MoveAfterUnshading = 1 Then NextPara
'      'THIS NEXT PART STILL NEEDS TO BE ADDED
'      '   but if the cursor is inside shaded text that doesn't start and end at a paragraph,
'      '   then removed the shading just from the text
'   Else
'      Call ShadeText(wdColorAutomatic)
'   End If
'   RestoreShadingSettings
'   End Sub
'Function ReshadeText(Optional TargetShade = wdColorAutomatic)
'   Call SetShadingSettings("ReshadeText", True)
'   With Selection
'      If .Font.SHADING.BackgroundPatternColorIndex = 0 Then Exit Function
'      .Find.Font.SHADING.BackgroundPatternColor = .Font.SHADING.BackgroundPatternColor
'      .Find.Execute
'      .collapse Direction:=wdCollapseEnd
'      .Find.Forward = False
'      .Find.Execute
'      .Font.SHADING.BackgroundPatternColor = TargetShade
'   End With
'   Call RestoreShadingSettings(True)
'   End Function
