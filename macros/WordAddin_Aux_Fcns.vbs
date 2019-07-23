Attribute VB_Name = "WordAddin_Aux_Fcns"
Dim obj_undo As UndoRecord
Private initial_screen_updating_state, initial_track_changes_state As Boolean
'ANNOUNCE
Function Announce(msg As String, Optional ttl As String = "Attention!", Optional duration As Integer = 1)
   Set InfoBox = CreateObject("WScript.Shell")
   Select Case InfoBox.Popup(msg, duration, ttl)
   End Select
   End Function
'EXCEEDS FIND & REPLACE CHARACTER LIMIT
Function exceedsFindReplaceCharLimit() As Boolean
   If Selection.Characters.Count > 254 Then
      exceedsFindReplaceCharLimit = True
   End If
   End Function
'MARK SPOT
Sub MarkSpot(mySpot As String)
   If ActiveDocument.Bookmarks.Exists(mySpot) = True Then ActiveDocument.Bookmarks(mySpot).Delete
   ActiveDocument.Bookmarks.Add Name:=mySpot
   End Sub
Function MoveToEndOfNextParagraph()
   Selection.collapse Direction:=wdCollapseStart
   Selection.MoveDown Unit:=wdParagraph, Count:=1
   'Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
   End Function
'WHOLE STORY SELECTED
Function whole_story_selected() As Boolean
   If Selection.Characters.Count = Selection.StoryLength Then
      whole_story_selected = True
   End If
   End Function
'FULL PARAGRAPHS SELECTED
Function FullParagraphsSelected() As Boolean
   Dim selection1 As Range
   Dim selection2 As Range
   MarkSpot ("mySpotB4FullParagraphsSelectedWasRan")
   Set selection1 = Selection.Range
   Selection.Expand Unit:=wdParagraph
   Set selection2 = Selection.Range
   If selection1 = selection2 Then
    FullParagraphsSelected = True
   Else
    FullParagraphsSelected = False
   End If
   ReturnToSpot ("mySpotB4FullParagraphsSelectedWasRan")
   End Function
'MOVE INSERTION POINT TO TOP OF NEXT PARAGRAPH
Sub NextPara()
   Dim c As Boolean
   If c = Selection.MoveDown(wdParagraph, 1) Then
    Selection.HomeKey Unit:=wdStory, Extend:=wdMove
   End If
   End Sub
Function RemoveHighlightingFromParagraph()
   Selection.PARAGRAPHS(1).Range.HighlightColorIndex = wdNoHighlight
   Selection.PARAGRAPHS(1).Range.Font.ColorIndex = wdAuto
   End Function
Function ReshadeContiguouslyShadedText(TargetShade)
   Dim txtRng As Range
   Set txtRng = Selection.Range
   With txtRng
      If (Selection.Type <> wdSelectionIP Or _
         .Font.Shading.BackgroundPatternColorIndex = 0) Then
            Exit Function
      End If
      .Find.Font.Shading.BackgroundPatternColor = .Font.Shading.BackgroundPatternColor
      .Find.Execute
      .collapse Direction:=wdCollapseEnd
      .Find.Forward = False
      .Find.Execute
      '.Find.Forward = False
      .Font.Shading.BackgroundPatternColor = TargetShade
   End With
End Function
'RETURN TO SPOT
Sub ReturnToSpot(mySpot As String)
   If ActiveDocument.Bookmarks.Exists(mySpot) = True Then
    Selection.GoTo what:=wdGoToBookmark, Name:=mySpot
   End If
   ActiveDocument.Bookmarks(mySpot).Delete
   End Sub
'NOT CURRENTLY USED
Function SelectContiguouslyColoredText()
   'Assumes that search-settings still as they whould be in parent calling subroutine
   With Selection
      'EXIT IF IP IN TEXT HAVING DEFAULT FONT COLOR
      If (.Type <> wdSelectionIP) Or (.Range.Font.ColorIndex = wdAuto) Then
         Exit Function
      Else
         .Find.Font.color = Selection.Font.color
         .Find.Execute
         .collapse Direction:=wdCollapseEnd
         .Find.Forward = False
         .Find.Execute
      End If
   End With
End Function
'NOT CURRENTLY USED
Function SelectContiguouslyHighlightedText()
   Call SetSearchSettings("SelectContiguouslyHighlightedText")
   With Selection
      If (.Type <> wdSelectionIP Or _
         .Range.HighlightColorIndex = wdNoHighlight) Then
            Exit Function
      Else
         .Find.Highlight = True
         .Find.Execute
         .collapse Direction:=wdCollapseEnd
         .Find.Forward = False
         .Find.Execute
         .Find.Forward = True
      End If
   End With
   End Function
'NOT CURRENTLY USED
Function SelectContiguouslyShadedText()
   With Selection
      If (.Type <> wdSelectionIP Or _
         .Font.Shading.BackgroundPatternColorIndex = 0) Then
            Exit Function
      End If
      .Find.Font.Shading.BackgroundPatternColor = Selection.Font.Shading.BackgroundPatternColor
      .Find.Execute
      .collapse Direction:=wdCollapseEnd
      .Find.Forward = False
      .Find.Execute
      '.Find.Forward = False
   End With
   End Function
'TRIM SELECTION - exclude trailing white space and non-alphanumeric characters from selection.
Sub TrimSelection()
   Selection.MoveEndWhile Cset:=(Chr$(10) & Chr$(32) & Chr$(13) & ".,?!"), Count:=-Selection.Characters.Count
   Selection.MoveStartWhile Cset:=(Chr$(10) & Chr$(32) & Chr$(13) & ".,?!"), Count:=Selection.Characters.Count
   End Sub

