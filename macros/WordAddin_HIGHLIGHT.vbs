Attribute VB_Name = "WordAddin_HIGHLIGHT"
Sub Macro1()
'   Dim rng As Range
'   Set rng = Selection.Range
'   MsgBox "rng.Font.Color --- " & rng.Font.color
'   MsgBox "rng.Characters(1).Font.Color --- " & rng.Characters(1).Font.color
'   MsgBox "rng.Characters.First.Previous.Font.color --- " & rng.Characters.First.Previous.Font.color
   MsgBox Selection.Words.Count
   Selection.Expand Unit:=wdWord
   MsgBox Selection.Words.Count
End Sub
Function HighlightAndChangeFont(HColor, Optional FColor = wdColorAutomatic)
   'NOTE - IN THE FOLLOWING THREE EXAMPLES   (1) GIVES HL_INDEX LEFT OF INSERTION POINT,
   '                                         (2) GIVES HL_INDEX RIGHT OF INSERTION POINT
   '                                         (3) GOES ONE CHARACTER TO THE LEFT TOO FAR
   ' (1) MsgBox "rng.HighlightColorIndex --- " & rng.HighlightColorIndex
   ' (2) MsgBox "rng.Characters(1).HighlightColorIndex --- " & rng.Characters(1).HighlightColorIndex
   ' (3) MsgBox "Selection.Characters.First.Previous.HighlightColorIndex --- " & Selection.Characters.First.Previous.HighlightColorIndex
   ' FONT BEHAVIOR IS DIFFERENT THAN HIGHLIGHTING
   Set obj_undo = Application.UndoRecord
   obj_undo.StartCustomRecord ("HighlightAndChangeFont()")
   On Error GoTo ErrorHandler
   Application.ScreenUpdating = False
   CursorHide
   Dim txtRng As Range
   Dim docRng As Range
   Dim MiddleOfWord As Boolean
   Dim OldColorIndex As String
   Set docRng = ActiveDocument.Range
  'DETERMINE TARGET TEXT
   Set txtRng = Selection.Range
   With txtRng
     'IF SELECTED TEXT ALREADY THAT COLOR, THEN ERASE IT INSTEAD
'      If (.HighlightColorIndex = HColor) Or (.Characters(1).HighlightColorIndex = HColor) Then
'         HColor = wdNoHighlight
'         FColor = wdColorAutomatic
'      End If
     'IF SELECTION, THEN ASSIGN TO RANGE, TRIM RANGE, COLLAPSE SELECTION
      If Selection.Type <> wdSelectionIP Then
         .MoveEndWhile Cset:=(Chr$(10) & Chr$(32) & Chr$(13) & ".,?!"), Count:=-.Characters.Count
         .MoveStartWhile Cset:=(Chr$(10) & Chr$(32) & Chr$(13) & ".,?!"), Count:=.Characters.Count
        'COLLAPSEEND IF IN MIDDLE OF WORD
         If .Words.Count = 1 And (.Characters.Count < .Words(1).Characters.Count) Then
            Selection.collapse Direction:=wdCollapseEnd
        'COLLAPSE START IF WHOLE WORD WAS SELECTED
         Else
            Selection.collapse Direction:=wdCollapseStart
         End If
      ElseIf Selection.Type = wdSelectionIP Then
        'BEHAVIOR IF NO SELECTION, BUT HIGHLIGHTED TEXT NEARBY
        .Find.Highlight = True
         If (.HighlightColorIndex <> wdNoHighlight) Or (.Characters(1).HighlightColorIndex <> wdNoHighlight) Then
           'IP CCT LEFT OF CCT OR BETWEEN DIFFERENT CCT1|CCT2, EXPAND RIGHT
            If (.Characters(1).HighlightColorIndex <> wdNoHighlight) And _
               (.HighlightColorIndex <> .Characters(1).HighlightColorIndex) Then
               .Find.Forward = True
               .Find.Execute
           'IP RIGHT OF CCT AND LEFT OF UNCOLORED TEXT, EXPAND LEFT
            ElseIf (.Characters(1).HighlightColorIndex = wdNoHighlight) And _
                   (.HighlightColorIndex <> wdNoHighlight) Then
               .Find.Forward = False
               .Find.Execute
               .Find.Forward = True
              'IP IN THE MIDDLE OF A WORD, BETWEEN DIFFERENT COLORED CHARACTERS, SELECT WHOLE WORD
               If (.Words.Count = 1) And (.Characters.Count < .Words(1).Characters.Count) Then
                  MiddleOfWord = True
                  .Expand Unit:=wdWord
                  .MoveEndWhile Chr(32), wdBackward
               End If
           'IP WITHIN CCT, EXPAND BOTH DIRECTIONS
            Else
               .Find.Forward = False
               .Find.Execute
               .collapse Direction:=wdCollapseStart
               .Find.Forward = True
               .Find.Execute
              'IP IN MIDDLE OF HIGHLIGHTED PART OF PARTLY WORDED TEXT
               If (.Words.Count = 1) And (.Characters.Count < .Words(1).Characters.Count) Then
                  .Expand Unit:=wdWord
                  .MoveEndWhile Chr(32), wdBackward
               End If
            End If
        'IF NO COLORED TEXT PRESENT
         Else
            .Expand Unit:=wdWord
            .MoveEndWhile Chr(32), wdBackward
         End If
      End If
     'IF SELECTED TEXT ALREADY THAT COLOR, THEN ERASE IT INSTEAD
      If (.HighlightColorIndex = HColor) And (MiddleOfWord = False) Then
        'E.G., HIGHLIGHT 'ETCHING', THEN TRY TO HIGHLIGHT 'ETCH' THE SAME COLOR
         If .HighlightColorIndex <> .Characters(1).HighlightColorIndex Then
            HColor = wdNoHighlight
            FColor = wdColorAutomatic
         End If
      End If
      End With
      OldColorIndex = Options.DefaultHighlightColorIndex
      Options.DefaultHighlightColorIndex = HColor
      With docRng.Find
      ' FIND CRITERIA
         .ClearFormatting
      ' REPLACE CRITERIA
         .Replacement.ClearFormatting
         .Replacement.Font.color = FColor
         .Replacement.Highlight = True
      ' EXECUTION
         .Execute Replace:=wdReplaceAll, FindText:=txtRng.Text, ReplaceWith:="", _
            FORMAT:=True, _
            MatchAllWordForms:=False, _
            MatchCase:=False, _
            MatchWholeWord:=False, _
            MatchWildcards:=False, _
            MatchSoundsLike:=False, _
            Wrap:=wdFindContinue
      End With
   Options.DefaultHighlightColorIndex = OldColorIndex
ErrorHandler:
   CursorShow
   Application.ScreenUpdating = True
   If Err <> 0 Then MsgBox "Error executing HighlightAndChangeFont() function."
   obj_undo.EndCustomRecord
End Function
'FONT COLOR, USING RANGES
Function FontColor(FColor)
   On Error GoTo ErrorHandler
   Set obj_undo = Application.UndoRecord
   obj_undo.StartCustomRecord ("FontColor()")
   Application.ScreenUpdating = False
   Dim txtRng As Range
   Dim docRng As Range
   Set txtRng = Selection.Range
   Set docRng = ActiveDocument.Range
   With txtRng
   'IF SELECTED TEXT ALREADY THAT COLOR, THEN ERASE IT INSTEAD
      If .Font.color = FColor Or .Characters.First.Previous.Font.color = FColor Then FColor = wdColorAutomatic
   
     'IF SELECTION, THEN ASSIGN TO RANGE, TRIM RANGE, COLLAPSE SELECTION
      If Selection.Type <> wdSelectionIP Then
         .MoveEndWhile Cset:=(Chr$(10) & Chr$(32) & Chr$(13) & ".,?!"), Count:=-.Characters.Count
         .MoveStartWhile Cset:=(Chr$(10) & Chr$(32) & Chr$(13) & ".,?!"), Count:=.Characters.Count
         If .Words.Count = 1 Then
            Selection.collapse Direction:=wdCollapseEnd
         Else
            Selection.collapse Direction:=wdCollapseStart
         End If
     'BEHAVIOR IF NO SELECTION, BUT FORMATTED TEXT NEARBY
      ElseIf Selection.Type = wdSelectionIP Then
         If .Font.color <> wdColorAutomatic Then
            .Find.Font.color = .Font.color
           'IP CCT LEFT OF CCT OR BETWEEN DIFFERENT CCT1|CCT2, EXPAND RIGHT
            If .Characters.First.Previous.Font.color = wdColorAutomatic Or _
               .Characters.First.Previous.Font.color <> .Font.color Then
               .Find.Forward = True
               .Find.Execute
           'IP WITHIN CCT, EXPAND BOTH DIRECTIONS
            ElseIf .Font.color <> wdColorAutomatic Then
               .Find.Forward = True
               .Find.Execute
               .collapse Direction:=wdCollapseEnd
               .Find.Forward = False
               .Find.Execute
               .Find.Forward = True
            End If
        'IP RIGHT OF CCT AND LEFT OF UNCOLORED TEXT, EXPAND LEFT
         ElseIf .Font.color = wdColorAutomatic And .Characters.First.Previous.Font.color <> wdColorAutomatic Then
            .Find.Font.color = .Characters.First.Previous.Font.color
            .Find.Forward = False
            .Find.Execute
            .Find.Forward = True
        'IF NO COLORED TEXT PRESENT
         Else
             .Expand Unit:=wdWord
             .MoveEndWhile Chr(32), wdBackward
          End If
      End If
   End With
   'REMOVE TRAILING WHITE SPACE
   With docRng.Find
   ' FIND CRITERIA
      .ClearFormatting
   ' REPLACE CRITERIA
      .Replacement.ClearFormatting
      .Replacement.Font.color = FColor
   ' EXECUTION
      .Execute Replace:=wdReplaceAll, FindText:=txtRng.Text, ReplaceWith:="", _
         FORMAT:=True, _
         MatchAllWordForms:=False, _
         MatchCase:=False, _
         MatchWholeWord:=False, _
         MatchWildcards:=False, _
         MatchSoundsLike:=False, _
         Wrap:=wdFindContinue
   End With
ErrorHandler:
   Application.ScreenUpdating = True
   If Err <> 0 Then MsgBox "Error executing FontColor() function."
   obj_undo.EndCustomRecord
End Function
Sub FindDelete()
   'Searches for and deletes all instances of selected text.
   'Trim ends of selection is intentionally not part of the subroutine.
   On Error GoTo ErrorHandler
   SetSearchSettings ("Find & Delete")
   Application.ScreenUpdating = False
   If Selection.Type = wdSelectionIP Then
      Selection.Expand Unit:=wdWord
   End If
   Selection.Find.FORMAT = False
   Selection.Find.Text = Selection
   Selection.Find.Execute Replace:=wdReplaceAll
ErrorHandler:
   Application.ScreenUpdating = True
   If Err <> 0 Then MsgBox "Data on clipboard is not text or is empty"
   RestoreSettings
End Sub
'FIND & REPLACE WITH CLIPBOARD TEXT
Sub FindReplaceClipboard()
   'Searches for and replaces with clipboard contents all instances of selected text.
   'Trim ends of selection is intentionally not part of the subroutine.
   SetSearchSettings ("Find & Delete")
   Application.ScreenUpdating = False
   Dim cbContainer As DataObject
   Dim cbText As String
   Set cbContainer = New DataObject
   cbContainer.GetFromClipboard
   cbText = cbContainer.GetText
   If Selection.Type = wdSelectionIP Then
      Selection.Expand Unit:=wdWord
   End If
   On Error GoTo ErrorHandler
   Selection.Find.Replacement.Text = cbText
   On Error Resume Next
   Selection.MoveEndWhile Chr(32), wdBackward
   Selection.Find.FORMAT = False
   Selection.Find.Text = Selection
   Selection.Find.Execute Replace:=wdReplaceAll
ErrorHandler:
   If Err <> 0 Then MsgBox "Data on clipboard is not text or is empty"
   Application.ScreenUpdating = True
   RestoreSettings
End Sub
Sub RestoreOrigFont()
   On Error GoTo ErrorHandler
   Set obj_undo = Application.UndoRecord
   obj_undo.StartCustomRecord ("RestoreOrigFont()")
   Application.ScreenUpdating = True
   FontColor (wdColorAutomatic)
   If Selection.Type = wdSelectionIP And Selection.Font.ColorIndex = 0 Then
      RemoveHighlightingFromParagraph 'Selection.PARAGRAPHS(1).Range.Font.ColorIndex = wdAuto
      Exit Sub
   Else
      FontColor (wdColorAutomatic)
   End If
ErrorHandler:
   Application.ScreenUpdating = True
   If Err <> 0 Then MsgBox "Error executing RestoreOrigFont() function."
   obj_undo.EndCustomRecord
   End Sub
Sub UnHighlight()
   On Error GoTo ErrorHandler
   Set obj_undo = Application.UndoRecord
   obj_undo.StartCustomRecord ("UnHighlight()")
   Application.ScreenUpdating = False
   Dim OldColorIndex As String
   OldColorIndex = Options.DefaultHighlightColorIndex
   Options.DefaultHighlightColorIndex = wdNoHighlight
   If Selection.Type = wdSelectionIP And Selection.Range.HighlightColorIndex = 0 Then
      RemoveHighlightingFromParagraph
      Exit Sub
   Else
      Call HighlightAndChangeFont(wdNoHighlight, wdColorAutomatic)
   End If
   Options.DefaultHighlightColorIndex = OldColorIndex
ErrorHandler:
   Application.ScreenUpdating = True
   If Err <> 0 Then MsgBox "Error executing UnHighlight() function."
   obj_undo.EndCustomRecord
   End Sub
