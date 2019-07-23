Attribute VB_Name = "WordAddin_SearchProperties"
'README - THERE IS MORE TO DO HERE.  I THINK YOU CAN AVOID THE WORD
' SELECTION ALMOST ALTOGETHER. JUST SET hl_def_ss.TEXT = SELECTION.TEXT
' IN FACT, THIS IS WHAT YOU WANT TO DO, SO THAT THE USER'S REGULAR
' FIND AND REPLACE ACTIVITY IS NOT AFFECTED BY THEIR USE OF THE APP.
 Dim obj_undo As UndoRecord
 Public search_settings_initialized As Boolean
 Public hl_def_ss As Find
 Public g_Foward, g_FORMAT, g_IgnorePunct, g_IgnoreSpace, _
      g_MatchCase, g_MatchWholeWord, g_MatchWildcards, _
      g_MatchSoundsLike, g_MatchAllWordForms _
      As Boolean
Sub ToggleMatchWholeWord()
   Selection.Find.MatchWholeWord = Not Selection.Find.MatchWholeWord
   End Sub
Sub ToggleMatchCase()
   Selection.Find.MatchCase = Not Selection.Find.MatchCase
   End Sub
Sub ToggleIgnorePunct()
   Selection.Find.IgnorePunct = Not Selection.Find.IgnorePunct
End Sub
Sub ToggleFindContinueFindStop()
   If Selection.Find.Wrap = wdFindStop Then
      Selection.Find.Wrap = wdFindContinue
   Else
      Selection.Find.Wrap = wdFindStop
   End If
   End Sub
Sub set_HL_Search_MatchWildCards()
   hl_def_ss.MatchWildcards = True
   MsgBox hl_def_ss.MatchWildcards
   End Sub
Sub Toggle_MatchWildcards()
   Selection.Find.MatchWildcards = Not Selection.Find.MatchWildcards
   End Sub
Sub initialize_highlighter_search_settings()
   Set hl_def_ss = ActiveDocument.Content.Find
   hl_def_ss.Text = ""
   hl_def_ss.Replacement.Text = ""
   hl_def_ss.Forward = True
   hl_def_ss.Wrap = wdFindContinue
   hl_def_ss.FORMAT = True
'   hl_def_ss.IgnorePunct = True
'   hl_def_ss.IgnoreSpace = True
   hl_def_ss.MatchCase = False
   hl_def_ss.MatchWholeWord = False
   hl_def_ss.MatchWildcards = False
   hl_def_ss.MatchSoundsLike = False
   hl_def_ss.MatchAllWordForms = False
   search_settings_initialized = True
   'MsgBox "Search Settings Initialized."
   End Sub
'SET HIGHLIGHTER SEARCH SETTINGS
'NOTE - Track changes does not record changes to highlighting. Turning wildcards on automatically sets WholeWordsOnly and MatchCase to False.
Sub SetSearchSettings(SubName As String)
   Set obj_undo = Application.UndoRecord
   obj_undo.StartCustomRecord (SubName & "()")
   If search_settings_initialized = False Then initialize_highlighter_search_settings
   With Selection.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .Text = hl_def_ss.Text
      .Replacement.Text = hl_def_ss.Replacement.Text
      .Forward = hl_def_ss.Forward
'      .IgnorePunct = hl_def_ss.IgnorePunct
'      .IgnoreSpace = hl_def_ss.IgnoreSpace
      .Wrap = hl_def_ss.Wrap
      .FORMAT = hl_def_ss.FORMAT
      .MatchCase = hl_def_ss.MatchCase
      .MatchWholeWord = hl_def_ss.MatchWholeWord
      .MatchWildcards = hl_def_ss.MatchWildcards
      .MatchSoundsLike = hl_def_ss.MatchSoundsLike
      .MatchAllWordForms = hl_def_ss.MatchAllWordForms
   End With
   End Sub
'RESTORE SETTINGS
Sub RestoreSettings()
   Selection.Find.MatchAllWordForms = False
   Selection.Find.MatchWildcards = False
   Selection.Find.FORMAT = False
   obj_undo.EndCustomRecord
   End Sub
