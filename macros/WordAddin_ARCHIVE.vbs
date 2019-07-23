Attribute VB_Name = "WordAddin_ARCHIVE"
Private Sub SearchReplace()
   Dim FindObject As Word.Find = Application.Selection.Find
   With FindObject
    .ClearFormatting()
    .Text = "find me"
    .Replacement.ClearFormatting()
    .Replacement.Text = "Found"
    .Execute(Replace:=Word.WdReplace.wdReplaceAll)
   End With
End Sub
Sub remember_doc_settings()
   With Selection.Find
   .ClearFormatting
   .Replacement.ClearFormatting
   .Text = ""
   .Replacement.Text = ""
   .Forward = True
   .Wrap = wdFindStop
   .FORMAT = False
   .MatchCase = False
   .MatchWholeWord = False
   .MatchWildcards = False
   .MatchSoundsLike = False
   .MatchAllWordForms = False
   End With
End Sub
