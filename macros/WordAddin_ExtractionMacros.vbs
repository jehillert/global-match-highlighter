Attribute VB_Name = "WordAddin_ExtractionMacros"
Sub CopySentencesContaiingSpecificWordsToNewDoc()
   doStart ("CopySentencesContaiingSpecificWordsToNewDoc")
   Dim target_term As String
   Dim SourceDoc As String
   Dim TargetDoc As String
   SourceDoc = ActiveDocument.Name
   ChangeFileOpenDirectory ActiveDocument.Path
   target_term = InputBox("Enter target term: ")
   Documents.Add
   TargetDoc = "Sentences with [" & target_term & "]"
   ActiveDocument.SaveAs2 FileName:=TargetDoc, FileFormat:= _
      wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
      :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
      :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
      SaveAsAOCELetter:=False, CompatibilityMode:=15
   Windows(SourceDoc).Activate
   With Selection
      .HomeKey Unit:=wdStory
      '  Find the entered texts.
      With Selection.Find
         .ClearFormatting
         .Text = target_term
         .Replacement.Text = ""
         .Forward = True
         .Wrap = wdFindContinue
         .FORMAT = False
         .MatchCase = False
         .MatchWholeWord = False
         .MatchWildcards = False
         .MatchSoundsLike = False
         .MatchAllWordForms = True
         .Execute
      End With
      Do While .Find.Found = True
      '  Expand the selection to the entire sentence.
      Selection.Expand Unit:=wdSentence
      Selection.Copy
      Windows(TargetDoc).Activate
      Selection.Paste
      Selection.TypeParagraph
      Selection.TypeParagraph
      Windows(SourceDoc).Activate
      .collapse wdCollapseEnd
      .Find.Execute
      Loop
      Windows(TargetDoc).Activate
   End With
   doStop
End Sub
