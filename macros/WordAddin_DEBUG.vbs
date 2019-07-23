Attribute VB_Name = "WordAddin_DEBUG"
Sub areFullPargraphsSelected()

   Dim selection1 As Range
   Dim selection2 As Range
   
   MarkSpot ("mySpotB4FullParagraphsSelectedWasRan")
   
   Set selection1 = Selection.Range
   Selection.Expand Unit:=wdParagraph
   Set selection2 = Selection.Range

   If selection1 = selection2 Then
    MsgBox "yes"
   Else
    MsgBox "no"
   End If
   ReturnToSpot ("mySpotB4FullParagraphsSelectedWasRan")
   
End Sub
