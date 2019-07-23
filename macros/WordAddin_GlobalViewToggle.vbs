Attribute VB_Name = "WordAddin_GlobalViewToggle"
Private Const globalZoomOverride = 28
Private webViewZoomPercent
Private wasDisplayBackgroundsOn, _
      wasGrammarOn, _
      wasSpellingOn, _
      wasDocMapOn, _
      wasRulerOn _
      As Boolean

Sub toggleGlobalViewState()
   
   Set objUndo = Application.UndoRecord
   objUndo.StartCustomRecord ("toggleGlobalViewState")
   Set myDoc = ActiveWindow.Document
   Set myWindow = ActiveWindow
   
   Dim row, col, pages, ribbonWidth As Integer
   
   On Error GoTo ErrorHandler
   
   myWindow.Visible = False
   myWindow.WindowState = wdWindowStateMaximize
   
   If myDoc.ActiveWindow.View.Type <> wdPrintView Or myDoc.ActiveWindow.View.Zoom.PageColumns = 1 Then
'        CAPATURE VIEW STATE
      webViewZoomPercent = myDoc.ActiveWindow.View.Zoom.Percentage
      wasDocMapOn = myDoc.ActiveWindow.DocumentMap
      wasGrammarOn = Options.CheckGrammarAsYouType
      wasSpellingOn = Options.CheckSpellingAsYouType
      wasRulerOn = myDoc.ActiveWindow.ActivePane.DisplayRulers
      wasDisplayBackgroundsOn = myDoc.ActiveWindow.View.DisplayBackgrounds
'        GLOBAL VIEW STATE
      ribbonWidth = CommandBars("Ribbon").Controls(1).Height
      If ribbonWidth = 194 Then CommandBars.ExecuteMso ("MinimizeRibbon")
      pages = myDoc.Range.Information(wdNumberOfPagesInDocument)
      If pages = 2 Then
         col = 2
         row = 1
      ElseIf pages = 3 Then
         col = 3
         row = 1
      ElseIf pages = 4 Then
         col = 2
         row = 2
      ElseIf pages = 5 Or pages = 6 Then
         col = 3
         row = 2
      ElseIf pages >= 7 Then
         col = 4
         row = pages \ col
         If pages Mod col > 0 Then
          row = row + 1
         End If
      End If
      myDoc.ActiveWindow.View.Type = wdPrintView
      myDoc.ActiveWindow.View.Zoom.PageColumns = col
      myDoc.ActiveWindow.View.Zoom.PageRows = row
      myDoc.ActiveWindow.View.Zoom.Percentage = globalZoomOverride
      myDoc.ActiveWindow.ActivePane.DisplayRulers = False
      myDoc.ActiveWindow.DocumentMap = False
      myDoc.ActiveWindow.View.DisplayBackgrounds = True
   Else 'INITIAL VIEW STATE
      ribbonWidth = CommandBars("Ribbon").Controls(1).Height
      If ribbonWidth = 108 Then CommandBars.ExecuteMso ("MinimizeRibbon")
      If myDoc.ActiveWindow.View.Type = wdPrintView Then
         myDoc.ActiveWindow.View.Zoom.PageColumns = 1
         myDoc.ActiveWindow.View.Zoom.Percentage = 100
'         ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitFullPage
      End If
      myDoc.ActiveWindow.DocumentMap = wasDocMapOn
      Options.CheckGrammarAsYouType = wasGrammarOn
      Options.CheckSpellingAsYouType = wasSpellingOn
      myDoc.ActiveWindow.View.Type = wdWebView
      myDoc.ActiveWindow.ActivePane.DisplayRulers = wasRulerOn
'      myDoc.ActiveWindow.View.Zoom.Percentage = webViewZoomPercent
      myDoc.ActiveWindow.View.DisplayBackgrounds = wasDisplayBackgroundsOn
   End If
   Application.ScreenRefresh
   myWindow.Visible = True
   Selection.MoveLeft Unit:=wdCharacter, Count:=1
   objUndo.EndCustomRecord
Exit Sub
ErrorHandler:
   MsgBox "Error in function 'toggleGlobalViewState.'"
   Application.ScreenRefresh
   myWindow.Visible = True
   objUndo.EndCustomRecord
End Sub
