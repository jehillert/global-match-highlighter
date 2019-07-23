Attribute VB_Name = "WordAddin_Colors"
'HIGHLIGHTER
Sub highlight_blue()
   Call HighlightAndChangeFont(wdBlue, wdColorWhite)
   End Sub
Sub highlight_bright_green()
   Call HighlightAndChangeFont(wdBrightGreen)
   End Sub
Sub highlight_dark_gray()
   Call HighlightAndChangeFont(wdGray50, wdColorWhite)
   End Sub
Sub highlight_light_gray()
   Call HighlightAndChangeFont(wdGray25)
   End Sub
Sub highlight_pink()
   Call HighlightAndChangeFont(wdPink, wdColorWhite)
   End Sub
Sub highlight_red()
   Call HighlightAndChangeFont(wdRed)
   End Sub
Sub highlight_turquoise()
   Call HighlightAndChangeFont(wdTurquoise)
   End Sub
Sub highlight_yellow()
   Call HighlightAndChangeFont(wdYellow)
   End Sub
Sub highlight_black()
   Call HighlightAndChangeFont(wdBlack, wdColorWhite)
   End Sub
Sub highlight_dark_blue()
   Call HighlightAndChangeFont(wdDarkBlue, wdColorWhite)
   End Sub
Sub highlight_dark_red()
   Call HighlightAndChangeFont(wdDarkRed, wdColorWhite)
   End Sub
Sub highlight_dark_yellow()
   Call HighlightAndChangeFont(wdDarkYellow, wdColorWhite)
   End Sub
Sub highlight_green()
   Call HighlightAndChangeFont(wdGreen, wdColorWhite)
   End Sub
Sub highlight_red_white()
'   start_timer
   Call HighlightAndChangeFont(wdRed, wdColorWhite)
'   stop_timer
   End Sub
Sub highlight_teal()
   Call HighlightAndChangeFont(wdTeal, wdColorWhite)
   End Sub
Sub highlight_violet()
   Call HighlightAndChangeFont(wdViolet, wdColorWhite)
   End Sub
Sub HighlightWHITE()
   Call HighlightAndChangeFont(wdWhite, wdColorWhite)
End Sub

'TEXT
Sub TextBlack()
   FontColor (wdColorBlack)
   End Sub
Sub TextBlue()
   FontColor (wdColorBlue)
   End Sub
Sub TextBrown()
   FontColor (wdColorBrown)
   End Sub
Sub TextDarkRed()
   FontColor (wdColorDarkRed)
   End Sub
Sub TextGold()
   FontColor (wdColorGold)
   End Sub
Sub TextGreen()
   FontColor (10132992)
   'FontColor (3437568)
   End Sub
Sub TextOrange()
   FontColor (wdColorOrange)
   End Sub
Sub TextPink()
   FontColor (wdColorPink)
   End Sub
Sub TextPurple()
   FontColor (10498160)
   End Sub
Sub TextRed()
   FontColor (wdColorRed)
   End Sub

