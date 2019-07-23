Attribute VB_Name = "WordAddin_Shade_Colors"
'SCROLL & COLLAPSE OPTIONS
'   Selection.MOVE Unit:=wdParagraph, Count:=1 'move to next para
'   Selection.MOVE Unit:=wdParagraph, Count:=-1 'start of this para
'   ScrollToUpperLeftCorner
Sub Shading_Black()
   Shade (-587137025)
   End Sub
Sub Shading_Blue()
   Shade (-738132122)
   End Sub
Sub Shading_AcquaBlue()
   Shade (12894220)
   End Sub
Sub Shading_BrightBlue()
   Shade (15773696)
   End Sub
Sub Shading_Green()
   'shade (13300954)
   Shade (10672756)
   End Sub
Sub Shading_BrightGreen()
   Shade (5287936)
   End Sub
Sub Shading_SpringGreen()
   Shade (11272065)
   End Sub
Sub Shading_Grey()
    Shade (-603923969)
    Selection.MOVE Unit:=wdParagraph, Count:=1 'move to next para
   End Sub
Sub Shading_Gray()
    Shade (-603923969)
    Selection.MOVE Unit:=wdParagraph, Count:=1 'move to next para
   End Sub
Sub Shading_Orange()
   Shade (10996734)
   End Sub
Sub Shading_BrightOrange()
   Shade (6004223)
   End Sub
Sub Shading_Red()
   Shade (8224255)
   End Sub
Sub Shading_Pink()
   'shade (16735231)
   Shade (16754687)
   'shade (wdColorPink)
   'shade (12616171) 'dull pink
   End Sub
Sub Shading_Purple()
   'shade (16740836)
   Shade (15316435)
   'shade (14598342)
   End Sub
Sub Shading_Tan()
   Shade (13427942)
   Selection.MOVE Unit:=wdParagraph, Count:=1 'move to next para
   End Sub
Sub Shading_Yellow()
   Shade (13434879)
   End Sub
Sub Shading_BrightYellow()
   Shade (6750207)
   End Sub
