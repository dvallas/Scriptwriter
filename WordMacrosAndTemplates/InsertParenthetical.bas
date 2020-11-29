Attribute VB_Name = "InsertParenthetical"
Sub InsertParenthetical()
'
' InsertParenthetical Macro
' Macro recorded 12/19/97 by Dean Vallas
'
    Selection.Style = ActiveDocument.Styles("sDirection")
    Selection.TypeText Text:="()"
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
