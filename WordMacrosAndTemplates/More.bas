Attribute VB_Name = "More"
Sub More()
'
' More Macro
' Breaks the page and adds (MORE) and (CONT'D)
'
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText Text:="(MORE)"
    Selection.Style = ActiveDocument.Styles("More")
    Selection.InsertBreak Type:=wdPageBreak
    Selection.Style = ActiveDocument.Styles("Character")
    Selection.TypeText Text:=" (CONT'D)"
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.HomeKey Unit:=wdLine
End Sub
