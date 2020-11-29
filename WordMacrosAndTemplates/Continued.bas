Attribute VB_Name = "Continued"
Sub Continued()
'
' Continued Macro
' Adds (CONTINUED) and CONTINUED:
'
    Selection.TypeParagraph
    Selection.Style = ActiveDocument.Styles("sAction")
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.Style = ActiveDocument.Styles("sDISSOLVE:")
    Selection.TypeText Text:="(CONTINUED)"
    Selection.InsertBreak Type:=wdPageBreak
    Selection.Style = ActiveDocument.Styles("sSlugline")
    Selection.TypeText Text:=vbTab & "continued: ()" & vbTab & vbTab
    Selection.TypeParagraph
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
End Sub
