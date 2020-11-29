Attribute VB_Name = "AddNameToAutoText"
Sub AddNameToAutoText()
'
' AddNameToAutoText Macro
' Macro written 12/20/98 by Dean Vallas
'
    Selection.Style = ActiveDocument.Styles("sCharacter Name")
    
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    'Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    ActiveDocument.AttachedTemplate.AutoTextEntries.Add _
    Name:=Selection, Range:=Selection.Range
    
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:=Paragraph
    'Selection.TypeText Text:=Paragraph
    
    Dim oAutoText As AutoTextEntry
  
    'Create an AutoText with the current selection as the content
    
    Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries _
        .Add(Name:=Selection.Text, Range:=Selection.Range)

    'Now replace the content of the AutoText with the desired value
    oAutoText.Value = Selection.Text
        
    'Clean up
    Set oAutoText = Nothing
    ActiveDocument.E
End Sub

