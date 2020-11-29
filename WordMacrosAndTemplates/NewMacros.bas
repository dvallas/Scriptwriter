Attribute VB_Name = "NewMacros"
Sub Continued()
Attribute Continued.VB_Description = "Macros created 12/16/98 by Dean Vallas"
Attribute Continued.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.Continued"
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
Sub Name()
'
' Character Macro
' Macro created 12/18/97 by Dean Vallas
'
WordBasic.Style "sCharacter Name"
If this.KeyPress() = wdKeyReturn Then
    MsgBox "name pressed"
    WordBasic.InsertPara
    WordBasic.Style "sDialog"
End If

End Sub
Sub Dialog()
Attribute Dialog.VB_Description = "Macro created 05/09/97 by Dean Vallas"
Attribute Dialog.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.Dialog"
'
' Dialog Macro
' Macro created 05/09/97 by Dean Vallas
'
WordBasic.Style "sDialog"
If this.KeyPress() = wdKeyReturn Then
    WordBasic.InsertPara
    WordBasic.Style "sSlugline"
End If

End Sub

Sub CreateRoadmapWindow()
Attribute CreateRoadmapWindow.VB_Description = "Macro recorded 05/09/97 by Dean Vallas"
Attribute CreateRoadmapWindow.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.CreateRoadmapWindow"
'
' CreateRoadmapWindow Macro
' Macro recorded 05/09/97 by Dean Vallas
'
Dim strRoadmap As String
Selection.HomeKey Unit:=wdStory
'
' Dialog Macro
' Macro created 05/09/97 by Dean Vallas
'

'Do While True
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("sSlugline")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Selection.Find.Execute
    Selection.Copy
    strRoadmap = strRoadmap & vbCrLf & Selection
    End With
   '     Exit Do
    
    'Loop
    MsgBox (strRoadmap)
End Sub
Sub InsertParenthetical()
Attribute InsertParenthetical.VB_Description = "Macro recorded 12/19/97 by Dean Vallas"
Attribute InsertParenthetical.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.InsertParenthetical"
'
' InsertParenthetical Macro
' Macro recorded 12/19/97 by Dean Vallas
'
    Selection.Style = ActiveDocument.Styles("sDirection")
    Selection.TypeText Text:="()"
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub UPPER()
Attribute UPPER.VB_Description = "Change Case to UPPER"
Attribute UPPER.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.UPPER"
'
' UPPER Macro
' Change Case to UPPER
'
    Selection.Range.Case = wdUpperCase
End Sub
Sub lower()
Attribute lower.VB_Description = "Change Case to lower"
Attribute lower.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.lower"
'
' lower Macro
' Change Case to lower
'
    Selection.Range.Case = wdLowerCase
End Sub
Sub Macro1()
Attribute Macro1.VB_Description = "Add Name to AutoText List"
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.Macro1"
'
' Macro1 Macro
' Add All Names to AutoText List
'
    Selection.Style = ActiveDocument.Styles("sCharacter Name")
    'Selection.TypeText Text:="joe"
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    Selection.Copy
    Application.DisplayAutoCompleteTips = False
End Sub
Sub AddNameToAutoText()
Attribute AddNameToAutoText.VB_Description = "Macro recorded 12/20/98 by Dean Vallas"
Attribute AddNameToAutoText.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.Macro2"
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
    Selection.TypeText Text:=Paragraph
    
End Sub

Sub CollectNamesToMsgBox()
'
' CollectNamesToMsgBox Macro
' Collect Character Names and Display
'
Dim strMsg As String

For Each i In ActiveDocument.AttachedTemplate.AutoTextEntries
   strMsg = strMsg & UCase(i.Name) & Chr(13) & Chr(10)
Next i

MsgBox strMsg

End Sub
