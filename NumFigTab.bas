' Developed by Chieh (Ross) Wang in 2016
' This macro is free to anyone who wish to use, modify, expand, and share.
' NumFigTab Macro
' Automatically count number of figures+tables (Style: Caption)
'

Sub NumFigTab(Control As IRibbonControl)
    Dim intCount As Integer
    Dim captionCount As Long
    
    intCount = 1
    captionCount = 0
       
    With ActiveDocument.Range
        Do
            If .Paragraphs(index:=intCount).style = "Caption" Then
                captionCount = captionCount + 1
            End If
            intCount = intCount + 1
        Loop Until intCount = .Paragraphs.Count
    End With
    
    For Each prop In ActiveDocument.CustomDocumentProperties
        If prop.Name = "FigTabCount" Then
            prop.Delete
        End If
    Next
    
    With ActiveDocument.CustomDocumentProperties
        .Add Name:="FigTabCount", _
            LinkToContent:=False, _
            Type:=msoPropertyTypeNumber, _
            Value:=captionCount
    End With
    
    Selection.WholeStory
    Selection.Fields.Update
End Sub
