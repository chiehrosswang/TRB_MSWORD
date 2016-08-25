Attribute VB_Name = "NumFigTab"
Sub NumFigTab()
Attribute NumFigTab.VB_Description = "Automatically count number of figures+tables (Style: Caption)"
Attribute NumFigTab.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.NumFigTab"
' Developed by Chieh (Ross) Wang in 2016
' This macro is free to anyone who wish to use and/or modify as long as the above line is kept in the macro file.
' NumFigTab Macro
' Automatically count number of figures+tables (Style: Caption)
'
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

