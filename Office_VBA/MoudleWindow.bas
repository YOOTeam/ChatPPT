Attribute VB_Name = "MoudleWindow"
Dim myForm As DeepSeekTool
Sub OpenTool()

    Set myForm = New DeepSeekTool
    myForm.Show
    
End Sub

Sub NoteInput()

    Dim noteStr As String
    noteStr = GetSlideNotesText
    If Len(noteStr) = 0 Then
        Exit Sub
    End If
    
    Dim codeStr As String
    Dim ret As Boolean
    codeStr = GetCodeStringByRequest(noteStr)
    
    If Len(codeStr) > 0 Then
        ret = RunDynamicCode(codeStr)
        SetSlideNotesText (codeStr)
    Else
        ret = False
    End If
    
End Sub


Function GetSlideNotesText() As String

    Dim slide As slide
    Dim notesText As String
    Dim shp As shape
    Dim foundNotes As Boolean
    
    ' 获取当前选中的幻灯片
    Set slide = ActiveWindow.Selection.SlideRange(1)
        
    ' 初始化标志
    foundNotes = False
        
    ' 遍历备注页中的所有形状
    For Each shp In slide.NotesPage.Shapes
        ' 检查形状是否有文本框并且包含文本
        If shp.HasTextFrame And shp.TextFrame.HasText Then
            notesText = shp.TextFrame.textRange.Text
            foundNotes = True
            Exit For ' 找到备注文本后退出循环
        End If
    Next shp
        
    ' 输出备注文本
    If foundNotes Then
        GetSlideNotesText = notesText
    Else
        MsgBox "未找到备注内容。"
        GetSlideNotesText = ""
    End If

End Function


Function SetSlideNotesText(noteStr As String) As String

    Dim slide As slide
    Dim notesText As String
    Dim shp As shape
    Dim foundNotes As Boolean
    
   
    ' 获取当前选中的幻灯片
    Set slide = ActiveWindow.Selection.SlideRange(1)
        
    ' 初始化标志
    foundNotes = False
        
    ' 遍历备注页中的所有形状
    For Each shp In slide.NotesPage.Shapes
        ' 检查形状是否有文本框
        If shp.HasTextFrame Then
            ' 设置备注文本
            shp.TextFrame.textRange.Text = noteStr
            foundNotes = True
            Exit For ' 找到文本框后退出循环
        End If
    Next shp
        
    ' 如果没有找到文本框，添加一个新的文本框
    If Not foundNotes Then
        Set shp = slide.NotesPage.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 500, 200)
        shp.TextFrame.textRange.Text = noteStr
    End If
End Function
