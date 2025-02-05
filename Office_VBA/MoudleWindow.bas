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
    
    ' ��ȡ��ǰѡ�еĻõ�Ƭ
    Set slide = ActiveWindow.Selection.SlideRange(1)
        
    ' ��ʼ����־
    foundNotes = False
        
    ' ������עҳ�е�������״
    For Each shp In slide.NotesPage.Shapes
        ' �����״�Ƿ����ı����Ұ����ı�
        If shp.HasTextFrame And shp.TextFrame.HasText Then
            notesText = shp.TextFrame.textRange.Text
            foundNotes = True
            Exit For ' �ҵ���ע�ı����˳�ѭ��
        End If
    Next shp
        
    ' �����ע�ı�
    If foundNotes Then
        GetSlideNotesText = notesText
    Else
        MsgBox "δ�ҵ���ע���ݡ�"
        GetSlideNotesText = ""
    End If

End Function


Function SetSlideNotesText(noteStr As String) As String

    Dim slide As slide
    Dim notesText As String
    Dim shp As shape
    Dim foundNotes As Boolean
    
   
    ' ��ȡ��ǰѡ�еĻõ�Ƭ
    Set slide = ActiveWindow.Selection.SlideRange(1)
        
    ' ��ʼ����־
    foundNotes = False
        
    ' ������עҳ�е�������״
    For Each shp In slide.NotesPage.Shapes
        ' �����״�Ƿ����ı���
        If shp.HasTextFrame Then
            ' ���ñ�ע�ı�
            shp.TextFrame.textRange.Text = noteStr
            foundNotes = True
            Exit For ' �ҵ��ı�����˳�ѭ��
        End If
    Next shp
        
    ' ���û���ҵ��ı������һ���µ��ı���
    If Not foundNotes Then
        Set shp = slide.NotesPage.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 500, 200)
        shp.TextFrame.textRange.Text = noteStr
    End If
End Function
