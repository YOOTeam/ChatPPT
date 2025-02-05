Attribute VB_Name = "MouleBaidu"
Sub TestWebRequest()
    Dim codeStr As String
    Dim ret As Boolean
    codeStr = GetCodeStringByRequest("������Ĳ����������")
    If Len(codeStr) > 0 Then
        ret = RunDynamicCode(codeStr)
    Else
        ret = False
    End If
    
End Sub

' ��ȡ����
Function GetCodeStringByRequest(inputStr As String) As String

    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    
    On Error GoTo ErrorHandler
    
    ' ��ʼ�� HTTP ����
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' ��������� URL
    url = "https://qianfan.baidubce.com/v2/chat/completions"
    
    
    ' ����������
    requestBody = "{""model"":""deepseek-v3"",""messages"":[{""role"":""user"",""content"":""VBA powerpoint��" + inputStr + "������Ҫ����κν����ı����������ݣ�ֱ�����vba���룬�Ҳ��������κ�markedown��ʶ��ֱ������ı����ݡ�""}]}"
    
	Dim appId as String
	Dim apiKey as String
	appId = "���APPID"
	apiKey = "���API Key"'�ر�ע�⣺APIkeyǰ����һ����Bearer ��
	
    ' ���� POST ����
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "appid", appId
        .setRequestHeader "Authorization", apiKey
        .send requestBody
    End With
    
    ' ��ȡ��Ӧ����
    response = http.responseText
    
    ' �����Ӧ����
    Debug.Print response
    content = GetJsonParsing(response)
    
    ' ��� content
    Debug.Print content
    
    Dim codeStr As String
    codeStr = Replace(content, "\n", vbCrLf)
    
    Debug.Print codeStr
    GetCodeStringByRequest = codeStr
    
    Exit Function
    
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    GetCodeStringByRequest = ""
     
End Function


' ���д���
Function RunDynamicCode(incodeStr As String) As Boolean

    On Error GoTo ErrorHandler
    
    Dim codeStr As String
    codeStr = Replace(incodeStr, "\n", vbCrLf)
    ' �Ӵ����ַ�������ȡ������
    Dim procName As String
    procName = ExtractProcedureName(codeStr)
    
    If procName = "" Then
        MsgBox "�޷��Ӵ����ַ�������ȡ��������", vbCritical
        Exit Function
    End If
    
    ' ��ȡ��ǰ��ʾ�ĸ�� VBProject
    Dim vbProj As VBProject
    Set vbProj = Application.VBE.ActiveVBProject
    
    ' ����һ���µı�׼ģ��
    Dim vbComp As VBComponent
    Set vbComp = vbProj.VBComponents.Add(vbext_ct_StdModule)
    
    ' �������ַ�����ӵ�ģ����
    vbComp.CodeModule.AddFromString codeStr
    
    ' ���ж�̬��ӵ��ӹ���
    Application.Run procName
    
    ' ɾ����̬������ģ��
    vbProj.VBComponents.Remove vbComp
    
    RunDynamicCode = True
    
   Exit Function
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    ' ȷ��ģ�鱻ɾ��
    If Not vbComp Is Nothing Then
        vbProj.VBComponents.Remove vbComp
    End If
     RunDynamicCode = False
     
End Function

Function ExtractProcedureName(codeStr As String) As String
    ' ʹ��������ʽ��ȡ������
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' ƥ�� Sub ��Ĺ�����
    regex.IgnoreCase = True
    regex.Global = False
    regex.Pattern = "Sub\s+([a-zA-Z_][a-zA-Z0-9_]*)"
    
    Dim matches As Object
    Set matches = regex.Execute(codeStr)
    
    If matches.Count > 0 Then
        ExtractProcedureName = matches(0).SubMatches(0)
    Else
        ExtractProcedureName = ""
    End If
End Function


' json��ȡֵ
Function GetJsonParsing(JsonString As String) As String

    Dim jsonDict As Object
    Dim content As String

    ' ���� JSON �ַ���
    On Error Resume Next
    Set jsonDict = JsonConverter.ParseJson(JsonString)
    On Error GoTo 0
    
    ' ����Ƿ�����ɹ�
    If jsonDict Is Nothing Then
        MsgBox "JSON ����ʧ�ܣ�"
        Exit Function
    End If
    
    ' ��ȡ content �ֶ��еĴ���
    content = jsonDict("choices")(1)("message")("content")
    
    ' �滻 \n Ϊʵ�ʻ��з�
    content = Replace(content, "\n", vbCrLf)
    
    ' ������
    Debug.Print "��ȡ�Ĵ��룺" & vbCrLf & content
    
    GetJsonParsing = content
   
End Function
