Attribute VB_Name = "MouleBaidu"
Sub TestWebRequest()
    Dim codeStr As String
    Dim ret As Boolean
    codeStr = GetCodeStringByRequest("这是你的测试语句内容")
    If Len(codeStr) > 0 Then
        ret = RunDynamicCode(codeStr)
    Else
        ret = False
    End If
    
End Sub

' 获取代码
Function GetCodeStringByRequest(inputStr As String) As String

    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    
    On Error GoTo ErrorHandler
    
    ' 初始化 HTTP 对象
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' 设置请求的 URL
    url = "https://qianfan.baidubce.com/v2/chat/completions"
    
    
    ' 设置请求体
    requestBody = "{""model"":""deepseek-v3"",""messages"":[{""role"":""user"",""content"":""VBA powerpoint，" + inputStr + "。不需要输出任何解释文本和引导内容，直接输出vba代码，且不可以有任何markedown标识，直接输出文本内容。""}]}"
    
	Dim appId as String
	Dim apiKey as String
	appId = "你的APPID"
	apiKey = "你的API Key"'特别注意：APIkey前面有一个“Bearer ”
	
    ' 发送 POST 请求
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "appid", appId
        .setRequestHeader "Authorization", apiKey
        .send requestBody
    End With
    
    ' 获取响应内容
    response = http.responseText
    
    ' 输出响应内容
    Debug.Print response
    content = GetJsonParsing(response)
    
    ' 输出 content
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


' 运行代码
Function RunDynamicCode(incodeStr As String) As Boolean

    On Error GoTo ErrorHandler
    
    Dim codeStr As String
    codeStr = Replace(incodeStr, "\n", vbCrLf)
    ' 从代码字符串中提取过程名
    Dim procName As String
    procName = ExtractProcedureName(codeStr)
    
    If procName = "" Then
        MsgBox "无法从代码字符串中提取过程名！", vbCritical
        Exit Function
    End If
    
    ' 获取当前演示文稿的 VBProject
    Dim vbProj As VBProject
    Set vbProj = Application.VBE.ActiveVBProject
    
    ' 创建一个新的标准模块
    Dim vbComp As VBComponent
    Set vbComp = vbProj.VBComponents.Add(vbext_ct_StdModule)
    
    ' 将代码字符串添加到模块中
    vbComp.CodeModule.AddFromString codeStr
    
    ' 运行动态添加的子过程
    Application.Run procName
    
    ' 删除动态创建的模块
    vbProj.VBComponents.Remove vbComp
    
    RunDynamicCode = True
    
   Exit Function
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    ' 确保模块被删除
    If Not vbComp Is Nothing Then
        vbProj.VBComponents.Remove vbComp
    End If
     RunDynamicCode = False
     
End Function

Function ExtractProcedureName(codeStr As String) As String
    ' 使用正则表达式提取过程名
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' 匹配 Sub 后的过程名
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


' json读取值
Function GetJsonParsing(JsonString As String) As String

    Dim jsonDict As Object
    Dim content As String

    ' 解析 JSON 字符串
    On Error Resume Next
    Set jsonDict = JsonConverter.ParseJson(JsonString)
    On Error GoTo 0
    
    ' 检查是否解析成功
    If jsonDict Is Nothing Then
        MsgBox "JSON 解析失败！"
        Exit Function
    End If
    
    ' 提取 content 字段中的代码
    content = jsonDict("choices")(1)("message")("content")
    
    ' 替换 \n 为实际换行符
    content = Replace(content, "\n", vbCrLf)
    
    ' 输出结果
    Debug.Print "提取的代码：" & vbCrLf & content
    
    GetJsonParsing = content
   
End Function
