Attribute VB_Name = "模块1"
' ========================================
' 宏名称: ExtractTablesToPPT
' 功能: 提取 Word 文档中指定页面的表格，并粘贴到 PowerPoint 中的幻灯片
' 作者: HW
' 版本: 1.0
' 创建日期: 2024/12/4
' 最后更新日期: 2024/12/4
' 更新历史:
'   v1.0 - 初始版本，实现基本功能
' ========================================



Sub ExtractTablesToPPT()
    Dim sourceDoc As Document
    Dim pptApp As Object
    Dim pptPresentation As Object
    Dim pptSlide As Object
    Dim specifiedPages As String
    Dim pagesArray() As String
    Dim pageRange As Variant
    Dim pageNum As Long, startPage As Long, endPage As Long
    Dim rng As Range
    Dim tbl As Table
    Dim i As Integer
    
    ' 提示用户输入页面范围
    specifiedPages = InputBox("请输入需要提取表格的页面范围（例如：1,2,3-5）：", "提取表格到PPT")
    If specifiedPages = "" Then
        MsgBox "未输入任何页面范围！", vbExclamation
        Exit Sub
    End If
    
    ' 设置当前文档为源文档
    Set sourceDoc = ActiveDocument

    ' 创建 PowerPoint 应用程序和新演示文稿
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application") ' 尝试获取 PowerPoint 应用
    If pptApp Is Nothing Then
        Set pptApp = CreateObject("PowerPoint.Application") ' 如果未运行，则创建新实例
    End If
    On Error GoTo 0
    pptApp.Visible = True
    Set pptPresentation = pptApp.Presentations.Add

    ' 分割输入的页面范围
    pagesArray = Split(specifiedPages, ",")
    i = 1
    
    ' 遍历每个指定的页面或范围
    For Each pageRange In pagesArray
        If InStr(pageRange, "-") > 0 Then
            ' 如果是范围（如3-5），提取起始和结束页面
            startPage = CLng(Split(pageRange, "-")(0))
            endPage = CLng(Split(pageRange, "-")(1))
        Else
            ' 如果是单个页面
            startPage = CLng(pageRange)
            endPage = startPage
        End If
        
        ' 检查页面范围是否有效
        If startPage <= 0 Or endPage < startPage Then
            MsgBox "页面范围无效：" & pageRange, vbExclamation
            Exit Sub
        End If
        
        ' 遍历范围内的每一页
        For pageNum = startPage To endPage
            ' 获取指定页面的范围
            Set rng = sourceDoc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=pageNum)
            rng.End = sourceDoc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=pageNum + 1).Start
            rng.End = rng.End - 1 ' 调整范围至页面结束
            
            ' 遍历页面内的表格
            For Each tbl In rng.Tables
                ' 复制表格
                tbl.Range.Copy
                
                ' 创建新幻灯片并粘贴表格
                Set pptSlide = pptPresentation.Slides.Add(Index:=pptPresentation.Slides.Count + 1, Layout:=12) ' 空白布局
                pptSlide.Shapes.Paste
                i = i + 1
            Next tbl
        Next pageNum
    Next pageRange

    ' 通知用户操作完成
    MsgBox "表格已成功提取并粘贴到新 PowerPoint 演示文稿！", vbInformation
End Sub

