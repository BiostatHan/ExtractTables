Attribute VB_Name = "ģ��1"
' ========================================
' ������: ExtractTablesToPPT
' ����: ��ȡ Word �ĵ���ָ��ҳ��ı�񣬲�ճ���� PowerPoint �еĻõ�Ƭ
' ����: �⺲
' �汾: 1.0
' ��������: 2024/12/4
' ����������: 2024/12/4
' ������ʷ:
'   v1.0 - ��ʼ�汾��ʵ�ֻ�������
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
    
    ' ��ʾ�û�����ҳ�淶Χ
    specifiedPages = InputBox("��������Ҫ��ȡ����ҳ�淶Χ�����磺1,2,3-5����", "��ȡ���PPT")
    If specifiedPages = "" Then
        MsgBox "δ�����κ�ҳ�淶Χ��", vbExclamation
        Exit Sub
    End If
    
    ' ���õ�ǰ�ĵ�ΪԴ�ĵ�
    Set sourceDoc = ActiveDocument

    ' ���� PowerPoint Ӧ�ó��������ʾ�ĸ�
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application") ' ���Ի�ȡ PowerPoint Ӧ��
    If pptApp Is Nothing Then
        Set pptApp = CreateObject("PowerPoint.Application") ' ���δ���У��򴴽���ʵ��
    End If
    On Error GoTo 0
    pptApp.Visible = True
    Set pptPresentation = pptApp.Presentations.Add

    ' �ָ������ҳ�淶Χ
    pagesArray = Split(specifiedPages, ",")
    i = 1
    
    ' ����ÿ��ָ����ҳ���Χ
    For Each pageRange In pagesArray
        If InStr(pageRange, "-") > 0 Then
            ' ����Ƿ�Χ����3-5������ȡ��ʼ�ͽ���ҳ��
            startPage = CLng(Split(pageRange, "-")(0))
            endPage = CLng(Split(pageRange, "-")(1))
        Else
            ' ����ǵ���ҳ��
            startPage = CLng(pageRange)
            endPage = startPage
        End If
        
        ' ���ҳ�淶Χ�Ƿ���Ч
        If startPage <= 0 Or endPage < startPage Then
            MsgBox "ҳ�淶Χ��Ч��" & pageRange, vbExclamation
            Exit Sub
        End If
        
        ' ������Χ�ڵ�ÿһҳ
        For pageNum = startPage To endPage
            ' ��ȡָ��ҳ��ķ�Χ
            Set rng = sourceDoc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=pageNum)
            rng.End = sourceDoc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=pageNum + 1).Start
            rng.End = rng.End - 1 ' ������Χ��ҳ�����
            
            ' ����ҳ���ڵı��
            For Each tbl In rng.Tables
                ' ���Ʊ��
                tbl.Range.Copy
                
                ' �����»õ�Ƭ��ճ�����
                Set pptSlide = pptPresentation.Slides.Add(Index:=pptPresentation.Slides.Count + 1, Layout:=12) ' �հײ���
                pptSlide.Shapes.Paste
                i = i + 1
            Next tbl
        Next pageNum
    Next pageRange

    ' ֪ͨ�û��������
    MsgBox "����ѳɹ���ȡ��ճ������ PowerPoint ��ʾ�ĸ壡", vbInformation
End Sub

