VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Dim FileOpen
    Dim X As Integer
    Dim Shengfen As String
    Application.ScreenUpdating = False
    ChDir (ThisWorkbook.Path)
    FileOpen = Application.GetOpenFilename(FileFilter:="Microsoft Excel文件(*.xls;*.xlsx),*.xls;*.xlsx", MultiSelect:=True, Title:="请选择需要合并的工作簿")
    X = 1
    If TypeName(FileOpen) = "Boolean" Then
            MsgBox "未选择任何文件, 退出."
            Exit Sub
        Else
            MsgBox "选择了" & UBound(FileOpen) & "个文件"
    End If

    While X <= UBound(FileOpen)
        Workbooks.Open Filename:=FileOpen(X)
        '提取省份名
        m = InStrRev(FileOpen(X), "-") + 1
        n = InStr(FileOpen(X), "省")
        If (n <> 0) Then
                Shengfen = Mid(FileOpen(X), m, n - m)
            Else
                Shengfen = Mid(FileOpen(X), m, InStrRev(FileOpen(X), ".") - m)
        End If
        
        '更改设备量信息表的省份
        Sheets("设备量信息表").Range("B4") = Shengfen
        '粘贴该省设备量信息表至汇总表
        Sheets("设备量信息表").Range("A4", Sheets("设备量信息表").Cells.SpecialCells(xlCellTypeLastCell)).Copy
        ThisWorkbook.Sheets("设备量信息表").Cells(ThisWorkbook.Sheets("设备量信息表").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
        
        '查找该省故障处理事件末尾，如果空表就加个标记
        DataRowEnd = Sheets("故障处理事件信息表").Range("E65536").End(xlUp).row + 1
        Sheets("故障处理事件信息表").Range("G6") = Shengfen
        '粘贴该省故障处理事件信息表至汇总表
        Sheets("故障处理事件信息表").Range("A6", Sheets("故障处理事件信息表").Cells(DataRowEnd, 20)).Copy
        ThisWorkbook.Sheets("故障处理事件信息表").Cells(ThisWorkbook.Sheets("故障处理事件信息表").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
        
        '查找该省重大故障事件末尾，如果空表就加个标记
        DataRowEnd = Sheets("重大故障事件信息表").Range("E65536").End(xlUp).row + 1
        Sheets("重大故障事件信息表").Range("G6") = Shengfen
        '粘贴该省重大故障事件信息表至汇总表
        Sheets("重大故障事件信息表").Range("A6", Sheets("重大故障事件信息表").Cells.Cells(DataRowEnd, 20)).Copy
        ThisWorkbook.Sheets("重大故障事件信息表").Cells(ThisWorkbook.Sheets("重大故障事件信息表").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
         
        '查找该省问题事件末尾，如果空表就加个标记
        DataRowEnd = Sheets("设备问题事件信息表").Range("B65536").End(xlUp).row + 1
        Sheets("设备问题事件信息表").Range("D6") = Shengfen
        '粘贴该省设备问题事件信息表至汇总表
        Sheets("设备问题事件信息表").Range("A6", Sheets("设备问题事件信息表").Cells(DataRowEnd, 20)).Copy
        ThisWorkbook.Sheets("设备问题事件信息表").Cells(ThisWorkbook.Sheets("设备问题事件信息表").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
        
        '查找该省问题事件末尾，如果空表就加个标记
        DataRowEnd = Sheets("割接事件信息表 ").Range("E65536").End(xlUp).row + 1
        Sheets("割接事件信息表 ").Range("G6") = Shengfen
        '粘贴该省设备量信息表至汇总表
        Sheets("割接事件信息表 ").Range("A6", Sheets("割接事件信息表 ").Cells(DataRowEnd, 20)).Copy
        ThisWorkbook.Sheets("割接事件信息表 ").Cells(ThisWorkbook.Sheets("割接事件信息表 ").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
                
        '更改软硬件入网质量的省份
        Sheets("软硬件入网质量").Range("B4") = Shengfen
        '粘贴该省软硬件入网质量至汇总表
        Sheets("软硬件入网质量").Range("A4", Sheets("软硬件入网质量").Cells.SpecialCells(xlCellTypeLastCell)).Copy
        ThisWorkbook.Sheets("软硬件入网质量").Cells(ThisWorkbook.Sheets("软硬件入网质量").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
        
        '更改服务配合度的省份
        Sheets("服务配合度").Range("B4") = Shengfen
        '粘贴该省服务配合度至汇总表
        Sheets("服务配合度").Range("A4", Sheets("服务配合度").Cells.SpecialCells(xlCellTypeLastCell)).Copy
        ThisWorkbook.Sheets("服务配合度").Cells(ThisWorkbook.Sheets("服务配合度").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
        
        '更改服务响应速度的省份
        Sheets("服务响应速度").Range("B4") = Shengfen
        '粘贴该省服务响应速度至汇总表
        Sheets("服务响应速度").Range("A4", Sheets("服务响应速度").Cells.SpecialCells(xlCellTypeLastCell)).Copy
        ThisWorkbook.Sheets("服务响应速度").Cells(ThisWorkbook.Sheets("服务响应速度").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
        
        '更改服务专业性的省份
        Sheets("服务专业性").Range("B4") = Shengfen
        '粘贴该省服务专业性至汇总表
        Sheets("服务专业性").Range("A4", Sheets("服务专业性").Cells.SpecialCells(xlCellTypeLastCell)).Copy
        ThisWorkbook.Sheets("服务专业性").Cells(ThisWorkbook.Sheets("服务专业性").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
        

        '关闭该省后评估表格
        Application.CutCopyMode = False
        ActiveWorkbook.Close SaveChanges:=False
        X = X + 1
    Wend
ExitHandler:
    Application.ScreenUpdating = True
    Exit Sub

errhadler:
    MsgBox Err.Description
End Sub

Private Sub CommandButton2_Click()
    Dim X, Y, Z, i, num, row, col, bShengRow, ShengChangNum, ChangShangRow, LeiXingCol, DataRowEnd, DataColEnd, XieFlag, JgSheng As Integer
    Dim Sheng
    Sheng = Array("安徽", "北京", "福建", "甘肃", "广东", "广西", "贵州", "海南", "河北", "河南", "黑龙江", "湖北", "湖南", "吉林", "江苏", "江西", "辽宁", "内蒙古", "宁夏", "青海", "山东", "山西", "陕西", "上海", "四川", "天津", "西藏", "新疆", "云南", "浙江", "重庆")
    Dim QueSheng(31), LeiXing, ChangShang, bSheng, Msg As String

    XieFlag = 28
    X = ThisWorkbook.Sheets.Count
    
    Do Until X = 1
        Set Sheet = ThisWorkbook.Sheets(X)
        
        '检查各表中是否包含全部省数据
        Y = UBound(Sheng)
        Z = 0
        Do Until Y < 0
            If Sheet.Range("A1:H65536").Find(Sheng(Y)) Is Nothing Then
                QueSheng(Z) = Sheng(Y)
                If Z = 0 Then
                        Msg = QueSheng(Z)
                    Else
                        Msg = Msg & "、" & QueSheng(Z)
                End If
                Z = Z + 1
            End If
            Y = Y - 1
        Loop
        '在第一个工作表B16行开始写入完成信息
        If Z > 0 Then
            Msg = Sheet.Name & "缺少" & Msg
            ThisWorkbook.Sheets(1).Cells(X + 15, 2) = Msg
        Else
            Msg = Sheet.Name & "OK"
            ThisWorkbook.Sheets(1).Cells(X + 15, 2) = Msg
        
        End If
                
        If X = 10 Or X = 9 Or X = 8 Or X = 7 Or X = 2 Then
        '根据T1开始的表格，将主观表格的二维表格转化过去，并检查数据有效性
            ThisWorkbook.Sheets(1).Cells(XieFlag, 1) = Sheet.Name
            DataColEnd = Sheet.Range("CC2").End(xlToLeft).Column
            bShengRow = 3
            For col = 21 To DataColEnd
                If Sheet.Cells(1, col) <> "" Then
                    LeiXing = Sheet.Cells(1, col)
                    num = Sheet.Cells(1, col).MergeArea.Columns.Count
                End If
                ChangShang = Sheet.Cells(2, col)
                num = num - 1
                If (num < 0) Then
                    Msg = ChangShang & "厂商超过" & LeiXing & "类型范围！"
                    ThisWorkbook.Sheets(1).Cells(XieFlag, 2) = Msg
                    XieFlag = XieFlag + 1
                    Exit Sub
                End If
                LeiXingCol = -1
                JgSheng = 0
                For row = 3 To 33
                    bSheng = Sheet.Cells(row, 20)
                    '大表中找省份
                    Set a = Sheet.Range("B3:B65536").Find(bSheng)
                    If Not a Is Nothing Then
                        bShengRow = a.row
                    Else
                        Msg = "找不到" & bSheng
                        ThisWorkbook.Sheets(1).Cells(XieFlag, 2) = Msg
                        XieFlag = XieFlag + 1
                    End If
                    '计算该省有多少厂家
                    ShengChangNum = 1
                    If row <> 33 Then
                        Do Until Sheet.Cells(bShengRow + ShengChangNum, 2) <> ""
                            ShengChangNum = ShengChangNum + 1
                        Loop
                    Else
                        ShengChangNum = Sheet.Range("C65536").End(xlUp).row + 1
                    End If
                    '查找指定的厂家
                    Set a = Sheet.Range(Cells(bShengRow, 3).Address, Cells(bShengRow + ShengChangNum - 1, 3).Address).Find(ChangShang)
                    If Not a Is Nothing Then
                        ChangShangRow = a.row
                    Else
                        Msg = bSheng & "找不到" & ChangShang
                        ThisWorkbook.Sheets(1).Cells(XieFlag, 2) = Msg
                        XieFlag = XieFlag + 1
                    End If
                    '找3排的类型列
                    If LeiXingCol = -1 Then
                        Set a = Sheet.Range("C3", "S3").Find(LeiXing)
                        If Not a Is Nothing Then
                            LeiXingCol = a.Column
                        Else
                            Msg = "找不到" & LeiXing
                            ThisWorkbook.Sheets(1).Cells(XieFlag, 2) = Msg
                            XieFlag = XieFlag + 1
                        End If
                    End If
                    '将大表数据写入小表
                    b = Sheet.Cells(ChangShangRow, LeiXingCol)
                    If b <> "" Then
                        If IsNumeric(b) Then
                            If (b < 0 Or ((X = 10 Or X = 9) And b > 6) Or (X = 8 And b > 8) Or (X = 7 And b > 10)) Then
                                Msg = bSheng & "的" & LeiXing & "厂商" & ChangShang & "是" & b & ",超过范围"
                                ThisWorkbook.Sheets(1).Cells(XieFlag, 2) = Msg
                                XieFlag = XieFlag + 1
                                Exit Sub
                            Else
                                If b <> Int(b) Then
                                    Msg = bSheng & "的" & LeiXing & "厂商" & ChangShang & "是" & b & ",不是整数"
                                    ThisWorkbook.Sheets(1).Cells(XieFlag, 2) = Msg
                                    XieFlag = XieFlag + 1
                               End If
                                If b <> 0 And b <> "" Then
                                    Sheet.Cells(row, col) = Sheet.Cells(ChangShangRow, LeiXingCol)
                                    JgSheng = JgSheng + 1
                                End If
                            End If
                        Else
                            Msg = bSheng & "的" & LeiXing & "厂商" & ChangShang & "是" & b & ",不是数字"
                            ThisWorkbook.Sheets(1).Cells(XieFlag, 2) = Msg
                            XieFlag = XieFlag + 1
                        End If
                    End If
                Next
                Sheet.Cells(35, col) = JgSheng
            Next

        '从B30开始写入错误信息
        Else
        '检查客观表格各行是否有示例、空行、无意义的行，并逐个删除
            DataRowEnd = Sheet.Range("E65536").End(xlUp).row + 1
            For i = DataRowEnd To 3 Step -1
                If X = 3 Or X = 4 Or X = 6 Then
                    If Sheet.Cells(i, 1) = "示例" Or ((Sheet.Cells(i, 5) = "" Or Sheet.Cells(i, 5) = "无") And Sheet.Cells(i, 9) = "") Then Sheet.Cells(i, 1).EntireRow.Delete
                ElseIf X = 5 Then
                    If Sheet.Cells(i, 1) = "示例" Or ((Sheet.Cells(i, 2) = "" Or Sheet.Cells(i, 2) = "无") And Sheet.Cells(i, 6) = "") Then Sheet.Cells(i, 1).EntireRow.Delete
                End If
            Next
        End If
        X = X - 1
    Loop
End Sub

