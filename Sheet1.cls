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
    FileOpen = Application.GetOpenFilename(FileFilter:="Microsoft Excel�ļ�(*.xls;*.xlsx),*.xls;*.xlsx", MultiSelect:=True, Title:="��ѡ����Ҫ�ϲ��Ĺ�����")
    X = 1
    If TypeName(FileOpen) = "Boolean" Then
            MsgBox "δѡ���κ��ļ�, �˳�."
            Exit Sub
        Else
            MsgBox "ѡ����" & UBound(FileOpen) & "���ļ�"
    End If

    While X <= UBound(FileOpen)
        Workbooks.Open Filename:=FileOpen(X)
        '��ȡʡ����
        m = InStrRev(FileOpen(X), "-") + 1
        n = InStr(FileOpen(X), "ʡ")
        If (n <> 0) Then
                Shengfen = Mid(FileOpen(X), m, n - m)
            Else
                Shengfen = Mid(FileOpen(X), m, InStrRev(FileOpen(X), ".") - m)
        End If
        
        '�����豸����Ϣ����ʡ��
        Sheets("�豸����Ϣ��").Range("B4") = Shengfen
        'ճ����ʡ�豸����Ϣ�������ܱ�
        Sheets("�豸����Ϣ��").Range("A4", Sheets("�豸����Ϣ��").Cells.SpecialCells(xlCellTypeLastCell)).Copy
        ThisWorkbook.Sheets("�豸����Ϣ��").Cells(ThisWorkbook.Sheets("�豸����Ϣ��").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
        
        '���Ҹ�ʡ���ϴ����¼�ĩβ������ձ��ͼӸ����
        DataRowEnd = Sheets("���ϴ����¼���Ϣ��").Range("E65536").End(xlUp).row + 1
        Sheets("���ϴ����¼���Ϣ��").Range("G6") = Shengfen
        'ճ����ʡ���ϴ����¼���Ϣ�������ܱ�
        Sheets("���ϴ����¼���Ϣ��").Range("A6", Sheets("���ϴ����¼���Ϣ��").Cells(DataRowEnd, 20)).Copy
        ThisWorkbook.Sheets("���ϴ����¼���Ϣ��").Cells(ThisWorkbook.Sheets("���ϴ����¼���Ϣ��").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
        
        '���Ҹ�ʡ�ش�����¼�ĩβ������ձ��ͼӸ����
        DataRowEnd = Sheets("�ش�����¼���Ϣ��").Range("E65536").End(xlUp).row + 1
        Sheets("�ش�����¼���Ϣ��").Range("G6") = Shengfen
        'ճ����ʡ�ش�����¼���Ϣ�������ܱ�
        Sheets("�ش�����¼���Ϣ��").Range("A6", Sheets("�ش�����¼���Ϣ��").Cells.Cells(DataRowEnd, 20)).Copy
        ThisWorkbook.Sheets("�ش�����¼���Ϣ��").Cells(ThisWorkbook.Sheets("�ش�����¼���Ϣ��").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
         
        '���Ҹ�ʡ�����¼�ĩβ������ձ��ͼӸ����
        DataRowEnd = Sheets("�豸�����¼���Ϣ��").Range("B65536").End(xlUp).row + 1
        Sheets("�豸�����¼���Ϣ��").Range("D6") = Shengfen
        'ճ����ʡ�豸�����¼���Ϣ�������ܱ�
        Sheets("�豸�����¼���Ϣ��").Range("A6", Sheets("�豸�����¼���Ϣ��").Cells(DataRowEnd, 20)).Copy
        ThisWorkbook.Sheets("�豸�����¼���Ϣ��").Cells(ThisWorkbook.Sheets("�豸�����¼���Ϣ��").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
        
        '���Ҹ�ʡ�����¼�ĩβ������ձ��ͼӸ����
        DataRowEnd = Sheets("����¼���Ϣ�� ").Range("E65536").End(xlUp).row + 1
        Sheets("����¼���Ϣ�� ").Range("G6") = Shengfen
        'ճ����ʡ�豸����Ϣ�������ܱ�
        Sheets("����¼���Ϣ�� ").Range("A6", Sheets("����¼���Ϣ�� ").Cells(DataRowEnd, 20)).Copy
        ThisWorkbook.Sheets("����¼���Ϣ�� ").Cells(ThisWorkbook.Sheets("����¼���Ϣ�� ").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
                
        '������Ӳ������������ʡ��
        Sheets("��Ӳ����������").Range("B4") = Shengfen
        'ճ����ʡ��Ӳ���������������ܱ�
        Sheets("��Ӳ����������").Range("A4", Sheets("��Ӳ����������").Cells.SpecialCells(xlCellTypeLastCell)).Copy
        ThisWorkbook.Sheets("��Ӳ����������").Cells(ThisWorkbook.Sheets("��Ӳ����������").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
        
        '���ķ�����϶ȵ�ʡ��
        Sheets("������϶�").Range("B4") = Shengfen
        'ճ����ʡ������϶������ܱ�
        Sheets("������϶�").Range("A4", Sheets("������϶�").Cells.SpecialCells(xlCellTypeLastCell)).Copy
        ThisWorkbook.Sheets("������϶�").Cells(ThisWorkbook.Sheets("������϶�").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
        
        '���ķ�����Ӧ�ٶȵ�ʡ��
        Sheets("������Ӧ�ٶ�").Range("B4") = Shengfen
        'ճ����ʡ������Ӧ�ٶ������ܱ�
        Sheets("������Ӧ�ٶ�").Range("A4", Sheets("������Ӧ�ٶ�").Cells.SpecialCells(xlCellTypeLastCell)).Copy
        ThisWorkbook.Sheets("������Ӧ�ٶ�").Cells(ThisWorkbook.Sheets("������Ӧ�ٶ�").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
        
        '���ķ���רҵ�Ե�ʡ��
        Sheets("����רҵ��").Range("B4") = Shengfen
        'ճ����ʡ����רҵ�������ܱ�
        Sheets("����רҵ��").Range("A4", Sheets("����רҵ��").Cells.SpecialCells(xlCellTypeLastCell)).Copy
        ThisWorkbook.Sheets("����רҵ��").Cells(ThisWorkbook.Sheets("����רҵ��").Range("A65536").End(xlUp).row + 1, 1).PasteSpecial Paste:=xlPasteAll
        

        '�رո�ʡ����������
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
    Sheng = Array("����", "����", "����", "����", "�㶫", "����", "����", "����", "�ӱ�", "����", "������", "����", "����", "����", "����", "����", "����", "���ɹ�", "����", "�ຣ", "ɽ��", "ɽ��", "����", "�Ϻ�", "�Ĵ�", "���", "����", "�½�", "����", "�㽭", "����")
    Dim QueSheng(31), LeiXing, ChangShang, bSheng, Msg As String

    XieFlag = 28
    X = ThisWorkbook.Sheets.Count
    
    Do Until X = 1
        Set Sheet = ThisWorkbook.Sheets(X)
        
        '���������Ƿ����ȫ��ʡ����
        Y = UBound(Sheng)
        Z = 0
        Do Until Y < 0
            If Sheet.Range("A1:H65536").Find(Sheng(Y)) Is Nothing Then
                QueSheng(Z) = Sheng(Y)
                If Z = 0 Then
                        Msg = QueSheng(Z)
                    Else
                        Msg = Msg & "��" & QueSheng(Z)
                End If
                Z = Z + 1
            End If
            Y = Y - 1
        Loop
        '�ڵ�һ��������B16�п�ʼд�������Ϣ
        If Z > 0 Then
            Msg = Sheet.Name & "ȱ��" & Msg
            ThisWorkbook.Sheets(1).Cells(X + 15, 2) = Msg
        Else
            Msg = Sheet.Name & "OK"
            ThisWorkbook.Sheets(1).Cells(X + 15, 2) = Msg
        
        End If
                
        If X = 10 Or X = 9 Or X = 8 Or X = 7 Or X = 2 Then
        '����T1��ʼ�ı��񣬽����۱���Ķ�ά����ת����ȥ�������������Ч��
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
                    Msg = ChangShang & "���̳���" & LeiXing & "���ͷ�Χ��"
                    ThisWorkbook.Sheets(1).Cells(XieFlag, 2) = Msg
                    XieFlag = XieFlag + 1
                    Exit Sub
                End If
                LeiXingCol = -1
                JgSheng = 0
                For row = 3 To 33
                    bSheng = Sheet.Cells(row, 20)
                    '�������ʡ��
                    Set a = Sheet.Range("B3:B65536").Find(bSheng)
                    If Not a Is Nothing Then
                        bShengRow = a.row
                    Else
                        Msg = "�Ҳ���" & bSheng
                        ThisWorkbook.Sheets(1).Cells(XieFlag, 2) = Msg
                        XieFlag = XieFlag + 1
                    End If
                    '�����ʡ�ж��ٳ���
                    ShengChangNum = 1
                    If row <> 33 Then
                        Do Until Sheet.Cells(bShengRow + ShengChangNum, 2) <> ""
                            ShengChangNum = ShengChangNum + 1
                        Loop
                    Else
                        ShengChangNum = Sheet.Range("C65536").End(xlUp).row + 1
                    End If
                    '����ָ���ĳ���
                    Set a = Sheet.Range(Cells(bShengRow, 3).Address, Cells(bShengRow + ShengChangNum - 1, 3).Address).Find(ChangShang)
                    If Not a Is Nothing Then
                        ChangShangRow = a.row
                    Else
                        Msg = bSheng & "�Ҳ���" & ChangShang
                        ThisWorkbook.Sheets(1).Cells(XieFlag, 2) = Msg
                        XieFlag = XieFlag + 1
                    End If
                    '��3�ŵ�������
                    If LeiXingCol = -1 Then
                        Set a = Sheet.Range("C3", "S3").Find(LeiXing)
                        If Not a Is Nothing Then
                            LeiXingCol = a.Column
                        Else
                            Msg = "�Ҳ���" & LeiXing
                            ThisWorkbook.Sheets(1).Cells(XieFlag, 2) = Msg
                            XieFlag = XieFlag + 1
                        End If
                    End If
                    '���������д��С��
                    b = Sheet.Cells(ChangShangRow, LeiXingCol)
                    If b <> "" Then
                        If IsNumeric(b) Then
                            If (b < 0 Or ((X = 10 Or X = 9) And b > 6) Or (X = 8 And b > 8) Or (X = 7 And b > 10)) Then
                                Msg = bSheng & "��" & LeiXing & "����" & ChangShang & "��" & b & ",������Χ"
                                ThisWorkbook.Sheets(1).Cells(XieFlag, 2) = Msg
                                XieFlag = XieFlag + 1
                                Exit Sub
                            Else
                                If b <> Int(b) Then
                                    Msg = bSheng & "��" & LeiXing & "����" & ChangShang & "��" & b & ",��������"
                                    ThisWorkbook.Sheets(1).Cells(XieFlag, 2) = Msg
                                    XieFlag = XieFlag + 1
                               End If
                                If b <> 0 And b <> "" Then
                                    Sheet.Cells(row, col) = Sheet.Cells(ChangShangRow, LeiXingCol)
                                    JgSheng = JgSheng + 1
                                End If
                            End If
                        Else
                            Msg = bSheng & "��" & LeiXing & "����" & ChangShang & "��" & b & ",��������"
                            ThisWorkbook.Sheets(1).Cells(XieFlag, 2) = Msg
                            XieFlag = XieFlag + 1
                        End If
                    End If
                Next
                Sheet.Cells(35, col) = JgSheng
            Next

        '��B30��ʼд�������Ϣ
        Else
        '���͹۱�������Ƿ���ʾ�������С���������У������ɾ��
            DataRowEnd = Sheet.Range("E65536").End(xlUp).row + 1
            For i = DataRowEnd To 3 Step -1
                If X = 3 Or X = 4 Or X = 6 Then
                    If Sheet.Cells(i, 1) = "ʾ��" Or ((Sheet.Cells(i, 5) = "" Or Sheet.Cells(i, 5) = "��") And Sheet.Cells(i, 9) = "") Then Sheet.Cells(i, 1).EntireRow.Delete
                ElseIf X = 5 Then
                    If Sheet.Cells(i, 1) = "ʾ��" Or ((Sheet.Cells(i, 2) = "" Or Sheet.Cells(i, 2) = "��") And Sheet.Cells(i, 6) = "") Then Sheet.Cells(i, 1).EntireRow.Delete
                End If
            Next
        End If
        X = X - 1
    Loop
End Sub
