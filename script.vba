
    ' ��������
    ' SrcForm - �������� �����, � ������� ����� �������� ��������
    ' SrcField - �������� ����, � ������� ����� �������� ��������
    ' NewForm - �����, ������� �� ����� ������� ��� ����������
    ' NewField - ����, � ������� �� ����� �������� ��������

    ' ���� ���������
    If (MsgBox("����� �������� �� �������! ��������?", vbYesNo, "��������������") = vbYes) Then
        
        ' ��������� ����� � ��������� ������ �� ���������� (��������� ��������)
        DoCmd.OpenForm "NewForm", acNormal, , , acFormAdd
        If CurrentProject.AllForms("NewForm").IsLoaded Then
            DoCmd.GoToRecord acDataForm, "NewForm", acNewRec
        End If
        
        ' �������� ��������� � ������ �������� ���������������� ����
        Forms![NewForm]![NewField] = NewData
        
        ' ��������� ��������
        DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
        
        ' �������� ��������� �������� � �������� ����� (���� ������)
        Forms![������]![SrcField].Undo
        
        ' ��������� �������� ������
        Forms![������]![SrcField].Requery
        
    Else
    
        Forms![������]![SrcField].Undo
        
    End If
    
    ' �������� ������
    Response = 0
