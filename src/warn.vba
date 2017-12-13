' �������� ������ Warn � ����� ��, �������� ���� ��� �� �����

Option Compare Database

Public Function DoNotInList(SrcFieldName As String, _
                            NewFieldName As String, _
                            SrcFormName As String, _
                            NewFormName As String, _
                            NewData As String _
                           ) As Integer
    
    ' ������� ��������� �� �������� �������� � ������
    If (MsgBox("����� �������� �� �������! ��������?", vbYesNo, "��������������") = vbYes) Then
        
        ' ��������� ����� � ��������� ������ �� ����������
        DoCmd.OpenForm NewFormName, acNormal, , , acFormAdd
        If CurrentProject.AllForms(NewFormName).IsLoaded Then
            DoCmd.GoToRecord acDataForm, NewFormName, acNewRec
        End If
        
        ' ���������� (�����������) � �������� �� �����
        Dim SrcField As ComboBox
        Set SrcField = Application.Forms(SrcFormName).Controls(SrcFieldName)
        
        Dim NewField As TextBox
        Set NewField = Application.Forms(NewFormName).Controls(NewFieldName)
        
        ' �������� ��������� � ������ �������� ���������������� ����
        NewField.Value = NewData
        
        ' ��������� ��������
        DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
        
        ' �������� ��������� �������� � �������� ����� (���� ������)
        SrcField.Undo
        
        ' ��������� �������� ������
        SrcField.Requery
        
    Else
    
        SrcField.Undo
        
    End If
    
    DoNotInList = 0 ' Response
    
End Function

'-----------------------------------------------------------
' ���������� ������:
'-----------------------------------------------------------
' ���������� ������� "��������� � ������" � ���. ������
Private Sub �������_ID_NotInList(NewData As String, Response As Integer)    
    ' ������������ ������
    ' ��� ��������� ����, ��� ������������ ����, ��� �������� �����, ��� ����������� �����
    Response = Warn.DoNotInList("�������_ID", "��� �������", "������", "�������", NewData)
End Sub
'-----------------------------------------------------------
' ���������� ������� "��������� � ������" ��� ���. ������
Private Sub ����_������_NotInList(NewData As String, Response As Integer)
    MsgBox ("������ �������� ���������� ������� ������ �� ������������!")
    Response = 0
End Sub
'-----------------------------------------------------------