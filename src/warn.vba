' Создайте модуль Warn в своей БД, добавьте туда код до черты

Option Compare Database

Public Function DoNotInList(SrcFieldName As String, _
                            NewFieldName As String, _
                            SrcFormName As String, _
                            NewFormName As String, _
                            NewData As String _
                           ) As Integer
    
    ' Выводим сообщение об отсутвии значения в списке
    If (MsgBox("Такое значение не найдено! Добавить?", vbYesNo, "Предупреждение") = vbYes) Then
        
        ' открываем форму с указанным именем на добавление
        DoCmd.OpenForm NewFormName, acNormal, , , acFormAdd
        If CurrentProject.AllForms(NewFormName).IsLoaded Then
            DoCmd.GoToRecord acDataForm, NewFormName, acNewRec
        End If
        
        ' обращаемся (динамически) к объектам на форме
        Dim SrcField As ComboBox
        Set SrcField = Application.Forms(SrcFormName).Controls(SrcFieldName)
        
        Dim NewField As TextBox
        Set NewField = Application.Forms(NewFormName).Controls(NewFieldName)
        
        ' присвоим введенное в список значение соответствующему полю
        NewField.Value = NewData
        
        ' сохраняем значение
        DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
        
        ' отменяем последнее действие в исходной форме (ввод текста)
        SrcField.Undo
        
        ' обновляем значения списка
        SrcField.Requery
        
    Else
    
        SrcField.Undo
        
    End If
    
    DoNotInList = 0 ' Response
    
End Function

'-----------------------------------------------------------
' Применение модуля:
'-----------------------------------------------------------
' Обработчик события "Отсутсвие в списке" с исп. модуля
Private Sub Учитель_ID_NotInList(NewData As String, Response As Integer)    
    ' Обрабатываем ошибку
    ' Имя исходного поля, имя заполняемого поля, имя исходной формы, имя открываемой формы
    Response = Warn.DoNotInList("Учитель_ID", "ФИО учителя", "Классы", "Учителя", NewData)
End Sub
'-----------------------------------------------------------
' Обработчик события "Отсутсвие в списке" без исп. модуля
Private Sub День_недели_NotInList(NewData As String, Response As Integer)
    MsgBox ("Данное значение необходимо выбрать строго из предложенных!")
    Response = 0
End Sub
'-----------------------------------------------------------