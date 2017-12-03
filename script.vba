
    ' «аменить
    ' SrcForm - исходна€ форма, в которую ввели неверное значение
    ' SrcField - исходное поле, в которое ввели неверное значение
    ' NewForm - форма, которую мы хотим открыть дл€ добавлени€
    ' NewField - поле, в которое мы хотим добавить значение

    ' окно сообщени€
    If (MsgBox("“акое значение не найдено! ƒобавить?", vbYesNo, "ѕредупреждение") = vbYes) Then
        
        ' открываем форму с указанным именем на добавление (последний аргумент)
        DoCmd.OpenForm "NewForm", acNormal, , , acFormAdd
        If CurrentProject.AllForms("NewForm").IsLoaded Then
            DoCmd.GoToRecord acDataForm, "NewForm", acNewRec
        End If
        
        ' присвоим введенное в список значение соответствующему полю
        Forms![NewForm]![NewField] = NewData
        
        ' сохран€ем значение
        DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
        
        ' отмен€ем последнее действие в исходной форме (ввод текста)
        Forms![ лассы]![SrcField].Undo
        
        ' обновл€ем значени€ списка
        Forms![ лассы]![SrcField].Requery
        
    Else
    
        Forms![ лассы]![SrcField].Undo
        
    End If
    
    ' обнул€ем ошибку
    Response = 0
