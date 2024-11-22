Private Sub ComboBox1_Change()

End Sub

Private Sub UserForm_Initialize()
    Dim NamesMatrix As Variant
    NamesMatrix = Array("Наталья", "Татьяна Ивановна", "Айрат", "Таня Собянина", "Надежда Николаевна", "Юля", "Марина Попова", "Марина Алексеевна", "Оля", "ОТиЗ Префаб")
    
    ' Заполняем ComboBox именами
    Dim name As Variant
    For Each name In NamesMatrix
        Me.ComboBox1.AddItem name
    Next name
    
End Sub
Private Sub CommandButton1_Click()
    Dim NamesMatrix As Variant, ChatIDsMatrix As Variant
    NamesMatrix = Array("Наталья", "Татьяна Ивановна", "Айрат", "Таня Собянина", "Надежда Николаевна", "Юля", "Марина Попова", "Марина Алексеевна", "Оля", "ОТиЗ Префаб")
    ChatIDsMatrix = Array("*******", "*******", "*******", "*******", "*******", "*******", "*******", "*******", "*******", "-*******")
    
    'Сохраняем выбранное имя в переменную
    Dim selectedName As String
    If Me.ComboBox1.ListIndex < 0 Then
        MsgBox "Выберите имя!"
        Exit Sub
    End If
        
    selectedName = Me.ComboBox1.Value
    
    Dim CHAT_ID As String
    Dim name As Variant
    Dim i As Long
    
    i = 0
    For Each name In NamesMatrix
        If name = selectedName Then
            CHAT_ID = ChatIDsMatrix(i)
            Exit For
        End If
        i = i + 1
    Next name
             
    
    'Передаем переменную в функцию из Module1
    'MsgBox selectedName
    Module1.Сохранить_и_Отправить_Скрин_Телеграм CHAT_ID, selectedName
    
    'Закрываем форму
    Unload Me
End Sub
Private Sub CommandButton3_Click()
    Сохранить_Скрин
    
    'Закрываем форму
    Unload Me
End Sub
Private Sub CommandButton2_Click()
    ' Закрываем форму
    Unload Me
End Sub
