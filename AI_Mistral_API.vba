Option Explicit

Sub Testing_Mistral_AI()

    Dim response As VbMsgBoxResult
    
    ' Вызываем диалоговое окно с вопросом и вариантами ответа "Да" и "Нет"
    response = MsgBox("Внимание! Отправляйте только ОБЕЗЛИЧЕННЫЕ данные! Вы хотите продолжить?", vbYesNo + vbQuestion, "Подтверждение")
    
    ' Проверяем, какой вариант выбрал пользователь
    If response = vbNo Then
        Exit Sub
    End If

    Dim question As String
    Dim systemRole As String
    Dim agent_id As String
        
    
    ' Проверяем, что выделен диапазон ячеек
    If TypeName(Selection) <> "Range" Then
        MsgBox "Пожалуйста, выделите диапазон ячеек."
        Exit Sub
    End If
    
    Dim cellsCount As Long
    cellsCount = Selection.Count
    
    If cellsCount < 2 Then
        MsgBox ("Выберите минимум 2 ячейки для анализа этих данных, потом снова нажмите кнопку.")
        Exit Sub
    End If
    
    Dim rng As Range
    Set rng = Selection
    
    Dim data
    data = GetSelectedCellsValue(rng)
    
    data = Replace(data, vbLf, " ")
    
    
    
    systemRole = "Есть вот такие данные из таблицы, где столбцы разделены знаком ; а строки разделены знаком | : " & data
    'systemRole = "Есть вот такие данные из таблицы: [[,%"
    
    
    ' Заменяем все вхождения " на ПУСТО
    systemRole = Replace(systemRole, """", "")
    'Debug.Print systemRole
    
    agent_id = "***YOUR_AGENT_ID***"
    
    
    Dim answer As String
    Dim outputString As String
    
    'question = "Сделай вывод по этим данным. Без уточнений."
    question = "Кто из них начальник?"
    
    answer = mistral_AI_API_agents_with_SystemRole(question, agent_id, systemRole)
    'answer = Mistral_AI_API(question)
    
    ' Заменяем все вхождения "\n" на " & vbNewLine & "
    outputString = Replace(answer, "\n", "" & vbNewLine & "")
    'outputString = answer
    'outputString = Replace(answer, "\n", " ")
    
    'MsgBox (outputString)
    
    Debug.Print data
    Debug.Print outputString
End Sub
Private Function GetSelectedCellsValue(rng As Range)
        
    Dim cell As Range
    Dim i As Integer
    Dim j As Integer
    
    Dim start_row As Long
    Dim start_column As Long

    start_row = rng.row
    start_column = rng.Column
    
    Dim rng_rows    As Long
    Dim rng_columns As Long

    rng_rows = rng.Rows.Count
    rng_columns = rng.Columns.Count
    
    Dim json        As String
    Dim vremenString As String

    If rng_rows = 1 And rng_columns = 1 Then
        vremenString = rng.Value
        vremenString = Replace(vremenString, "\", "??? ")
        vremenString = Replace(vremenString, "'", "???? ")
        vremenString = Replace(vremenString, """", "????? ")
        vremenString = Replace(vremenString, "&", "?????? ")

        json = "[[R" & start_row & "C" & start_column & ":" & vremenString & "]]"
    Else

        Dim massiv_rng
        massiv_rng = rng.Value
        json = "["

        For i = 1 To UBound(massiv_rng, 1)
            json = json & "["
            For j = 1 To UBound(massiv_rng, 2)

                vremenString = massiv_rng(i, j)
                vremenString = Replace(vremenString, "\", "??? ")
                vremenString = Replace(vremenString, "'", "???? ")
                vremenString = Replace(vremenString, """", "????? ")
                vremenString = Replace(vremenString, "&", "?????? ")

                json = json & "R" & start_row + i - 1 & "C" & start_column + j - 1 & ":" & vremenString & ";"
            Next j
            json = json & "]|"
        Next i

        json = Replace(json, ";]", "]")
        json = json & "]"
        json = Replace(json, "|]", "]")

    End If
    
    GetSelectedCellsValue = json
End Function
Private Function Mistral_AI_API(question As String) As String
    Dim mistral_key As String
    Dim URL As String
    Dim data As String
    Dim http As Object
    Dim response As String
    Dim responseJSON As Object
    Dim text As String
    
    mistral_key = "***YOUR_MISTRAL_KEY***"
    URL = "https://api.mistral.ai/v1/chat/completions"
    
    ' Создаем JSON payload
    data = "{""messages"": [{""role"": ""user"", ""content"": """ & question & """}], ""model"": ""mistral-large-latest"", ""stream"": false, ""temperature"": 0.7}"
    
    ' Создаем объект для HTTP запроса
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", URL, False
    http.setRequestHeader "Authorization", "Bearer " & mistral_key
    http.setRequestHeader "Content-Type", "application/json"
    http.send data
    
    ' Получаем ответ
    response = http.responseText
    
   
    Mistral_AI_API = ExtractContent(response)
       
    
End Function

Private Function mistral_AI_API_agents_with_SystemRole(question As String, agent_id As String, systemRole As String) As String
    Dim mistral_key As String
    Dim URL As String
    Dim data As String
    Dim http As Object
    Dim response As String
    
    mistral_key = "***YOUR_MISTRAL_KEY***"
    URL = "https://api.mistral.ai/v1/agents/completions"
    
    ' Создаем JSON payload
    data = "{""messages"": [{""role"": ""system"", ""content"": """ & systemRole & """},{""role"": ""user"", ""content"": """ & question & """}], ""stream"": false, ""agent_id"": """ & agent_id & """}"
    'Debug.Print data
    ' Создаем объект для HTTP запроса
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "POST", URL, False
    http.setRequestHeader "Authorization", "Bearer " & mistral_key
    http.setRequestHeader "Content-Type", "application/json"
    http.send data
    
    ' Получаем ответ
    response = http.responseText
    
   
    mistral_AI_API_agents_with_SystemRole = ExtractContent(response)
       
    
End Function


Private Function ExtractContent(inputString As String) As String
    
    Dim contentStart As Long
    Dim contentEnd As Long
    Dim content As String
    
    ' Находим позицию слова "content"
    contentStart = InStr(inputString, """content"":""")
    
    ' Если слово "content" найдено
    If contentStart > 0 Then
        ' Смещаемся на длину строки "content":"
        contentStart = contentStart + Len("""content"":""")
        
        ' Находим позицию закрывающей кавычки
        contentEnd = InStr(contentStart, inputString, """,""")
        
        ' Извлекаем текст между кавычками
        content = Mid(inputString, contentStart, contentEnd - contentStart)
                
    Else
        MsgBox "Слово 'content' не найдено в строке."
    End If
    
    ExtractContent = content
End Function


Sub Testing_Mistral_AI_2()

    Dim response As VbMsgBoxResult
    
    ' Вызываем диалоговое окно с вопросом и вариантами ответа "Да" и "Нет"
    response = MsgBox("Внимание! Отправляйте только ОБЕЗЛИЧЕННЫЕ данные! Вы хотите продолжить?", vbYesNo + vbQuestion, "Подтверждение")
    
    ' Проверяем, какой вариант выбрал пользователь
    If response = vbNo Then
        Exit Sub
    End If

    Dim question As String
    Dim systemRole As String
    Dim agent_id As String
        
    
    ' Проверяем, что выделен диапазон ячеек
    If TypeName(Selection) <> "Range" Then
        MsgBox "Пожалуйста, выделите диапазон ячеек."
        Exit Sub
    End If
    
    Dim cellsCount As Long
    cellsCount = Selection.Count
    
    If cellsCount < 2 Then
        MsgBox ("Выберите минимум 2 ячейки для анализа этих данных, потом снова нажмите кнопку.")
        Exit Sub
    End If
    
    Dim dopQuestion As String
    dopQuestion = InputBox("Введите запрос:")
    
    Debug.Print dopQuestion
    
    Dim rng As Range
    Set rng = Selection
    
    Dim data
    data = GetSelectedCellsValue(rng)
    
    data = Replace(data, vbLf, " ")
    
    Debug.Print data
    
    systemRole = "В таблице, столбцы разделены знаком ; а строки разделены знаком |. Каждый элемент имеет номер ячейки типа R1C1 и само значение ячейки. Вот эта таблица: " & data
    'systemRole = "Есть вот такие данные из таблицы: [[,%"
    
    
    ' Заменяем все вхождения " на ПУСТО
    systemRole = Replace(systemRole, """", "")
    'Debug.Print systemRole
    
    agent_id = "***YOUR_AGENT_ID***"
    
    
    Dim answer As String
    Dim outputString As String
    
    'question = "Сделай вывод по этим данным. Без уточнений."
    question = dopQuestion & ".Напиши ответ в таком же виде R1C1, строки раздели знаком | а столбцы раздели знаком ;. Не расписывай никаких промежуточных расчётов, напиши только точный ответ"
    
    answer = mistral_AI_API_agents_with_SystemRole(question, agent_id, systemRole)
    'answer = Mistral_AI_API(question)
    Debug.Print answer
    
'    Dim splitArray() As String
'    ' Разделяем строку на элементы по разделителю "|"
'    splitArray = Split(answer, "|")
    
    ' Выводим элементы массива в ячейки ниже
'    Dim i As Long
'    For i = LBound(splitArray) To UBound(splitArray)
'        Debug.Print splitArray(i)
'    Next i
    
    
    ' Заменяем все вхождения "\n" на " & vbNewLine & "
    'outputString = Replace(answer, "\n", "" & vbNewLine & "")
    'outputString = answer
    'outputString = Replace(answer, "\n", " ")
    
    'MsgBox (outputString)
    
    'Debug.Print outputString
    
    Dim regex As Object
    Dim matches As Object
    Dim match As Variant
    Dim cellAddress As String
    Dim cellValue As String
    Dim row As Integer
    Dim col As Integer
       
    
    ' Создаем объект регулярного выражения
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "\[R(\d+)C(\d+):([^]]+)\]"
    
    ' Ищем все совпадения в строке DATA
    Set matches = regex.Execute(answer)
        
    
    ' Проходим по каждому совпадению и заполняем ячейки
    For Each match In matches
        row = CInt(match.SubMatches(0))
        col = CInt(match.SubMatches(1))
        cellValue = match.SubMatches(2)
        
        ' Заполняем ячейку
        Cells(row, col).Value = cellValue
    Next match
    
    MsgBox answer
    
End Sub

Sub FillCellsFromData()
    Dim data As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Variant
    Dim cellAddress As String
    Dim cellValue As String
    Dim row As Integer
    Dim col As Integer
    
    ' Пример строки DATA
    data = "[[R2C1:Январь]|[R3C1:Февраль]|[R4C1:Март]|[R5C1:Апрель]|[R6C1:Май]|[R7C1:Июнь]|[R8C1:Июль]|[R9C1:Август]|[R10C1:Сентябрь]|[R11C1:Октябрь]|[R12C1:Ноябрь]|[R13C1:Декабрь]]"
    
    ' Создаем объект регулярного выражения
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "\[R(\d+)C(\d+):([^]]+)\]"
    
    ' Ищем все совпадения в строке DATA
    Set matches = regex.Execute(data)
    
    ' Проходим по каждому совпадению и заполняем ячейки
    For Each match In matches
        row = CInt(match.SubMatches(0))
        col = CInt(match.SubMatches(1))
        cellValue = match.SubMatches(2)
        
        ' Заполняем ячейку
        Cells(row, col).Value = cellValue
    Next match
End Sub

Sub FillCellsFromData2()
    Dim data As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Variant
    Dim cellAddress As String
    Dim cellValue As String
    Dim row As Integer
    Dim col As Integer
    
    ' Пример строки DATA
    data = "[[R2C18:2861000;R2C19:4013000]|[R3C18:2861000;R3C19:4013000]]"
    
    ' Создаем объект регулярного выражения
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "\[R(\d+)C(\d+):([^;]+)\;"
    
    ' Ищем все совпадения в строке DATA
    Set matches = regex.Execute(data)
    
    Dim yyyy
    
    
    ' Проходим по каждому совпадению и заполняем ячейки
    For Each match In matches
        yyyy = UBound(matches)
    
        row = CInt(match.SubMatches(0))
        col = CInt(match.SubMatches(1))
        cellValue = match.SubMatches(2)
        
        ' Заполняем ячейку
        Cells(row, col).Value = cellValue
    Next match
End Sub
