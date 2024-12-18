
'Option Explicit
'Option Compare Text

Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long

Dim myRegExp As New RegExp ' создаем экземпляр RegExp
Dim aMatch As match ' один из совпавших образцов
Dim colMatches As MatchCollection ' коллекция этих образцов

Dim SpisokFiles As FileDialogSelectedItems, File

Private Sub ScreenShotTelegramm(control As IRibbonControl)
    Call Выбрать_Имена_для_отправки_Скриншота_вТелеграм
End Sub
Private Sub SelectAll(control As IRibbonControl)
    Call Выбрать_всё
End Sub
Private Sub AI_Ribbon(control As IRibbonControl)
    Call Testing_Mistral_AI_2
End Sub
Private Sub SaveToXML(control As IRibbonControl)
    Call SelectionToXML
End Sub
Private Sub ImportFromXML(control As IRibbonControl)
    Call LoadDataFromXML_to_Range
End Sub
'airatButton4 (элемент: button, атрибут: onAction), 2007
Private Sub Fill_to_Right_Graded(control As IRibbonControl)
    Call Заполнить_вправо_с_уровнями
End Sub
'airatButton5 (элемент: button, атрибут: onAction), 2007
Private Sub Fill_to_Right(control As IRibbonControl)
    Call Заполнить_вправо
End Sub
'airatButton6 (элемент: button, атрибут: onAction), 2007
Private Sub Fill_to_Left(control As IRibbonControl)
    Call Заполнить_влево
End Sub

Function GetFilenamesCollection(Optional ByVal Title As String, Optional ByVal InitialPath As String) As FileDialogSelectedItems

    Dim FOLDER$
    ' функция выводит диалоговое окно выбора нескольких файлов с заголовком Title,
    ' начиная обзор диска с папки InitialPath
    ' возвращает массив путей к выбранным файлам, или пустую строку в случае отказа от выбора
    
    With Application.FileDialog(3) ' msoFileDialogFilePicker
        .ButtonName = "Выбрать"
        .Title = Title
        '.InitialFileName = InitialPath
        
        .InitialFileName = GetSetting(Application.name, "GetFilenamesCollection", "folder", InitialPath)
        
        If .Show <> -1 Then Exit Function
        Set GetFilenamesCollection = .SelectedItems
        
        FOLDER$ = Left(.SelectedItems(1), InStrRev(.SelectedItems(1), "\"))
        SaveSetting Application.name, "GetFilenamesCollection", "folder", FOLDER$
        
    End With
            
End Function
Function GetMemUsage()
    ' Returns the current Excel.Application memory usage in MB
    Set objSWbemServices = GetObject("winmgmts:")
    GetMemUsage = objSWbemServices.Get("Win32_Process.Handle='" & GetCurrentProcessId & "'").WorkingSetSize / 1024 / 1024
    MsgBox "Excel использует " & Round(GetMemUsage, 1) & " MB оперативки"
    Set objSWbemServices = Nothing
End Function

Sub FixNumberToTextByArray(ByRef rng As Range)
    Dim cell As Range
    
    Dim FileName As String
    Dim SheetName As String
    
    FileName = rng.Parent.Parent.name
    SheetName = rng.Parent.name
    
    Dim Matrix As Variant
    Matrix = rng.Value
       
    
    Dim ReturnedMatrix()
    ReDim ReturnedMatrix(UBound(Matrix) - 1, UBound(Matrix, 2) - 1)
        
    ' Пройдитесь по каждой ячейке в диапазоне
    
    For j = 1 To UBound(Matrix, 2)
        For i = 1 To UBound(Matrix)
            If IsNumeric(Matrix(i, j)) Then
                If Matrix(i, j) = "" Then
                    ReturnedMatrix(i - 1, j - 1) = ""
                Else
                    ReturnedMatrix(i - 1, j - 1) = CStr(Matrix(i, j))
                End If
            Else
                ReturnedMatrix(i - 1, j - 1) = Matrix(i, j)
            End If
        Next i
    Next j
          
    
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = Workbooks(FileName)
    Set ws = wb.Sheets(SheetName)
    
    ws.Cells(rng.row, rng.Column).Resize(UBound(ReturnedMatrix) + 1, UBound(ReturnedMatrix, 2) + 1) = ReturnedMatrix
    
End Sub

Sub FixTextToNumberByArray(ByRef rng As Range)
    Dim cell As Range
    
    Dim FileName As String
    Dim SheetName As String
    
    FileName = rng.Parent.Parent.name
    SheetName = rng.Parent.name
    
    Dim Matrix As Variant
    Matrix = rng.Value
       
    
    Dim ReturnedMatrix()
    ReDim ReturnedMatrix(UBound(Matrix) - 1, UBound(Matrix, 2) - 1)
        
    ' Пройдитесь по каждой ячейке в диапазоне
    
    For j = 1 To UBound(Matrix, 2)
        For i = 1 To UBound(Matrix)
            If IsNumeric(Matrix(i, j)) Then
                If Matrix(i, j) = "" Then
                    ReturnedMatrix(i - 1, j - 1) = ""
                Else
                    ReturnedMatrix(i - 1, j - 1) = CDbl(Matrix(i, j))
                End If
            Else
                ReturnedMatrix(i - 1, j - 1) = Matrix(i, j)
            End If
        Next i
    Next j
          
    
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = Workbooks(FileName)
    Set ws = wb.Sheets(SheetName)
    
    ws.Cells(rng.row, rng.Column).Resize(UBound(ReturnedMatrix) + 1, UBound(ReturnedMatrix, 2) + 1) = ReturnedMatrix
    
End Sub
Sub FixTextToNumberByObj(ByRef rng As Range)
    
    Dim cell As Range
    
    ' Пройдитесь по каждой ячейке в диапазоне
    For Each cell In rng
        Debug.Print IsNumeric(cell.Value)
        Debug.Print cell.NumberFormat
        If IsNumeric(cell.Value) And cell.NumberFormat = "General" Then
            ' Если ячейка содержит число, но отформатирована как текст, исправьте это
            cell.Value = CDbl(cell.Value)
            cell.NumberFormat = "General"
        End If
    Next cell
    
    MsgBox "Ошибки исправлены!", vbInformation
End Sub

Sub FixTextToNumber()
    Dim rng As Range
    Set rng = Selection
    
    Call FixTextToNumberByArray(rng)
End Sub

Sub FilesConnection()
    
    Set SpisokFiles = GetFilenamesCollection("Выбор ФАЙЛОВ", ThisWorkbook.Path)   'выводим окно выбора
    If SpisokFiles Is Nothing Then Exit Sub  'выход, если пользователь отказался от выбора файлов
    
Prepare
    
    
    Dim fileAdresses() As String, fileNames() As String
    Dim i As Integer
    i = 1
    
    For Each Item In SpisokFiles
        ReDim Preserve fileAdresses(i - 1) As String
        fileAdresses(i - 1) = SpisokFiles.Item(i)
        
        ReDim Preserve fileNames(i - 1) As String
        fileNames(i - 1) = Mid(SpisokFiles.Item(i), InStrRev(SpisokFiles.Item(i), "\") + 1, Len(SpisokFiles.Item(i)) - InStrRev(SpisokFiles.Item(i), "\"))
        
        i = i + 1
    Next Item
    

    
    Dim fileAdress As String, FileName As String

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\чек лист\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Dim indicatorImeniFila As Boolean
    indicatorImeniFila = False
    
    For i = 0 To UBound(fileNames)
        If InStr(1, fileNames(i), "чек") <> Empty Or InStr(1, fileNames(i), "Чек") <> Empty Then
            fileAdress = fileAdresses(i)
            FileName = fileNames(i)
            indicatorImeniFila = True
            Exit For
        End If
    Next i
    
    If indicatorImeniFila = False Then
        MsgBox ("Файла с именем содержащим слово чек среди выбранных файлов нет")
        Ended
        Exit Sub
    End If
    indicatorImeniFila = False
    
    'открытие файла чек лист
    Dim ChekListFile As Workbook
    For Each ChekListFile In Workbooks
        If ChekListFile.name = FileName Then
            GoTo PerehodChekList
        End If
    Next ChekListFile

    Workbooks.Open (fileAdress)
    
PerehodChekList:

    Set ChekListFile = Workbooks(FileName)
    
    ChekListFile.Activate
    
    
'Проверяем наличие листа с именем содержащим слово чек
    Dim indicatorImeniLista As Boolean
    indicatorImeniLista = False
    
    Dim ChekListSheet As Worksheet
    
    For Each ChekListSheet In ChekListFile.Sheets
        If InStr(1, ChekListSheet.name, "чек") <> Empty Or InStr(1, ChekListSheet.name, "Чек") <> Empty Then
            Set ChekListSheet = ChekListSheet
            indicatorImeniLista = True
            Exit For
        End If
    Next ChekListSheet
    
    If indicatorImeniLista = False Then
        MsgBox ("Листа с именем содержащим слово чек нет в файле " & ChekListFile.name)
        Set ChekListFile = Nothing
        Set ChekListSheet = Nothing
        Ended
        Exit Sub
    End If
    indicatorImeniLista = False
    
       
    
    'создадим общий файл
    Dim CommonFile As Workbook
    Set CommonFile = Workbooks.Add
    CommonFile.SaveAs FileName:=ChekListFile.Path & "\АвтоАнализ.xlsx"
    
    
    
    ChekListSheet.Copy after:=CommonFile.Sheets(CommonFile.Sheets.Count) 'копируем лист в Общий файл
    Set ChekListSheet = ActiveSheet
    
    ChekListFile.Close False
    Set ChekListFile = Nothing
    
    ChekListSheet.name = "чек_лист"
    
    Application.DisplayAlerts = False
    CommonFile.Sheets(1).Delete
    Application.DisplayAlerts = True
    
    
    
    Dim DateCell As Range
    Set DateCell = ChekListSheet.Cells.Find(What:="Дата", Lookat:=xlPart, LookIn:=xlFormulas, searchorder:=xlByRows, searchdirection:=xlNext)
    
    Dim DateCellRow As Long
    Dim DateCellColumn As Long

    DateCellRow = DateCell.row
    DateCellColumn = DateCell.Column
    
    Dim obrazecDate As String
    obrazecDate = ChekListSheet.Cells(DateCellRow + 1, DateCellColumn).Value
    
    ' Извлекаем год и месяц из даты
    Dim origYear As Integer
    Dim origMonth As Integer
    
    origYear = CInt(Format(obrazecDate, "YYYY"))
    origMonth = CInt(Format(obrazecDate, "mm"))
    
    ' Создаем новую дату, представляющую первый день этого месяца и года
    Dim obrazecMonth As Date
    obrazecMonth = DateSerial(origYear, origMonth, 1)
       

    'преобразуем Числовые Табельные номера в текстовые формат
    Dim chekListTabNomCell As Range
    Set chekListTabNomCell = ChekListSheet.Cells.Find(What:="ТабельныйНомер", Lookat:=xlPart, LookIn:=xlFormulas, searchorder:=xlByRows, searchdirection:=xlNext)
    
    Dim chekListTabNomCellRow As Long
    Dim chekListTabNomCellColumn As Long

    chekListTabNomCellRow = chekListTabNomCell.row
    chekListTabNomCellColumn = chekListTabNomCell.Column
    
    Dim chekListSheetlastRow As Long
    chekListSheetlastRow = ChekListSheet.Cells(ChekListSheet.Rows.Count, chekListTabNomCellColumn).End(xlUp).row
    
    
    Dim chekListTabNomRng As Range
    Set chekListTabNomRng = ChekListSheet.Range(ChekListSheet.Cells(chekListTabNomCellRow + 1, chekListTabNomCellColumn), ChekListSheet.Cells(chekListSheetlastRow, chekListTabNomCellColumn))
    
    'Call FixTextToNumberByArray(chekListTabNomRng)
    Call FixNumberToTextByArray(chekListTabNomRng)
    
    'так как в Yandex DataLense столбец с Таб.Номерами автоматически определяется как числовой, нам необходимо добавить любой текст в конец списка, чтобы столбец стал текстовым
    ChekListSheet.Cells(chekListSheetlastRow + 1, chekListTabNomCellColumn).Value = "тестовая строка, не удалять"
    
'////////////////////////////////////////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\счётчик\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Dim CounterSheet As Worksheet
    Set CounterSheet = CommonFile.Sheets.Add(after:=CommonFile.Sheets(CommonFile.Sheets.Count))
    CounterSheet.name = "счётчик"
    CounterSheet.Cells(1, 1).FormulaR1C1 = "=COUNTA(" & ChekListSheet.name & "!C10)"
    
'//////////////////////////////////////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\kpi\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'//////////////////////////////////////////////////////////////////////////////////////////////////////////
    indicatorImeniFila = False
    
    For i = 0 To UBound(fileNames)
        If InStr(1, fileNames(i), "kpi") <> Empty Or InStr(1, fileNames(i), "Kpi") <> Empty Then
            fileAdress = fileAdresses(i)
            FileName = fileNames(i)
            indicatorImeniFila = True
            Exit For
        End If
    Next i
    
    If indicatorImeniFila = False Then
        MsgBox ("Файла с именем содержащим слово kpi среди выбранных файлов нет")
        Ended
        Exit Sub
    End If
    indicatorImeniFila = False
    
    'открытие файла kpi
    Dim kpiFile As Workbook
    For Each kpiFile In Workbooks
        If kpiFile.name = FileName Then
            GoTo PerehodKpi
        End If
    Next kpiFile

    Workbooks.Open (fileAdress)
    
PerehodKpi:

    Set kpiFile = Workbooks(FileName)
    
    kpiFile.Activate
    
    
'Проверяем наличие листа с именем содержащим слово kpi
    indicatorImeniLista = False
    
    Dim kpiSheet As Worksheet
    
    For Each kpiSheet In kpiFile.Sheets
        If InStr(1, kpiSheet.name, "kpi") <> Empty Or InStr(1, kpiSheet.name, "Kpi") <> Empty Then
            Set kpiSheet = kpiSheet
            indicatorImeniLista = True
            Exit For
        End If
    Next kpiSheet
    
    If indicatorImeniLista = False Then
        MsgBox ("Листа с именем содержащим слово kpi нет в файле " & kpiFile.name)
        Set kpiFile = Nothing
        Set kpiSheet = Nothing
        Ended
        Exit Sub
    End If
    indicatorImeniLista = False
    
    
    kpiSheet.Copy after:=CommonFile.Sheets(CommonFile.Sheets.Count) 'копируем лист в Общий файл
    Set kpiSheet = ActiveSheet
    
    kpiFile.Close False
    Set kpiFile = Nothing
    
    kpiSheet.name = "kpi"
    
    kpiSheet.Columns(1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Dim kpiSheetlastRow As Long
    kpiSheetlastRow = kpiSheet.Cells(kpiSheet.Rows.Count, 2).End(xlUp).row
          
    
    'проставляем месяц
    kpiSheet.Cells(1, 1) = "Месяц kpi"
    kpiSheet.Range(kpiSheet.Cells(2, 1), kpiSheet.Cells(kpiSheetlastRow, 1)).Value = obrazecMonth
    
    kpiSheet.Range(kpiSheet.Cells(2, 1), kpiSheet.Cells(kpiSheetlastRow, 1)).NumberFormat = "dd.mm.yyyy;@"
    
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\штатные сотрудники\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    indicatorImeniFila = False
    
    For i = 0 To UBound(fileNames)
        If InStr(1, fileNames(i), "шр") <> Empty Or InStr(1, fileNames(i), "ШР") <> Empty Then
            fileAdress = fileAdresses(i)
            FileName = fileNames(i)
            indicatorImeniFila = True
            Exit For
        End If
    Next i
    
    If indicatorImeniFila = False Then
        MsgBox ("Файла с именем содержащим слово шр среди выбранных файлов нет")
        Ended
        Exit Sub
    End If
    indicatorImeniFila = False
    
    'открытие файла штатные сотрудники
    Dim ShtatFile As Workbook
    For Each ShtatFile In Workbooks
        If ShtatFile.name = FileName Then
            GoTo PerehodShtat
        End If
    Next ShtatFile

    Workbooks.Open (fileAdress)
    
PerehodShtat:

    Set ShtatFile = Workbooks(FileName)
    
    ShtatFile.Activate
    
    
'Проверяем наличие листа с именем содержащим слово Лист_1
    indicatorImeniLista = False
    
    Dim ShtatSheet As Worksheet
    
    For Each ShtatSheet In ShtatFile.Sheets
        If InStr(1, ShtatSheet.name, "Лист_1") <> Empty Or InStr(1, ShtatSheet.name, "лист_1") <> Empty Then
            Set ShtatSheet = ShtatSheet
            indicatorImeniLista = True
            Exit For
        End If
    Next ShtatSheet
    
    If indicatorImeniLista = False Then
        MsgBox ("Листа с именем содержащим слово Лист_1 нет в файле " & ShtatFile.name)
        Set ShtatFile = Nothing
        Set ShtatSheet = Nothing
        Ended
        Exit Sub
    End If
    indicatorImeniLista = False
    
    
    ShtatSheet.Copy after:=CommonFile.Sheets(CommonFile.Sheets.Count) 'копируем лист в Общий файл
    Set ShtatSheet = ActiveSheet
    
    ShtatFile.Close False
    Set ShtatFile = Nothing
    
    ShtatSheet.name = "штатные сотрудники"
    
    'делаем преобразование структуры (удаляем лишнее)
    ShtatSheet.Cells.UnMerge
    
    Dim sostoyaniyeCell As Range
    Set sostoyaniyeCell = ShtatSheet.Cells.Find(What:="Состояние", Lookat:=xlPart, LookIn:=xlFormulas, searchorder:=xlByRows, searchdirection:=xlNext)
    
    Dim sostoyaniyeCellRow As Long
    Dim sostoyaniyeCellColumn As Long

    sostoyaniyeCellRow = sostoyaniyeCell.row
    sostoyaniyeCellColumn = sostoyaniyeCell.Column

    ShtatSheet.Range(ShtatSheet.Cells(sostoyaniyeCellRow + 1, 1), ShtatSheet.Cells(sostoyaniyeCellRow + 1, sostoyaniyeCellColumn)).Value = ShtatSheet.Range(ShtatSheet.Cells(sostoyaniyeCellRow, 1), ShtatSheet.Cells(sostoyaniyeCellRow, sostoyaniyeCellColumn)).Value

    ShtatSheet.Range(ShtatSheet.Rows(1), ShtatSheet.Rows(sostoyaniyeCellRow)).Delete
    
    Dim ShtatSheetlastColumn As Long
    ShtatSheetlastColumn = ShtatSheet.Cells(1, ShtatSheet.Columns.Count).End(xlToLeft).Column
    
    
    
    For i = ShtatSheetlastColumn To 1 Step -1
        If ShtatSheet.Cells(1, i).Value = "" Then
            ShtatSheet.Columns(i).Delete
        End If
    Next i
    
    ShtatSheet.Columns(1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    
    
    'преобразуем Числовые Табельные номера в текстовые формат
    Dim tabNomCell As Range
    Set tabNomCell = ShtatSheet.Cells.Find(What:="Табельный номер", Lookat:=xlPart, LookIn:=xlFormulas, searchorder:=xlByRows, searchdirection:=xlNext)
    
    Dim tabNomCellRow As Long
    Dim tabNomCellColumn As Long

    tabNomCellRow = tabNomCell.row
    tabNomCellColumn = tabNomCell.Column
    
    Dim ShtatSheetlastRow As Long
    ShtatSheetlastRow = ShtatSheet.Cells(ShtatSheet.Rows.Count, 2).End(xlUp).row
    
    
    Dim tabNomRng As Range
    Set tabNomRng = ShtatSheet.Range(ShtatSheet.Cells(tabNomCellRow + 1, tabNomCellColumn), ShtatSheet.Cells(ShtatSheetlastRow, tabNomCellColumn))
    
    'Call FixTextToNumberByArray(tabNomRng)
    Call FixNumberToTextByArray(tabNomRng)
    
    'так как в Yandex DataLense столбец с Таб.Номерами автоматически определяется как числовой, нам необходимо добавить любой текст в конец списка, чтобы столбец стал текстовым
    ShtatSheet.Cells(ShtatSheetlastRow + 1, tabNomCellColumn).Value = "тестовая строка, не удалять"
    
    'проставляем месяц
    ShtatSheet.Cells(1, 1) = "Месяц ШР"
    ShtatSheet.Range(ShtatSheet.Cells(tabNomCellRow + 1, 1), ShtatSheet.Cells(ShtatSheetlastRow, 1)).Value = obrazecMonth
    
    
    
'\\\\\\\\\\\\\\\\\формулы для звеньев с вахтой\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Dim lastrChekListSheet As Long
    
    lastrChekListSheet = ChekListSheet.Cells(ChekListSheet.Rows.Count, 10).End(xlUp).row
    ChekListSheet.Range(ChekListSheet.Cells(2, 16), ChekListSheet.Cells(lastrChekListSheet, 16)).FormulaR1C1 = "=IF(IFERROR(SEARCH(""ахта"",VLOOKUP(RC[-7],'" & ShtatSheet.name & "'!C4:C9,6,0)),0),5&VLOOKUP(RC[-7],'" & ShtatSheet.name & "'!C4:C9,6,0),IF(IFERROR(SEARCH(""кользящийграфик №20"",VLOOKUP(RC[-7],'" & ShtatSheet.name & "'!C4:C9,6,0)),0),RC[-9]&VLOOKUP(RC[-7],'" & ShtatSheet.name & "'!C4:C9,6,0),RC[-9]))"
    'ChekListSheet.Range(ChekListSheet.Cells(2, 16), ChekListSheet.Cells(lastrChekListSheet, 16)).FormulaR1C1 = "=IF(IFERROR(SEARCH(""ахта"",VLOOKUP(RC[-7],'Часы 1с анализ сверка часов'!C[-14]:C[-10],5,0)),0),5&VLOOKUP(RC[-7],'Часы 1с анализ сверка часов'!C[-14]:C[-10],5,0),RC[-9])"
    ChekListSheet.Cells(1, 16).Value = "Звено с вахтой"

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\аутстафф\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////
    indicatorImeniFila = False
    
    For i = 0 To UBound(fileNames)
        If InStr(1, fileNames(i), "аутс") <> Empty Or InStr(1, fileNames(i), "Аутс") <> Empty Then
            fileAdress = fileAdresses(i)
            FileName = fileNames(i)
            indicatorImeniFila = True
            Exit For
        End If
    Next i
    
    If indicatorImeniFila = False Then
        MsgBox ("Файла с именем содержащим слово аутстафф среди выбранных файлов нет")
        Ended
        Exit Sub
    End If
    indicatorImeniFila = False
    
    'открытие файла Аутстафф
    Dim AutstaffFile As Workbook
    For Each AutstaffFile In Workbooks
        If AutstaffFile.name = FileName Then
            GoTo PerehodAutstaff
        End If
    Next AutstaffFile

    Workbooks.Open (fileAdress)
    
PerehodAutstaff:

    Set AutstaffFile = Workbooks(FileName)
    
    AutstaffFile.Activate
    
    
'Проверяем наличие листа с именем содержащим слово аутстафф
    indicatorImeniLista = False
    
    Dim AutstaffSheet As Worksheet
    
    For Each AutstaffSheet In AutstaffFile.Sheets
        If InStr(1, AutstaffSheet.name, "аутс") <> Empty Or InStr(1, AutstaffSheet.name, "Аутс") <> Empty Then
            Set AutstaffSheet = AutstaffSheet
            indicatorImeniLista = True
            Exit For
        End If
    Next AutstaffSheet
    
    If indicatorImeniLista = False Then
        MsgBox ("Листа с именем содержащим слово аутстафф нет в файле " & AutstaffFile.name)
        Set AutstaffFile = Nothing
        Set AutstaffSheet = Nothing
        Ended
        Exit Sub
    End If
    indicatorImeniLista = False
    
    
    AutstaffSheet.Copy after:=CommonFile.Sheets(CommonFile.Sheets.Count) 'копируем лист в Общий файл
    Set AutstaffSheet = ActiveSheet
    
    AutstaffFile.Close False
    Set AutstaffFile = Nothing
    
    AutstaffSheet.name = "аутстафф"
    
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\Вне табеля\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    indicatorImeniFila = False
    
    For i = 0 To UBound(fileNames)
        If InStr(1, fileNames(i), "вне таб") <> Empty Or InStr(1, fileNames(i), "Вне таб") <> Empty Then
            fileAdress = fileAdresses(i)
            FileName = fileNames(i)
            indicatorImeniFila = True
            Exit For
        End If
    Next i
    
    If indicatorImeniFila = False Then
        MsgBox ("Файла с именем содержащим слово Вне табеля среди выбранных файлов нет")
        Ended
        Exit Sub
    End If
    indicatorImeniFila = False
    
    'открытие файла Вне табеля
    Dim VneTabFile As Workbook
    For Each VneTabFile In Workbooks
        If VneTabFile.name = FileName Then
            GoTo PerehodVneTab
        End If
    Next VneTabFile

    Workbooks.Open (fileAdress)
    
PerehodVneTab:

    Set VneTabFile = Workbooks(FileName)
    
    VneTabFile.Activate
    
    
'Проверяем наличие листа с именем содержащим слово Лист_1
    indicatorImeniLista = False
    
    Dim VneTabSheet As Worksheet
    
    For Each VneTabSheet In VneTabFile.Sheets
        If InStr(1, VneTabSheet.name, "Лист_1") <> Empty Or InStr(1, VneTabSheet.name, "лист_1") <> Empty Then
            Set VneTabSheet = VneTabSheet
            indicatorImeniLista = True
            Exit For
        End If
    Next VneTabSheet
    
    If indicatorImeniLista = False Then
        MsgBox ("Листа с именем содержащим слово Лист_1 нет в файле " & VneTabFile.name)
        Set VneTabFile = Nothing
        Set VneTabSheet = Nothing
        Ended
        Exit Sub
    End If
    indicatorImeniLista = False
    
    
    VneTabSheet.Copy after:=CommonFile.Sheets(CommonFile.Sheets.Count) 'копируем лист в Общий файл
    Set VneTabSheet = ActiveSheet
    
    VneTabFile.Close False
    Set VneTabFile = Nothing
    
    VneTabSheet.name = "Вне табеля"
     
    
    
    
    
    
    
    'делаем преобразование структуры (удаляем лишнее)
    VneTabSheet.Cells.UnMerge
    
    Dim sotrudnikCell As Range
    Set sotrudnikCell = VneTabSheet.Cells.Find(What:="Сотрудник", Lookat:=xlPart, LookIn:=xlFormulas, searchorder:=xlByRows, searchdirection:=xlNext)
    
    Dim sotrudnikCellRow As Long
    Dim sotrudnikCellColumn As Long

    sotrudnikCellRow = sotrudnikCell.row
    sotrudnikCellColumn = sotrudnikCell.Column
    
    VneTabSheet.Range(VneTabSheet.Rows(1), VneTabSheet.Rows(sotrudnikCellRow - 1)).Delete
    
    Dim lastColumnVneTabSheet As Long
    lastColumnVneTabSheet = VneTabSheet.Cells(1, VneTabSheet.Columns.Count).End(xlToLeft).Column
    
    
    
    For i = lastColumnVneTabSheet To 1 Step -1
        If VneTabSheet.Cells(1, i).Value = "" Then
            VneTabSheet.Columns(i).Delete
        End If
    Next i
    
    
    Dim lastrVneTabSheet As Long
    lastrVneTabSheet = VneTabSheet.Cells(VneTabSheet.Rows.Count, 2).End(xlUp).row
    
    
    VneTabSheet.Range(VneTabSheet.Cells(2, 1), VneTabSheet.Cells(lastrVneTabSheet, 1)).FormulaR1C1 = "=RC3&RC6&TEXT(RC9,""0"")"
    VneTabSheet.Cells(1, 1).Value = "Для ВПР"
    
    'поиск шапки Отработано часов
    Dim OtrabotanoChasCell As Range
    Set OtrabotanoChasCell = VneTabSheet.Cells.Find(What:="Отработано часов", Lookat:=xlPart, LookIn:=xlFormulas, searchorder:=xlByRows, searchdirection:=xlNext)
        
    Dim OtrabotanoChasCellColumn As Long
    OtrabotanoChasCellColumn = OtrabotanoChasCell.Column
            
    'поиск шапки Таб. номер
    Dim TabNomVneTabCell As Range
    Set TabNomVneTabCell = VneTabSheet.Cells.Find(What:="Таб. номер", Lookat:=xlPart, LookIn:=xlFormulas, searchorder:=xlByRows, searchdirection:=xlNext)
        
    Dim TabNomVneTabCellColumn As Long
    TabNomVneTabCellColumn = TabNomVneTabCell.Column
        
    'преобразуем Числовые Табельные номера в текстовые формат
    Dim VneTabTabNomRng As Range
    Set VneTabTabNomRng = VneTabSheet.Range(VneTabSheet.Cells(2, TabNomVneTabCellColumn), VneTabSheet.Cells(lastrVneTabSheet, TabNomVneTabCellColumn))
    
    'Call FixTextToNumberByArray(VneTabTabNomRng)
    Call FixNumberToTextByArray(VneTabTabNomRng)
    
           
    'поиск шапки Дата
    Dim DateVneTabCell As Range
    Set DateVneTabCell = VneTabSheet.Cells.Find(What:="Дата", Lookat:=xlPart, LookIn:=xlFormulas, searchorder:=xlByRows, searchdirection:=xlNext)
        
    Dim DateVneTabCellColumn As Long
    DateVneTabCellColumn = DateVneTabCell.Column
    
    'поиск шапки Проведен
    Dim ProvedenVneTabCell As Range
    Set ProvedenVneTabCell = VneTabSheet.Cells.Find(What:="Проведен", Lookat:=xlPart, LookIn:=xlFormulas, searchorder:=xlByRows, searchdirection:=xlNext)
        
    Dim ProvedenVneTabCellColumn As Long
    ProvedenVneTabCellColumn = ProvedenVneTabCell.Column
    
    
    
    ChekListSheet.Range(ChekListSheet.Cells(2, 17), ChekListSheet.Cells(lastrChekListSheet, 17)).FormulaR1C1 = "=SUMIFS('" & VneTabSheet.name & "'!C" & OtrabotanoChasCellColumn & ",'" & VneTabSheet.name & "'!C" & TabNomVneTabCellColumn & ",RC9,'" & VneTabSheet.name & "'!C" & DateVneTabCellColumn & ",RC5,'" & VneTabSheet.name & "'!C" & ProvedenVneTabCellColumn & ",""Да"")"
    ChekListSheet.Cells(1, 17).Value = "Часы (Вне Табеля)"
    
    ChekListSheet.Range(ChekListSheet.Cells(2, 18), ChekListSheet.Cells(lastrChekListSheet, 18)).FormulaR1C1 = "=SUMIFS('" & VneTabSheet.name & "'!C" & OtrabotanoChasCellColumn & ",'" & VneTabSheet.name & "'!C" & TabNomVneTabCellColumn & ",RC9,'" & VneTabSheet.name & "'!C" & DateVneTabCellColumn & ",RC5,'" & VneTabSheet.name & "'!C" & ProvedenVneTabCellColumn & ",""Нет"")"
    ChekListSheet.Cells(1, 18).Value = "Часы (Празд.Дни)"
    
    
    ChekListSheet.Range(ChekListSheet.Cells(2, 19), ChekListSheet.Cells(lastrChekListSheet, 19)).FormulaR1C1 = "=IFERROR(VLOOKUP(""Да""&RC9&RC5,'" & VneTabSheet.name & "'!C1:C" & OtrabotanoChasCellColumn & "," & OtrabotanoChasCellColumn - 1 & ",0),0)"
    ChekListSheet.Cells(1, 19).Value = "Способ компенсации (Вне табеля)"
    
    ChekListSheet.Range(ChekListSheet.Cells(2, 20), ChekListSheet.Cells(lastrChekListSheet, 20)).FormulaR1C1 = "=IFERROR(VLOOKUP(""Нет""&RC9&RC5,'" & VneTabSheet.name & "'!C1:C" & OtrabotanoChasCellColumn & "," & OtrabotanoChasCellColumn - 1 & ",0),0)"
    ChekListSheet.Cells(1, 20).Value = "Способ компенсации (Празд.Дни)"
    
    
    
    
   '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\Количество суток\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    indicatorImeniFila = False
    
    For i = 0 To UBound(fileNames)
        If InStr(1, fileNames(i), "кол_во суток") <> Empty Or InStr(1, fileNames(i), "Кол_во суток") <> Empty Then
            fileAdress = fileAdresses(i)
            FileName = fileNames(i)
            indicatorImeniFila = True
            Exit For
        End If
    Next i
    
    If indicatorImeniFila = False Then
        MsgBox ("Файла с именем содержащим слово Кол_во суток среди выбранных файлов нет")
        Ended
        Exit Sub
    End If
    indicatorImeniFila = False
    
    'открытие файла Кол_во суток
    Dim Kol_vo_sutokFile As Workbook
    For Each Kol_vo_sutokFile In Workbooks
        If Kol_vo_sutokFile.name = FileName Then
            GoTo PerehodKol_vo_sutok
        End If
    Next Kol_vo_sutokFile

    Workbooks.Open (fileAdress)
    
PerehodKol_vo_sutok:

    Set Kol_vo_sutokFile = Workbooks(FileName)
    
    Kol_vo_sutokFile.Activate
    
    
'Проверяем наличие листа с именем содержащим слово кол_во суток
    indicatorImeniLista = False
    
    Dim Kol_vo_sutokSheet As Worksheet
    
    For Each Kol_vo_sutokSheet In Kol_vo_sutokFile.Sheets
        If InStr(1, Kol_vo_sutokSheet.name, "кол_во суток") <> Empty Or InStr(1, Kol_vo_sutokSheet.name, "Кол_во суток") <> Empty Then
            Set Kol_vo_sutokSheet = Kol_vo_sutokSheet
            indicatorImeniLista = True
            Exit For
        End If
    Next Kol_vo_sutokSheet
    
    If indicatorImeniLista = False Then
        MsgBox ("Листа с именем содержащим слово кол_во суток нет в файле " & Kol_vo_sutokFile.name)
        Set Kol_vo_sutokFile = Nothing
        Set Kol_vo_sutokSheet = Nothing
        Ended
        Exit Sub
    End If
    indicatorImeniLista = False
    
    
    Kol_vo_sutokSheet.Copy after:=CommonFile.Sheets(CommonFile.Sheets.Count) 'копируем лист в Общий файл
    Set Kol_vo_sutokSheet = ActiveSheet
    
    Kol_vo_sutokFile.Close False
    Set Kol_vo_sutokFile = Nothing
    
    Kol_vo_sutokSheet.name = "кол_во суток"
    
    
Ended
    
    
    CommonFile.Save
       
    MsgBox "Файл для АвтоАнализа создан. Замените им старый файл на Яндекс Диске"
    
End Sub

Public Sub Prepare()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.CalculateBeforeSave = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
End Sub

Public Sub Ended()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateBeforeSave = True
    Application.EnableEvents = True
    'ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
End Sub

Function ВПР2(Искомое_значение As Variant, Таблица As Variant, Номер_столбца_где_ищем As Long, Номер_столбца_из_которого_берем_значение As Long, Номер_повтора As Long)
    Dim i As Long, iCount As Long
    Select Case TypeName(Таблица)
    Case "Range"
    For i = 1 To Таблица.Rows.Count
        If Таблица.Cells(i, Номер_столбца_где_ищем) = Искомое_значение Then
            iCount = iCount + 1
        End If
        If iCount = Номер_повтора Then
            ВПР2 = Таблица.Cells(i, Номер_столбца_из_которого_берем_значение)
    Exit For
    End If
    Next i
    Case "Variant()"
        For i = 1 To UBound(Table)
            If Таблица(i, 1) = Искомое_значение Then iCount = iCount + 1
                If iCount = Номер_повтора Then
                    ВПР2 = Таблица(i, Номер_столбца_из_которого_берем_значение)
        Exit For
        End If
        Next i
    End Select
End Function

Function ВПР_Быстрый(Искомое_значение As Variant, Таблица As Range, Номер_столбца_где_ищем As Integer, Номер_столбца_из_которого_берем_значение As Integer, Номер_повтора As Integer)
    Dim i As Long, iCount As Long
    Dim avArr
    avArr = Таблица.Value
    For i = 1 To UBound(avArr, 1)
        If avArr(i, Номер_столбца_где_ищем) = Искомое_значение Then
            iCount = iCount + 1
        End If
        If iCount = Номер_повтора Then
            ВПР_Быстрый = avArr(i, Номер_столбца_из_которого_берем_значение)
            Exit For
        End If
    Next i
End Function

'Function regEx(strInput As String, matchPattern As String, Optional ByVal outputPattern As String = "$0") As Variant
'    Dim inputRegexObj As New VBScript_RegExp_55.RegExp, outputRegexObj As New VBScript_RegExp_55.RegExp, outReplaceRegexObj As New VBScript_RegExp_55.RegExp
'    Dim inputMatches As Object, replaceMatches As Object, replaceMatch As Object
'    Dim replaceNumber As Integer
'
'    With inputRegexObj
'        .Global = True
'        .MultiLine = True
'        .IgnoreCase = False
'        .Pattern = matchPattern
'    End With
'    With outputRegexObj
'        .Global = True
'        .MultiLine = True
'        .IgnoreCase = False
'        .Pattern = "\$(\d+)"
'    End With
'    With outReplaceRegexObj
'        .Global = True
'        .MultiLine = True
'        .IgnoreCase = False
'    End With
'
'    Set inputMatches = inputRegexObj.Execute(strInput)
'    If inputMatches.Count = 0 Then
'        regEx = False
'    Else
'        Set replaceMatches = outputRegexObj.Execute(outputPattern)
'        For Each replaceMatch In replaceMatches
'            replaceNumber = replaceMatch.SubMatches(0)
'            outReplaceRegexObj.Pattern = "\$" & replaceNumber
'
'            If replaceNumber = 0 Then
'                outputPattern = outReplaceRegexObj.Replace(outputPattern, inputMatches(0).Value)
'            Else
'                If replaceNumber > inputMatches(0).SubMatches.Count Then
'                    'regex = "A to high $ tag found. Largest allowed is $" & inputMatches(0).SubMatches.Count & "."
'                    regEx = CVErr(xlErrValue)
'                    Exit Function
'                Else
'                    outputPattern = outReplaceRegexObj.Replace(outputPattern, inputMatches(0).SubMatches(replaceNumber - 1))
'                End If
'            End If
'        Next
'        regEx = outputPattern
'    End If
'End Function
'
'Function simpleCellRegex(Myrange As Range) As String
'    Dim regEx As New RegExp
'    Dim strPattern As String
'    Dim strInput As String
'    Dim strReplace As String
'    Dim strOutput As String
'
'
'    strPattern = "^[0-9]{1,3}"
'
'    If strPattern <> "" Then
'        strInput = Myrange.Value
'        strReplace = ""
'
'        With regEx
'            .Global = True
'            .MultiLine = True
'            .IgnoreCase = False
'            .Pattern = strPattern
'        End With
'
'        If regEx.Test(strInput) Then
'            simpleCellRegex = regEx.Replace(strInput, strReplace)
'        Else
'            simpleCellRegex = "Not matched"
'        End If
'    End If
'End Function
'
Public Function RegExpExtract(text As String, Pattern As String, Optional Item As Integer = 1) As String
    Dim regex As New RegExp
    On Error GoTo ErrHandl
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = Pattern
    regex.Global = True
    If regex.Test(text) Then
        Set matches = regex.Execute(text)
        RegExpExtract = matches.Item(Item - 1)
        Exit Function
    End If
ErrHandl:
    RegExpExtract = CVErr(xlErrValue)
End Function

Public Function RegExpExtractMulti(text As String, Pattern As String, Optional Item As Integer = 1) As String
    Dim regex As New RegExp
    Dim Virazhenie
    On Error GoTo ErrHandl
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = Pattern
    regex.Global = True
    If regex.Test(text) Then
        Set matches = regex.Execute(text)
        For i = 0 To matches.Count - 1
            Virazhenie = Virazhenie & matches(i)
        Next i
        RegExpExtractMulti = Virazhenie
        Exit Function
    End If
ErrHandl:
    RegExpExtractMulti = CVErr(xlErrValue)
End Function

Public Function RegExpExtractMultiCell(text As String, Pattern As String, Optional Item As Integer = 1) As String
    Dim regex As New RegExp
    Dim Virazhenie
    On Error GoTo ErrHandl
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = Pattern
    regex.Global = True
    If regex.Test(text) Then
        Set matches = regex.Execute(text)
        For i = 1 To matches.Count - 1
            ActiveSheet.Cells(ActiveCell.row, ActiveCell.Column + i).Value = matches(i)
        Next i
        RegExpExtractMultiCell = matches(0)
        Exit Function
    End If
ErrHandl:
    RegExpExtractMultiCell = CVErr(xlErrValue)
End Function

Function my_sum1(A As Range, B As Range)
    Dim i&, j&, av(), bv()
    av = A.Value
    bv = B.Value
    ReDim v(1 To UBound(av), 1 To UBound(av, 2))
    For i = 1 To UBound(v)
        For j = 1 To UBound(v, 2)
            v(i, j) = av(i, j) + bv(i, j)
        Next
    Next
    my_sum1 = v
End Function

Sub Fix_Numbers_From_Dates()
    Dim num As Double, cell As Range
 
    For Each cell In Selection
        If Not IsEmpty(cell) Then
            If cell.NumberFormat = "General" Then
                num = CDbl(Replace(cell, ".", ","))
            Else
                num = CDbl(Format(cell, "m,yyyy"))
            End If
            cell.Clear
            cell.Value = num
        End If
    Next cell
End Sub

Sub Заполнить_влево()

On Error GoTo ErrorHandler

    Dim response
        response = MsgBox("Точно запустить?", vbYesNo)
        If response = vbNo Then Exit Sub


Dim oldSheet As Worksheet: Set oldSheet = ActiveSheet

Dim newSheet As Worksheet
ActiveWorkbook.Sheets(oldSheet.name).Copy after:=Sheets(1)
Set newSheet = ActiveSheet
newSheet.name = oldSheet.name & " Копия"

newSheet.Activate


    Dim cell As Range, av()
    
    Selection.UnMerge
    av = Selection.Value
    
    Dim lastStr As Long, lastClmn As Long
    lastStr = UBound(av)
    lastClmn = UBound(av, 2)
    
    Dim firstRow As Long, firstClmn As Long
    
    For Each cell In Selection
        firstRow = cell.row
        firstClmn = cell.Column
        Exit For
    Next cell
    
    'Debug.Print (firstRow & " " & firstClmn)
    'Debug.Print (av(2, 2))

'инвертируем плоскую таблицу справа на лево
    For i = 1 To lastStr
        For j = lastClmn To 2 Step -1
            If av(i, j - 1) = "" Then
                av(i, j - 1) = av(i, j)
            End If
        Next j
    Next i
    
    newSheet.Range(newSheet.Cells(firstRow, firstClmn), newSheet.Cells(firstRow + lastStr - 1, firstClmn + lastClmn - 1)) = av

Exit Sub
ErrorHandler:
MsgBox ("Необходимо выбрать минимум 2 ячейки!!!")

End Sub

Sub Заполнить_вправо()

On Error GoTo ErrorHandler

    Dim response
        response = MsgBox("Точно запустить?", vbYesNo)
        If response = vbNo Then Exit Sub


Dim oldSheet As Worksheet: Set oldSheet = ActiveSheet

Dim newSheet As Worksheet
ActiveWorkbook.Sheets(oldSheet.name).Copy after:=Sheets(1)
Set newSheet = ActiveSheet
newSheet.name = oldSheet.name & " Копия"

newSheet.Activate


    Dim cell As Range, av()
    
    Selection.UnMerge
    av = Selection.Value
    
    Dim lastStr As Long, lastClmn As Long
    lastStr = UBound(av)
    lastClmn = UBound(av, 2)
    
    Dim firstRow As Long, firstClmn As Long
    
    For Each cell In Selection
        firstRow = cell.row
        firstClmn = cell.Column
        Exit For
    Next cell
    
    'Debug.Print (firstRow & " " & firstClmn)
    'Debug.Print (av(2, 2))

'инвертируем плоскую таблицу слева на право
    For i = 1 To lastStr
        For j = 1 To lastClmn - 1
            If av(i, j + 1) = "" Then
                av(i, j + 1) = av(i, j)
            End If
        Next j
    Next i
    
    newSheet.Range(newSheet.Cells(firstRow, firstClmn), newSheet.Cells(firstRow + lastStr - 1, firstClmn + lastClmn - 1)) = av

Exit Sub
ErrorHandler:
MsgBox ("Необходимо выбрать минимум 2 ячейки!!!")

End Sub

Sub Заполнить_вправо_с_уровнями()


'On Error GoTo ErrorHandler

    Dim response
        response = MsgBox("Точно запустить?", vbYesNo)
        If response = vbNo Then Exit Sub
        
Dim oldSheet As Worksheet: Set oldSheet = ActiveSheet

Dim newSheet As Worksheet
ActiveWorkbook.Sheets(oldSheet.name).Copy after:=Sheets(1)
Set newSheet = ActiveSheet
newSheet.name = oldSheet.name & " Копия"

newSheet.Activate


    Dim cell As Range, av(), bv()
    Dim k As Long
    
    
    'Поиск слова СОТРУДНИК////////////
    Dim RowSotrudnik As Long, clmnSotrudnik As Long
    For k = 1 To 1000
        For j = 1 To 100
            If newSheet.Cells(k, j).Value = "Сотрудник" Then
                RowSotrudnik = k
                clmnSotrudnik = j
                GoTo VihodSotrudnik
            End If
        Next j
    Next k
    
    If clmnSotrudnik = Empty Then
        MsgBox ("В файле " & ActiveWorkbook.name & " на листе " & newSheet.name & " нет шапки со словом СОТРУДНИК")
        Exit Sub
    End If
    
VihodSotrudnik:


Dim firstRow As Long, firstClmn As Long
    
    For Each cell In Selection
        firstRow = cell.row
        firstClmn = cell.Column
        Exit For
    Next cell
    
    bv = Selection.Value
    
    Dim lastStr As Long, lastClmn As Long
    lastStr = UBound(bv)
    lastClmn = UBound(bv, 2)
    
    
    av = newSheet.Range(newSheet.Cells(firstRow, clmnSotrudnik), newSheet.Cells(firstRow + lastStr - 1, clmnSotrudnik)).Value
      
    'проверяем, есть ли пустые данные в столбце Сотрудник. Если есть пустые данные, то это означает корявость выгрузки,
    'и необходимо дополнительно найти столбец Подразделение, чтобы восполнить данными массив av() столбца Сотрудник
    Dim indicatorPustot As Boolean: indicatorPustot = False
    
    For i = 1 To UBound(av)
        If av(i, 1) = Empty Then
            indicatorPustot = True
            Exit For
        End If
    Next i
    
    
    
      
    If indicatorPustot = True Then
        'Поиск слова Подразделение////////////
        Dim RowPodrazd As Long, clmnPodrazd As Long
        For k = 1 To 1000
            For j = 1 To 100
                If newSheet.Cells(k, j).Value = "Подразделение" Then
                    RowPodrazd = k
                    clmnPodrazd = j
                    GoTo VihodPodrazd
                End If
            Next j
        Next k
        
        If clmnPodrazd = Empty Then
            MsgBox ("Так как в файле " & ActiveWorkbook.name & " на листе " & newSheet.name & " в столбце Сотрудник нет данных о вышестоящих подразделениях, а только ФИО, то необходим столбец Подразделение, для вытягивания информации, а Подразделение нет в шапке")
            Exit Sub
        End If
        
VihodPodrazd:


        Dim massivPodrazd()
        massivPodrazd = newSheet.Range(newSheet.Cells(firstRow, clmnPodrazd), newSheet.Cells(firstRow + lastStr - 1, clmnPodrazd)).Value
        
        For i = 1 To UBound(av)
            If av(i, 1) = Empty Then
                av(i, 1) = massivPodrazd(i, 1)
            End If
        Next i
    
    End If
    
    
    
    
    Dim urovniVhodimosti()
    ReDim urovniVhodimosti(1 To lastStr)
            
'проставляем уровни входимости
    Dim MaxUrovenVhodimosti As Long, MinUrovenVhodimosti As Long
    
    
    For i = firstRow To firstRow + lastStr - 1
        urovniVhodimosti(i - firstRow + 1) = Selection.Rows(i - firstRow + 1).OutlineLevel
        If Selection.Rows(i - firstRow + 1).OutlineLevel > MaxUrovenVhodimosti Then
            MaxUrovenVhodimosti = Selection.Rows(i - firstRow + 1).OutlineLevel
        End If
    Next i
    
    MinUrovenVhodimosti = MaxUrovenVhodimosti
    
    For i = firstRow To firstRow + lastStr - 1
        If Selection.Rows(i - firstRow + 1).OutlineLevel < MinUrovenVhodimosti Then
            MinUrovenVhodimosti = Selection.Rows(i - firstRow + 1).OutlineLevel
        End If
    Next i
    
    
    
'матрица уровней в плоскости
    Dim Matrix(), razmerClmn As Long
    razmerClmn = MaxUrovenVhodimosti - MinUrovenVhodimosti + 1
    ReDim Matrix(1 To lastStr, 1 To razmerClmn)
    
    Dim shapkaUrovney()
    ReDim shapkaUrovney(1 To razmerClmn)
    
    k = 0
    For j = 1 To razmerClmn
        shapkaUrovney(j) = MinUrovenVhodimosti + k
        k = k + 1
    Next j



 'заполняем влево
    For i = 1 To UBound(av)
        For j = 1 To UBound(Matrix, 2)
            If urovniVhodimosti(i) = shapkaUrovney(j) Then
                For k = j To UBound(Matrix, 2)
                    Matrix(i, k) = av(i, 1)
                Next k
                Exit For
            End If
        Next j
    Next i
   
    
    
 'заполняем вниз
    Dim VremennoePodrazd As String
    For j = 1 To UBound(Matrix, 2) - 1
        For i = 1 To UBound(av)
            If Matrix(i, j) <> Empty Then
                VremennoePodrazd = Matrix(i, j)
            Else
                Matrix(i, j) = VremennoePodrazd
            End If
        Next i
    Next j
    

    

    
'инвертируем плоскую таблицу справа на лево
    Dim indicatorNepovtorimosti As Boolean
    Dim p As Long
    
    For i = 1 To UBound(av)
        For j = UBound(Matrix, 2) To 1 Step -1
        indicatorNepovtorimosti = False
            For k = j - 1 To 1 Step -1
                If Matrix(i, j) <> Matrix(i, k) Then
                    For p = k To j - 1
                    Matrix(i, p) = Matrix(i, k)
                    
                    indicatorNepovtorimosti = True
                    Next p
                    Exit For
                End If
            Next k
            If indicatorNepovtorimosti = False Then Exit For
        Next j
    Next i
    
    
'For i = 1 To UBound(matrix)
'    For j = 1 To UBound(matrix, 2)
'        Debug.Print (matrix(i, j))
'    Next j
'Next i




    
'Добавляем столбцы
    newSheet.Range(newSheet.Columns(1), newSheet.Columns(UBound(Matrix, 2))).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
'Вставляем Матрицу
    newSheet.Cells(firstRow, 1).Resize(UBound(Matrix), UBound(Matrix, 2)) = Matrix
    
    Exit Sub
ErrorHandler:
MsgBox ("Необходимо выбрать минимум 2 вертикальные ячейки!!!")

    
    
End Sub


Sub Выбрать_всё()
    On Error GoTo Err
    
    lastRow = ActiveSheet.Cells.Find(What:="*", Lookat:=xlPart, LookIn:=xlFormulas, searchorder:=xlByRows, searchdirection:=xlPrevious).row
    lastColumn = ActiveSheet.Cells.Find(What:="*", Lookat:=xlPart, LookIn:=xlFormulas, searchorder:=xlByColumns, searchdirection:=xlPrevious).Column
    ActiveSheet.Range(Cells(1, 1), Cells(lastRow, lastColumn)).Select
    Exit Sub
            
Err:
        MsgBox ("нет заполненных ячеек")
        
End Sub
Sub Выбрать_Имена_для_отправки_Скриншота_вТелеграм()
    UserForm1.Show
End Sub
Sub Сохранить_Скрин()
    Dim MassivFilesAdresses() As String
    MassivFilesAdresses = Save_to_jpg_HQ
    
    Dim FilesAdressesCount As Long
    FilesAdressesCount = UBound(MassivFilesAdresses)
    
    If FilesAdressesCount = 0 Then
        MsgBox "Скриншот сохранён по адресу: " & MassivFilesAdresses(0)
    Else
        MsgBox "Скриншоты сохранены по адресу: " & Application.ActiveWorkbook.Path
    End If
End Sub
Sub Сохранить_и_Отправить_Скрин_Телеграм(CHAT_ID As String, selectedName As String)
      
    Dim MassivFilesAdresses() As String
    MassivFilesAdresses = Save_to_jpg_HQ
      
    Dim botToken As String
    botToken = "***YOUR_TELEGRAM_BOT_TOKEN"
        
    Dim FileName As String
    Dim SheetName As String
    Dim Diapazon As String
    Dim data As String
    Dim Number As String
    
    Dim textMessage As String
    Dim fullAdress As String
    
    For i = 0 To UBound(MassivFilesAdresses)
        fullAdress = MassivFilesAdresses(i)
                                  
        ' Извлечение названия файла
        FileName = ExtractBetween(fullAdress, "файл(", ")")
        
        ' Извлечение названия листа
        SheetName = ExtractBetween(fullAdress, "лист(", ")")
        
        ' Извлечение диапазона
        Diapazon = ExtractBetween(fullAdress, "диапазон(", ")")
        
        ' Извлечение даты
        data = ExtractBetween(fullAdress, "дата(", ")")
        
        ' Извлечение цифры перед ".jpg"
        Number = ExtractNumberBeforeJpg(fullAdress)
        
        textMessage = "Пользователь: " & Application.UserName & vbCrLf & "Файл: " & FileName & vbCrLf & "Лист: " & SheetName & vbCrLf & "Диапазон: " & Diapazon & vbCrLf & "Дата: " & data & vbCrLf & "№: " & Number
    
        Call sendTextToTelegram(textMessage, CHAT_ID, botToken)
        
        Call telegram_send_JPG(fullAdress, CHAT_ID, botToken)
        
        
    Next i
    
    Dim FilesAdressesCount As Long
    FilesAdressesCount = UBound(MassivFilesAdresses)
    
    If FilesAdressesCount = 0 Then
        MsgBox "Скриншот отправлен пользователю " & selectedName & "Скриншот сохранён по адресу: " & MassivFilesAdresses(0)
    Else
        MsgBox "Скриншоты отправлены пользователю " & selectedName & "Скриншоты сохранены по адресу: " & Application.ActiveWorkbook.Path
    End If
                
End Sub
Function Save_to_jpg_HQ()
    
    Dim FileName As String, SheetName As String, rngAdress As String
    Dim sPath As String, name As String, formatF As String, fullName As String, fullAdress As String
    Dim rngs As Range, rng As Range
    
    Dim i As Long, j As Long
    
    
    Set rngs = Selection
    
    Dim timeNow
    timeNow = "дата(" & Replace(Replace(Now, ":", "."), " ", "_") & ")"
     
    Dim MassivFilesAdresses() As String
        
    i = 1
    For Each rng In rngs.Areas
        sPath = ActiveWorkbook.Path
        FileName = rng.Parent.Parent.name
        SheetName = rng.Parent.name
        fullName = sPath & "\файл(" & FileName & ")лист(" & SheetName & ")"
        formatF = "jpg"
    
        rngAdress = "диапазон(" & Replace(Replace(rng.Address, "$", ""), ":", "_") & ")"
        fullAdress = fullName & rngAdress & timeNow & "_" & i & "." & formatF
    
          Debug.Print fullAdress
          
          Dim shTemp As Worksheet
          Dim chrt As Chart
          
          rng.CopyPicture Appearance:=xlScreen, Format:=xlBitmap '(если поставить xlPicture, то качество будет хуже, но файл будет меньше весить)
          
          Set shTemp = Workbooks.Add.Worksheets(1)
          
          Set chrt = shTemp.ChartObjects.Add(Left:=0, Top:=0, Width:=rng.Width, Height:=rng.Height).Chart
          
          'создаём иммитацию действий, чтобы скриншот успел сохраниться, иначе будет белый скриншот
          For j = 0 To 5
              DoEvents
          Next j
        
          With chrt
              .ChartArea.Border.LineStyle = 0
              .Paste
              .Export FileName:=fullAdress
              .Parent.Delete
          End With
          shTemp.Parent.Close 0
          
          ReDim Preserve MassivFilesAdresses(i - 1) As String
          MassivFilesAdresses(i - 1) = fullAdress
        
        i = i + 1
    Next rng
    
    Save_to_jpg_HQ = MassivFilesAdresses
    
End Function
Sub telegram_send_JPG(fullAdress As String, Optional CHAT_ID As String = "***YOUR_TELEGRAM_CHAT_ID***", Optional botToken As String = "***YOUR_TELEGRAM_BOT_TOKEN***")
    ' Получаем размер файла в мегабайтах
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim fileSize As Double
        fileSize = fso.GetFile(fullAdress).Size / 1024 / 1024 ' Размер в мегабайтах
        Debug.Print "Размер файла: " & fileSize & " МБ"
          
     ' Проверка размера файла
    Const MAX_FILE_SIZE As Double = 50 ' Максимальный размер файла в МБ
    If fileSize > MAX_FILE_SIZE Then
        MsgBox "Файл слишком большой для отправки через Telegram. Максимальный размер: " & MAX_FILE_SIZE & " МБ", vbExclamation
        Exit Sub
    End If
          
    Const URL = "https://api.telegram.org/bot"
    Const METHOD_NAME = "/sendDocument?"
      
    Dim SheetName As String
    Dim Diapazon As String
    
    ' Извлечение названия листа
    SheetName = ExtractBetween(fullAdress, "лист(", ")")
    
    ' Извлечение диапазона
    Diapazon = ExtractBetween(fullAdress, "диапазон(", ")")


               
    Dim data As Object, key
    Set data = CreateObject("Scripting.Dictionary")
    data.Add "chat_id", CHAT_ID
    
          
    ' generate boundary
    Dim BOUNDARY, s As String, n As Integer
    For n = 1 To 16: s = s & Chr(65 + Int(Rnd * 25)): Next
    BOUNDARY = s & CDbl(Now)
  
    Dim part As String, ado As Object
    For Each key In data.keys
        part = part & "--" & BOUNDARY & vbCrLf
        part = part & "Content-Disposition: form-data; name=""" & key & """" & vbCrLf & vbCrLf
        part = part & data(key) & vbCrLf
    Next
    ' filename
    part = part & "--" & BOUNDARY & vbCrLf
    part = part & "Content-Disposition: form-data; name=""document""; filename=""" & Russian_utf(SheetName & "_" & Diapazon & ".jpg") & """" & vbCrLf & vbCrLf
                                                                    
    ' read file as binary
    Dim jpg
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 1 'binary
    ado.Open
    ado.LoadFromFile fullAdress
    ado.Position = 0
    jpg = ado.read
    ado.Close
  
    ' combine part, jpg , end
    ado.Open
    ado.Position = 0
    ado.Type = 1 ' binary
    ado.Write ToBytes(part)
    ado.Write jpg
    ado.Write ToBytes(vbCrLf & "--" & BOUNDARY & "---")
    ado.Position = 0
  
    Dim req As Object, reqURL As String
    Set req = CreateObject("MSXML2.XMLHTTP")
    reqURL = URL & botToken & METHOD_NAME
    With req
        .Open "POST", reqURL, False
        .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & BOUNDARY
        .send ado.read
        'MsgBox .responseText
    End With
  
End Sub
Sub sendTextToTelegram(Optional textMessage As String = "Привет!", Optional CHAT_ID As String = "***YOUR_TELEGRAM_CHAT_ID***", Optional botToken As String = "***YOUR_TELEGRAM_BOT_TOKEN***")
       
    textMessage = Russian_utf(textMessage)
    
    sURL = "https://api.telegram.org/bot" & botToken & "/sendMessage?chat_id=" & CHAT_ID & "&text=" & "%0D%0A" & textMessage
    Set oHttp = CreateObject("Msxml2.XMLHTTP")
    oHttp.Open "POST", sURL, False
    oHttp.send
    Set oHttp = Nothing
    
End Sub
Sub Выбрать_Имена_для_отправки_Листа_вТелеграм()
    UserForm2.Show
End Sub
Sub Сохранить_Лист()
    MsgBox "Лист сохранён по адресу: " & SaveActiveSheetAsValues
End Sub
Sub СохранитьЛистИОтправитьТелеграм(CHAT_ID As String, selectedName As String)
    Dim filePath As String
    
    ' Отправляем файл в Telegram
    Dim botToken As String
    botToken = "***YOUR_TELEGRAM_BOT_TOKEN***"
    Call SendSheetToTelegram(SaveActiveSheetAsValues, CHAT_ID, botToken)
End Sub
Function SaveActiveSheetAsValues()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim newBook As Workbook
    Dim filePath As String
    Dim FileName As String
    Dim OldFileName As String
    Dim OldSheetName As String
    
    Application.DisplayAlerts = False
    
    Set wb = ActiveWorkbook
    ' Устанавливаем текущий лист
    Set ws = ActiveSheet
    
    ' Создаем новый файл
    Set newBook = Workbooks.Add
    
    ' Копируем активный лист в новый файл
    ws.Copy Before:=newBook.Sheets(1)
    
    ' Удаляем все листы, кроме скопированного
    Application.DisplayAlerts = False
    While newBook.Sheets.Count > 1
        newBook.Sheets(2).Delete
    Wend
    Application.DisplayAlerts = True
    
    ' Сохраняем все значения в новый файл
    newBook.Sheets(1).Cells.Copy
    newBook.Sheets(1).Cells.PasteSpecial Paste:=xlPasteValues
    
    ' Определяем путь и имя файла
    filePath = wb.Path & "\" ' Укажите путь, куда сохранить файл
    
    OldFileName = wb.name
    OldSheetName = ws.name
    
    FileName = "файл(" & OldFileName & ")лист(" & OldSheetName & ")" & "дата(" & Format(Now, "yyyymmdd_hhmmss") & ")" & ".xlsx"
    
    Debug.Print FileName
    
    ' Сохраняем файл
    newBook.SaveAs FileName:=filePath & FileName, FileFormat:=xlOpenXMLWorkbook
    
    ' Закрываем новый файл
    newBook.Close SaveChanges:=False
    
    ' Возвращаем путь к сохраненному файлу
    SaveActiveSheetAsValues = filePath & FileName
End Function
Sub SendSheetToTelegram(filePath As String, Optional chatID As String = "***YOUR_TELEGRAM_CHAT_ID", Optional botToken As String = "***YOUR_TELEGRAM_BOT_TOKEN***")
    
    Dim URL As String
    Dim http As Object
    Dim BOUNDARY As String
    Dim body As String
    Dim fileData As String
    Dim fileSize As Long
    Dim fileContent As String
    
    Dim FileName As String
    Dim SheetName As String
    
    ' Извлечение названия файла
    FileName = ExtractBetween(filePath, "файл(", ")")
    ' Извлечение названия листа
    SheetName = ExtractBetween(filePath, "лист(", ")")
            
    
    ' URL для отправки файла
    URL = "https://api.telegram.org/bot" & botToken & "/sendDocument"
    
    ' Создаем объект XMLHTTP
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Определяем границу для multipart/form-data
    BOUNDARY = "---------------------------" & Format(Now, "yyyymmddhhmmss")
    
    ' Открываем файл и читаем его содержимое
    Open filePath For Binary Access Read As #1
    fileSize = LOF(1)
    fileContent = Space$(fileSize)
    Get #1, , fileContent
    Close #1
    
    ' Формируем тело запроса
    body = "--" & BOUNDARY & vbCrLf
    body = body & "Content-Disposition: form-data; name=""chat_id""" & vbCrLf & vbCrLf
    body = body & chatID & vbCrLf
    body = body & "--" & BOUNDARY & vbCrLf
    body = body & "Content-Disposition: form-data; name=""document""; filename=""" & Mid(filePath, InStrRev(filePath, "\") + 1) & """" & vbCrLf
    'body = body & "Content-Disposition: form-data; name=""document""; filename=""" & Russian_utf(FileName & "_" & SheetName & ".xlsx") & """" & vbCrLf
    
    body = body & "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" & vbCrLf & vbCrLf
    body = body & fileContent & vbCrLf
    body = body & "--" & BOUNDARY & "--" & vbCrLf
    
    ' Отправляем запрос
    http.Open "POST", URL, False
    http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & BOUNDARY
    http.send body
    
    ' Проверяем статус ответа
    If http.Status = 200 Then
        MsgBox "Файл успешно отправлен в Telegram!"
    Else
        MsgBox "Ошибка при отправке файла: " & http.Status & " " & http.responseText
    End If
End Sub
Sub ExtractData()
    Dim inputString As String
    Dim FileName As String
    Dim SheetName As String
    Dim Diapazon As String
    Dim data As String
    Dim Number As String
    
    ' Исходная строка
    inputString = "файл(БЮДЖЕТ 2025 - шаблон - ООТиЗ -29.10.2024.xlsx)лист(Индексация 24)диапазон(C2_D6)дата(02.11.2024_14.11.14)_3.jpg"
    
    ' Извлечение названия файла
    FileName = ExtractBetween(inputString, "файл(", ")")
    
    ' Извлечение названия листа
    SheetName = ExtractBetween(inputString, "лист(", ")")
    
    ' Извлечение диапазона
    Diapazon = ExtractBetween(inputString, "диапазон(", ")")
    
    ' Извлечение даты
    data = ExtractBetween(inputString, "дата(", ")")
    
    ' Извлечение цифры перед ".jpg"
    Number = ExtractNumberBeforeJpg(inputString)
    
    ' Вывод результатов в окно Immediate (Ctrl+G)
    Debug.Print "FileName: " & FileName
    Debug.Print "SheetName: " & SheetName
    Debug.Print "Diapazon: " & Diapazon
    Debug.Print "Data: " & data
    Debug.Print "Number: " & Number
End Sub

Function ExtractBetween(ByVal inputString As String, ByVal startString As String, ByVal endString As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
    ' Находим позицию начальной строки
    startPos = InStr(inputString, startString) + Len(startString)
    
    ' Находим позицию конечной строки
    endPos = InStr(startPos, inputString, endString)
    
    ' Извлекаем текст между начальной и конечной строками
    ExtractBetween = Mid(inputString, startPos, endPos - startPos)
End Function

Function ExtractNumberBeforeJpg(ByVal inputString As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
    ' Находим позицию ".jpg"
    endPos = InStrRev(inputString, ".jpg")
    
    ' Находим позицию цифры перед ".jpg"
    startPos = endPos - 1
    Do While IsNumeric(Mid(inputString, startPos, 1))
        startPos = startPos - 1
    Loop
    
    ' Извлекаем цифру
    ExtractNumberBeforeJpg = Mid(inputString, startPos + 1, endPos - startPos - 1)
End Function
Private Function ToBytes(str As String) As Variant
    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Open
    ado.Type = 2 ' text
    'ado.Charset = "_autodetect"
    ado.Charset = "UTF-8"
    ado.WriteText str
    ado.Position = 0
    ado.Type = 1
    ToBytes = ado.read
    ado.Close
  
End Function
Public Function Russian_utf(str)
    Static objHtmlfile As Object
    If objHtmlfile Is Nothing Then
        Set objHtmlfile = CreateObject("htmlfile")
        objHtmlfile.parentWindow.execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
    End If
    Russian_utf = objHtmlfile.parentWindow.encode(str)
End Function

Sub send_Document()

    Const URL = "https://api.telegram.org/bot"
    Const TOKEN = "***YOUR_TELEGRAM_BOT_TOKEN***"
    Const METHOD_NAME = "/sendDocument?"
    Const CHAT_ID = "***YOUR_TELEGRAM_CHAT_ID***"
    
    Const FOLDER = "C:\test\"
    Const DOCUMENT_FILE = "qqq.xlsx"
        
    Dim data As Object, key
    Set data = CreateObject("Scripting.Dictionary")
    data.Add "chat_id", CHAT_ID
    data.Add "caption", "dfgdfg"
    ' generate boundary
    Dim BOUNDARY, s As String, n As Integer
    For n = 1 To 16: s = s & Chr(65 + Int(Rnd * 25)): Next
    BOUNDARY = s & CDbl(Now)

    Dim part As String, ado As Object
    For Each key In data.keys
        part = part & "--" & BOUNDARY & vbCrLf
        part = part & "Content-Disposition: form-data; name=""" & key & """" & vbCrLf & vbCrLf
        part = part & data(key) & vbCrLf
    Next
    ' filename
    part = part & "--" & BOUNDARY & vbCrLf
    part = part & "Content-Disposition: form-data; name=""document""; filename=""" & DOCUMENT_FILE & """" & vbCrLf & vbCrLf
    
    ' read document file as binary
    Dim doc
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 1 'binary
    ado.Open
    ado.LoadFromFile FOLDER & DOCUMENT_FILE
    ado.Position = 0
    doc = ado.read
    ado.Close

    ' combine part, document, end
    ado.Open
    ado.Position = 0
    ado.Type = 1 ' binary
    ado.Write ToBytes(part)
    ado.Write doc
    ado.Write ToBytes(vbCrLf & "--" & BOUNDARY & "--")
    ado.Position = 0

    Dim req As Object, reqURL As String
    Set req = CreateObject("MSXML2.XMLHTTP")
    reqURL = URL & TOKEN & METHOD_NAME
    With req
        .Open "POST", reqURL, False
        .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & BOUNDARY
        .send ado.read
        MsgBox .responseText
    End With

End Sub

Sub SelectionToXML()

    Dim dataArray As Variant
    dataArray = Selection.Value

    Call SaveArrayToXML(dataArray)
    
End Sub

Sub SaveArrayToXML(dataArray As Variant)
       
    Dim cell As Range
    Dim xmlDoc As Object
    Dim rootNode As Object
    Dim rowNode As Object
    Dim cellNode As Object
    Dim filePath As String
    Dim i As Long
    Dim j As Long
    
    ' Получаем путь к текущему файлу Excel и добавляем имя XML файла
    filePath = ActiveWorkbook.Path & "\Data.xml"
    
    ' Проверяем существование файла
    fileExists = Dir(filePath) <> ""
    
    ' Создаем новый XML документ
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    
    ' Если файл существует, загружаем его
    If fileExists Then
        xmlDoc.async = False
        xmlDoc.Load filePath
        Set rootNode = xmlDoc.DocumentElement
    Else
        ' Если файл не существует, создаем новый XML документ и корневой элемент
        Set rootNode = xmlDoc.createElement("Data")
        xmlDoc.appendChild rootNode
    End If
    
    ' Проходим по каждой строке в диапазоне
    For i = 1 To UBound(dataArray, 1)
        ' Создаем узел для строки
        Set rowNode = xmlDoc.createElement("Row")
        rootNode.appendChild rowNode
        
        ' Проходим по каждой ячейке в строке
        For j = 1 To UBound(dataArray, 2)
            ' Создаем узел для ячейки
            Set cellNode = xmlDoc.createElement("Cell")
            cellNode.text = dataArray(i, j)
            rowNode.appendChild cellNode
        Next j
    Next i
    
    ' Сохраняем XML документ в файл
    xmlDoc.Save filePath
    
    ' Очищаем объекты
    Set xmlDoc = Nothing
    Set rootNode = Nothing
    Set rowNode = Nothing
    Set cellNode = Nothing
    
    MsgBox "XML файл успешно создан и сохранен по пути: " & filePath
        
End Sub
Sub LoadDataFromXML_to_Range()
    
    Set SpisokFiles = GetFilenamesCollection("Выбери XML файл", ThisWorkbook.Path)   'выводим окно выбора
    If SpisokFiles Is Nothing Then Exit Sub  'выход, если пользователь отказался от выбора файлов
        
    
    Dim fileAdresses() As String, fileNames() As String
    Dim i As Integer
    i = 1
    
    For Each Item In SpisokFiles
        ReDim Preserve fileAdresses(i - 1) As String
        fileAdresses(i - 1) = SpisokFiles.Item(i)
        
'        ReDim Preserve fileNames(i - 1) As String
'        fileNames(i - 1) = Mid(SpisokFiles.Item(i), InStrRev(SpisokFiles.Item(i), "\") + 1, Len(SpisokFiles.Item(i)) - InStrRev(SpisokFiles.Item(i), "\"))
        
        i = i + 1
    Next Item
    
    
    Dim xmlDoc As Object
    Dim rootNode As Object
    Dim rowNode As Object
    Dim cellNode As Object
    Dim filePath As String
    Dim ws As Worksheet
    Dim myCell As Range
    Dim rowIndex As Integer
    Dim colIndex As Integer
    
    ' Укажите путь к XML файлу
    'filePath = ActiveWorkbook.Path & "\Data.xml"
    filePath = fileAdresses(0)
    
    ' Создаем новый XML документ и загружаем файл
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.async = False
    xmlDoc.Load filePath
    
    ' Проверяем, успешно ли загружен файл
    If xmlDoc.parseError.ErrorCode <> 0 Then
        MsgBox "Ошибка при загрузке XML файла: " & xmlDoc.parseError.reason
        Exit Sub
    End If
    
    ' Получаем активный лист и активную ячейку
    Set ws = ActiveSheet
    Set myCell = ActiveCell
    
    ' Получаем корневой элемент
    Set rootNode = xmlDoc.DocumentElement
    
    ' Инициализируем индексы строк и столбцов
    rowIndex = 1
    colIndex = 1
    
    ' Проходим по каждому узлу Row
    For Each rowNode In rootNode.ChildNodes
        ' Проходим по каждому узлу Cell в текущем Row
        For Each cellNode In rowNode.ChildNodes
            ' Вставляем значение ячейки в активную ячейку и смещаемся вправо
            
                ws.Cells(myCell.row + rowIndex - 1, myCell.Column + colIndex - 1).Value = cellNode.text
            
            colIndex = colIndex + 1
        Next cellNode
        
        ' Сбрасываем индекс столбца и увеличиваем индекс строки
        colIndex = 1
        rowIndex = rowIndex + 1
    Next rowNode
    
    ' Очищаем объекты
    Set xmlDoc = Nothing
    Set rootNode = Nothing
    Set rowNode = Nothing
    Set cellNode = Nothing
    
    MsgBox "Данные успешно загружены из XML файла и вставлены в активный лист."
End Sub
