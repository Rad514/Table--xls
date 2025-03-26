Sub CreateTable()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) ' Выбираем первый лист

    ' Удаляем предыдущие данные (если есть)
    ws.Cells.Clear

    ' Заголовки таблицы
    ws.Cells(1, 1).Value = "Параметры"
    ws.Cells(1, 2).Value = "26.03"
    ws.Cells(1, 3).Value = "27.03"
    ws.Cells(1, 4).Value = "28.03"

    ' Строки таблицы
    ws.Cells(2, 1).Value = "Инструкция / План работ / Обучение ERP"
    ws.Cells(2, 2).Value = "Розлить три танка / Обучить мастера (Зенищев)"
    
    ws.Cells(3, 1).Value = "Факт"
    ws.Cells(3, 2).Value = "Розлили 4 танка (с 24 по 25 марта)"

    ' Форматируем таблицу
    With ws.Range("A1:D3")
        .Borders.LineStyle = xlContinuous ' Добавляем границы
        .Interior.Color = RGB(220, 230, 241) ' Цвет фона
        .Font.Bold = True ' Жирный шрифт для заголовков
    End With

    ' Авторазмер столбцов
    ws.Columns("A:D").AutoFit

    MsgBox "Таблица успешно создана!"
End Sub