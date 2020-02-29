Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports System.Threading
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Class Form1
    Dim thr As Thread
    Dim findName As String
    Dim findUsing As String

    'Public Sub doOne()
    '    Dim excelList As New List(Of String)
    '    Using GamePathf As New OpenFileDialog
    '        GamePathf.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
    '        GamePathf.Title = "請選擇檔案"
    '        GamePathf.Multiselect = True
    '        GamePathf.RestoreDirectory = True
    '        GamePathf.ShowDialog()
    '        excelList.AddRange(GamePathf.FileNames)

    '    End Using
    '    If excelList.Count = 0 Then
    '        Exit Sub
    '    End If

    '    Dim app As New Excel.Application
    '    Dim Book As Excel.Workbook

    '    app.DisplayAlerts = False
    '    app.Visible = False

    '    Dim isStart As Boolean = False


    '    Dim tmpBook As Excel.Workbook
    '    Dim tmpSheet As Excel.Worksheet
    '    tmpBook = app.Workbooks.Add

    '    Dim tmpID As Integer = 0
    '    For Each tmpPath In excelList
    '        tmpID += 1
    '        Try
    '            Book = app.Workbooks.Open(tmpPath, False, True)

    '            Dim data1, data2, data3, data4, data5, data6, data7, data8 As String
    '            data1 = ""
    '            data2 = ""
    '            data3 = ""
    '            data4 = ""
    '            data5 = ""
    '            data6 = ""
    '            data7 = ""
    '            data8 = ""
    '            '----------

    '            tmpSheet = tmpBook.Sheets(tmpID)
    '            tmpSheet = tmpBook.Worksheets.Add()
    '            tmpSheet.Name = Book.Name
    '            tmpSheet.Cells(1, 1).Value = "日期"
    '            tmpSheet.Cells(1, 2).Value = "廠商"
    '            tmpSheet.Cells(1, 3).Value = "品名"
    '            tmpSheet.Cells(1, 4).Value = "數量"
    '            tmpSheet.Cells(1, 5).Value = "單位"
    '            tmpSheet.Cells(1, 6).Value = "單價"
    '            tmpSheet.Cells(1, 7).Value = "金額"
    '            tmpSheet.Cells(1, 8).Value = "用途"

    '            Dim nowLastID As Integer = 2
    '            '----------

    '            For Each sheet As Excel.Worksheet In Book.Worksheets
    '                If Regex.IsMatch(sheet.Name.ToString, "0*([5-8][0-9]|9[0-9]|[1-8][0-9]{2}|9[0-8][0-9]|99[0-9])\.0*([1-9]|1[0-2])") Then

    '                    sheet.AutoFilterMode = False
    '                    setShow("進度：" & sheet.Name.ToString)
    '                    isStart = False

    '                    For Each row As Excel.Range In sheet.Rows
    '                        If isStart = True AndAlso row.Cells(1).value Is Nothing AndAlso row.Cells(2).value Is Nothing AndAlso row.Cells(3).value Is Nothing Then
    '                            Exit For '底部
    '                        End If

    '                        If isStart = False Then

    '                            If (row.Cells(1).value IsNot Nothing AndAlso row.Cells(1).Value.ToString.Contains("日期")) Then isStart = True
    '                            If (row.Cells(2).value IsNot Nothing AndAlso row.Cells(2).Value.ToString.Contains("廠商")) Then isStart = True
    '                            If (row.Cells(3).value IsNot Nothing AndAlso row.Cells(3).Value.ToString.Contains("品名")) Then isStart = True
    '                            If (row.Cells(4).value IsNot Nothing AndAlso row.Cells(4).Value.ToString.Contains("數量")) Then isStart = True
    '                            If (row.Cells(5).value IsNot Nothing AndAlso row.Cells(5).Value.ToString.Contains("單位")) Then isStart = True
    '                            If (row.Cells(6).value IsNot Nothing AndAlso row.Cells(6).Value.ToString.Contains("單價")) Then isStart = True
    '                            If (row.Cells(7).value IsNot Nothing AndAlso row.Cells(7).Value.ToString.Contains("金額")) Then isStart = True
    '                            If (row.Cells(8).value IsNot Nothing AndAlso row.Cells(8).Value.ToString.Contains("用途")) Then isStart = True

    '                        End If

    '                        If isStart = True Then
    '                            If row.Cells(1).value IsNot Nothing Then data1 = row.Cells(1).Value.ToString
    '                            If row.Cells(2).value IsNot Nothing Then data2 = row.Cells(2).Value.ToString
    '                            If row.Cells(8).value IsNot Nothing Then data8 = row.Cells(8).Value.ToString

    '                            If data2.Contains(findName) AndAlso (data8.Contains(findUsing) Or Len(findUsing) = 0) Then
    '                                data3 = ""
    '                                data4 = ""
    '                                data5 = ""
    '                                data6 = ""
    '                                data7 = ""
    '                                data8 = ""
    '                                If row.Cells(3).value IsNot Nothing Then data3 = row.Cells(3).Value.ToString
    '                                If row.Cells(4).value IsNot Nothing Then data4 = row.Cells(4).Value.ToString
    '                                If row.Cells(5).value IsNot Nothing Then data5 = row.Cells(5).Value.ToString
    '                                If row.Cells(6).value IsNot Nothing Then data6 = row.Cells(6).Value.ToString
    '                                If row.Cells(7).value IsNot Nothing Then data7 = row.Cells(7).Value.ToString
    '                                If row.Cells(8).value IsNot Nothing Then data8 = row.Cells(8).Value.ToString

    '                                '2013/1/5 上午 12:00:00

    '                                tmpSheet.Cells(nowLastID, 1).Value = data1.Split(" ")(0).Replace(data1.Split("/")(0), (Val(data1.Split("/")(0)) - 1911).ToString)
    '                                If Len(data1.Split(".")(0)) = 3 Then tmpSheet.Cells(nowLastID, 1).Value = data1
    '                                tmpSheet.Cells(nowLastID, 2).Value = data2
    '                                tmpSheet.Cells(nowLastID, 3).Value = data3
    '                                tmpSheet.Cells(nowLastID, 4).Value = data4
    '                                tmpSheet.Cells(nowLastID, 5).Value = data5
    '                                tmpSheet.Cells(nowLastID, 6).Value = data6
    '                                tmpSheet.Cells(nowLastID, 7).Value = data7
    '                                tmpSheet.Cells(nowLastID, 8).Value = data8

    '                                nowLastID += 1
    '                            End If

    '                        End If
    '                    Next

    '                End If
    '            Next

    '        Catch ex As Exception
    '            MsgBox("出錯")
    '            Exit Try
    '        End Try

    '    Next
    '    Using GamePathf As New SaveFileDialog
    '        GamePathf.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm"
    '        GamePathf.Title = "儲存"
    '        GamePathf.ShowDialog()
    '        tmpBook.SaveAs(GamePathf.FileName)
    '    End Using

    '    tmpBook.Close()

    '    Book.Close()
    '    app.Quit()
    '    bg_busy = False
    'End Sub
    Public Sub doOne()
        Dim excelList As New List(Of String)
        Using GamePathf As New OpenFileDialog
            GamePathf.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
            GamePathf.Title = "請選擇檔案"
            GamePathf.Multiselect = True
            GamePathf.RestoreDirectory = True
            GamePathf.ShowDialog()
            excelList.AddRange(GamePathf.FileNames)

        End Using
        If excelList.Count = 0 Then
            bg_busy = False
            Exit Sub
        End If

        Dim app As New Excel.Application
        Dim Book As Excel.Workbook

        app.DisplayAlerts = False
        app.Visible = False

        Dim isStart As Boolean = False


        Dim tmpBook As Excel.Workbook
        Dim tmpSheet As Excel.Worksheet
        tmpBook = app.Workbooks.Add

        Dim tmpID As Integer = 0
        For Each tmpPath In excelList
            tmpID += 1
            Try
                Book = app.Workbooks.Open(tmpPath, False, True)

                Dim data1, data2, data3, data4, data5, data6, data7, data8 As String
                data1 = ""
                data2 = ""
                data3 = ""
                data4 = ""
                data5 = ""
                data6 = ""
                data7 = ""
                data8 = ""
                '----------

                tmpSheet = tmpBook.Sheets(tmpID)
                tmpSheet = tmpBook.Worksheets.Add()
                tmpSheet.Name = Book.Name
                tmpSheet.Cells(1, 1).Value = "日期"
                tmpSheet.Cells(1, 2).Value = "廠商"
                tmpSheet.Cells(1, 3).Value = "品名"
                tmpSheet.Cells(1, 4).Value = "數量"
                tmpSheet.Cells(1, 5).Value = "單位"
                tmpSheet.Cells(1, 6).Value = "單價"
                tmpSheet.Cells(1, 7).Value = "金額"
                tmpSheet.Cells(1, 8).Value = "用途"
                tmpSheet.Columns("C").ColumnWidth = 50
                tmpSheet.Columns("H").ColumnWidth = 35

                Dim nowLastID As Integer = 2
                '----------

                For Each sheet As Excel.Worksheet In Book.Worksheets
                    If Regex.IsMatch(sheet.Name.ToString, "0*([5-8][0-9]|9[0-9]|[1-8][0-9]{2}|9[0-8][0-9]|99[0-9])\.0*([1-9]|1[0-2])") Then
                        sheet.AutoFilterMode = False

                        sheet.Range("B:B").NumberFormatLocal = "@"

                        setShow("進度：" & sheet.Name.ToString)
                        isStart = False


                        Dim findList As New List(Of Integer)

                        'If sheet.Range("B:B").Find(What:=findName) Is Nothing Then Continue For
                        ' MsgBox(sheet.Range("B:B").Find(What:=findName).Count)
                        'For Each RowTmp As Excel.Range In sheet.Range("B:B").Find(What:=findName)
                        '    findList.Add(RowTmp.Row)
                        '    'MsgBox(a.Row)
                        '    'MsgBox(sheet.Cells(a.Row, 2).value)
                        'Next

                        Dim RowTmp As Excel.Range
                        With sheet.Range("B:B")
                            RowTmp = .Find(What:=findName)
                            Do
                                If Not RowTmp Is Nothing Then
                                    If findList.Contains(RowTmp.Row) Then Exit Do
                                    findList.Add(RowTmp.Row)
                                    RowTmp = .FindNext(RowTmp)
                                Else
                                    Exit Do
                                End If
                            Loop
                        End With




                        For Each tmpRow As Integer In findList
                            Dim idOffest As Integer = 0
                            Do
                                data3 = ""
                                data4 = ""
                                data5 = ""
                                data6 = ""
                                data7 = ""
                                data8 = ""
                                If sheet.Cells(tmpRow + idOffest, 1).value IsNot Nothing Then data1 = sheet.Cells(tmpRow + idOffest, 1).value
                                If sheet.Cells(tmpRow + idOffest, 2).value IsNot Nothing Then data2 = sheet.Cells(tmpRow + idOffest, 2).value
                                If sheet.Cells(tmpRow + idOffest, 3).value IsNot Nothing Then data3 = sheet.Cells(tmpRow + idOffest, 3).value
                                If sheet.Cells(tmpRow + idOffest, 4).value IsNot Nothing Then data4 = sheet.Cells(tmpRow + idOffest, 4).value
                                If sheet.Cells(tmpRow + idOffest, 5).value IsNot Nothing Then data5 = sheet.Cells(tmpRow + idOffest, 5).value
                                If sheet.Cells(tmpRow + idOffest, 6).value IsNot Nothing Then data6 = sheet.Cells(tmpRow + idOffest, 6).value
                                If sheet.Cells(tmpRow + idOffest, 7).value IsNot Nothing Then data7 = sheet.Cells(tmpRow + idOffest, 7).value
                                If sheet.Cells(tmpRow + idOffest, 8).value IsNot Nothing Then data8 = sheet.Cells(tmpRow + idOffest, 8).value

                                If idOffest > 0 And findList.Contains(tmpRow + idOffest) Then
                                    Exit Do
                                End If



                                If Not data2.Contains(findName) Or Len(data3) = 0 Then Exit Do

                                If Len(findUsing) > 0 And Not data8.Contains(findUsing) Then
                                    idOffest += 1
                                    Continue Do
                                End If

                                Try
                                    tmpSheet.Cells(nowLastID, 1).Value = data1.Split(" ")(0).Replace(data1.Split("/")(0), (Val(data1.Split("/")(0)) - 1911).ToString)
                                    If Len(data1.Split(".")(0)) = 3 Then tmpSheet.Cells(nowLastID, 1).Value = data1
                                Catch ex As Exception
                                    tmpSheet.Cells(nowLastID, 1).Value = data1
                                End Try
                                tmpSheet.Cells(nowLastID, 2).Value = data2
                                tmpSheet.Cells(nowLastID, 3).Value = data3
                                tmpSheet.Cells(nowLastID, 4).Value = data4
                                tmpSheet.Cells(nowLastID, 5).Value = data5
                                tmpSheet.Cells(nowLastID, 6).Value = data6
                                tmpSheet.Cells(nowLastID, 7).Value = data7
                                tmpSheet.Cells(nowLastID, 8).Value = data8

                                idOffest += 1
                                nowLastID += 1
                            Loop
                        Next

                        findList.Clear()

                    End If
                Next

            Catch ex As Exception
                MsgBox("出錯")
                Exit Try
            End Try

        Next
        Using GamePathf As New SaveFileDialog
            GamePathf.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm"
            GamePathf.Title = "儲存"
            GamePathf.ShowDialog()
            tmpBook.SaveAs(GamePathf.FileName)
        End Using

        tmpBook.Close()

        Book.Close()
        app.Quit()
        bg_busy = False
    End Sub

    'Public Sub doOneA()
    '    Dim excelList As New List(Of String)
    '    Using GamePathf As New OpenFileDialog
    '        GamePathf.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
    '        GamePathf.Title = "請選擇檔案"
    '        GamePathf.Multiselect = True
    '        GamePathf.RestoreDirectory = True
    '        GamePathf.ShowDialog()
    '        excelList.AddRange(GamePathf.FileNames)

    '    End Using
    '    If excelList.Count = 0 Then
    '        Exit Sub
    '    End If

    '    Dim app As New Excel.Application
    '    Dim Book As Excel.Workbook

    '    app.DisplayAlerts = False
    '    app.Visible = False

    '    Dim isStart As Boolean = False


    '    Dim tmpBook As Excel.Workbook
    '    Dim tmpSheet As Excel.Worksheet
    '    tmpBook = app.Workbooks.Add

    '    Dim tmpID As Integer = 0
    '    For Each tmpPath In excelList
    '        tmpID += 1
    '        Try
    '            Book = app.Workbooks.Open(tmpPath, False, True)

    '            Dim data1, data2, data3, data4, data5, data6, data7, data8 As String
    '            data1 = ""
    '            data2 = ""
    '            data3 = ""
    '            data4 = ""
    '            data5 = ""
    '            data6 = ""
    '            data7 = ""
    '            data8 = ""
    '            '----------

    '            tmpSheet = tmpBook.Sheets(tmpID)
    '            tmpSheet = tmpBook.Worksheets.Add()
    '            tmpSheet.Name = Book.Name
    '            tmpSheet.Cells(1, 1).Value = "日期"
    '            tmpSheet.Cells(1, 2).Value = "廠商"
    '            tmpSheet.Cells(1, 3).Value = "品名"
    '            tmpSheet.Cells(1, 4).Value = "數量"
    '            tmpSheet.Cells(1, 5).Value = "單位"
    '            tmpSheet.Cells(1, 6).Value = "單價"
    '            tmpSheet.Cells(1, 7).Value = "金額"
    '            tmpSheet.Cells(1, 8).Value = "用途"

    '            Dim nowLastID As Integer = 2
    '            '----------

    '            For Each sheet As Excel.Worksheet In Book.Worksheets
    '                'If Regex.IsMatch(sheet.Name.ToString, "0*([5-8][0-9]|9[0-9]|[1-8][0-9]{2}|9[0-8][0-9]|99[0-9])\.0*([1-9]|1[0-2])") Then 
    '                If sheet.Cells.Find(What:="日期") IsNot Nothing Or sheet.Cells.Find(What:="廠商") IsNot Nothing Or sheet.Cells.Find(What:="品名") IsNot Nothing Or sheet.Cells.Find(What:="數量") IsNot Nothing Or sheet.Cells.Find(What:="單位") IsNot Nothing Or sheet.Cells.Find(What:="用途") IsNot Nothing Then

    '                    sheet.AutoFilterMode = False
    '                    setShow("進度：" & sheet.Name.ToString)
    '                    isStart = False

    '                    For Each row As Excel.Range In sheet.Rows
    '                        If isStart = True AndAlso row.Cells(1).value Is Nothing AndAlso row.Cells(2).value Is Nothing AndAlso row.Cells(3).value Is Nothing Then
    '                            Exit For '底部
    '                        End If

    '                        If isStart = False Then

    '                            If (row.Cells(1).value IsNot Nothing AndAlso row.Cells(1).Value.ToString.Contains("日期")) Then isStart = True
    '                            If (row.Cells(2).value IsNot Nothing AndAlso row.Cells(2).Value.ToString.Contains("廠商")) Then isStart = True
    '                            If (row.Cells(3).value IsNot Nothing AndAlso row.Cells(3).Value.ToString.Contains("品名")) Then isStart = True
    '                            If (row.Cells(4).value IsNot Nothing AndAlso row.Cells(4).Value.ToString.Contains("數量")) Then isStart = True
    '                            If (row.Cells(5).value IsNot Nothing AndAlso row.Cells(5).Value.ToString.Contains("單位")) Then isStart = True
    '                            If (row.Cells(6).value IsNot Nothing AndAlso row.Cells(6).Value.ToString.Contains("單價")) Then isStart = True
    '                            If (row.Cells(7).value IsNot Nothing AndAlso row.Cells(7).Value.ToString.Contains("金額")) Then isStart = True
    '                            If (row.Cells(8).value IsNot Nothing AndAlso row.Cells(8).Value.ToString.Contains("用途")) Then isStart = True

    '                        End If

    '                        If isStart = True Then
    '                            If row.Cells(1).value IsNot Nothing Then data1 = row.Cells(1).Value.ToString
    '                            If row.Cells(2).value IsNot Nothing Then data2 = row.Cells(2).Value.ToString
    '                            If row.Cells(8).value IsNot Nothing Then data8 = row.Cells(8).Value.ToString

    '                            If data2.Contains(findName) AndAlso (data8.Contains(findUsing) Or Len(findUsing) = 0) Then
    '                                data3 = ""
    '                                data4 = ""
    '                                data5 = ""
    '                                data6 = ""
    '                                data7 = ""
    '                                data8 = ""
    '                                If row.Cells(3).value IsNot Nothing Then data3 = row.Cells(3).Value.ToString
    '                                If row.Cells(4).value IsNot Nothing Then data4 = row.Cells(4).Value.ToString
    '                                If row.Cells(5).value IsNot Nothing Then data5 = row.Cells(5).Value.ToString
    '                                If row.Cells(6).value IsNot Nothing Then data6 = row.Cells(6).Value.ToString
    '                                If row.Cells(7).value IsNot Nothing Then data7 = row.Cells(7).Value.ToString
    '                                If row.Cells(8).value IsNot Nothing Then data8 = row.Cells(8).Value.ToString

    '                                '2013/1/5 上午 12:00:00

    '                                tmpSheet.Cells(nowLastID, 1).Value = data1.Split(" ")(0).Replace(data1.Split("/")(0), (Val(data1.Split("/")(0)) - 1911).ToString)
    '                                If Len(data1.Split(".")(0)) = 3 Then tmpSheet.Cells(nowLastID, 1).Value = data1
    '                                tmpSheet.Cells(nowLastID, 2).Value = data2
    '                                tmpSheet.Cells(nowLastID, 3).Value = data3
    '                                tmpSheet.Cells(nowLastID, 4).Value = data4
    '                                tmpSheet.Cells(nowLastID, 5).Value = data5
    '                                tmpSheet.Cells(nowLastID, 6).Value = data6
    '                                tmpSheet.Cells(nowLastID, 7).Value = data7
    '                                tmpSheet.Cells(nowLastID, 8).Value = data8

    '                                nowLastID += 1
    '                            End If

    '                        End If
    '                    Next

    '                End If
    '            Next

    '        Catch ex As Exception
    '            MsgBox("出錯")
    '            Exit Try
    '        End Try

    '    Next
    '    Using GamePathf As New SaveFileDialog
    '        GamePathf.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm"
    '        GamePathf.Title = "儲存"
    '        GamePathf.ShowDialog()
    '        tmpBook.SaveAs(GamePathf.FileName)
    '    End Using

    '    tmpBook.Close()

    '    Book.Close()
    '    app.Quit()
    '    bg_busy = False
    'End Sub
    Public Sub doOneA()
        Dim excelList As New List(Of String)
        Using GamePathf As New OpenFileDialog
            GamePathf.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
            GamePathf.Title = "請選擇檔案"
            GamePathf.Multiselect = True
            GamePathf.RestoreDirectory = True
            GamePathf.ShowDialog()
            excelList.AddRange(GamePathf.FileNames)

        End Using
        If excelList.Count = 0 Then
            bg_busy = False
            Exit Sub
        End If

        Dim app As New Excel.Application
        Dim Book As Excel.Workbook

        app.DisplayAlerts = False
        app.Visible = False

        Dim isStart As Boolean = False


        Dim tmpBook As Excel.Workbook
        Dim tmpSheet As Excel.Worksheet
        tmpBook = app.Workbooks.Add

        Dim tmpID As Integer = 0
        For Each tmpPath In excelList
            tmpID += 1
            Try
                Book = app.Workbooks.Open(tmpPath, False, True)

                Dim data1, data2, data3, data4, data5, data6, data7, data8 As String
                data1 = ""
                data2 = ""
                data3 = ""
                data4 = ""
                data5 = ""
                data6 = ""
                data7 = ""
                data8 = ""
                '----------

                tmpSheet = tmpBook.Sheets(tmpID)
                tmpSheet = tmpBook.Worksheets.Add()
                tmpSheet.Name = Book.Name
                tmpSheet.Cells(1, 1).Value = "日期"
                tmpSheet.Cells(1, 2).Value = "廠商"
                tmpSheet.Cells(1, 3).Value = "品名"
                tmpSheet.Cells(1, 4).Value = "數量"
                tmpSheet.Cells(1, 5).Value = "單位"
                tmpSheet.Cells(1, 6).Value = "單價"
                tmpSheet.Cells(1, 7).Value = "金額"
                tmpSheet.Cells(1, 8).Value = "用途"
                tmpSheet.Columns("C").ColumnWidth = 50
                tmpSheet.Columns("H").ColumnWidth = 35

                Dim nowLastID As Integer = 2
                '----------

                For Each sheet As Excel.Worksheet In Book.Worksheets
                    If 1 = 1 Then
                        sheet.AutoFilterMode = False

                        sheet.Range("B:B").NumberFormatLocal = "@"

                        setShow("進度：" & sheet.Name.ToString)
                        isStart = False


                        Dim findList As New List(Of Integer)

                        'If sheet.Range("B:B").Find(What:=findName) Is Nothing Then Continue For
                        ' MsgBox(sheet.Range("B:B").Find(What:=findName).Count)
                        'For Each RowTmp As Excel.Range In sheet.Range("B:B").Find(What:=findName)
                        '    findList.Add(RowTmp.Row)
                        '    'MsgBox(a.Row)
                        '    'MsgBox(sheet.Cells(a.Row, 2).value)
                        'Next

                        Dim RowTmp As Excel.Range
                        With sheet.Range("B:B")
                            RowTmp = .Find(What:=findName)
                            Do
                                If Not RowTmp Is Nothing Then
                                    If findList.Contains(RowTmp.Row) Then Exit Do
                                    findList.Add(RowTmp.Row)
                                    RowTmp = .FindNext(RowTmp)
                                Else
                                    Exit Do
                                End If
                            Loop
                        End With




                        For Each tmpRow As Integer In findList
                            Dim idOffest As Integer = 0
                            Do
                                data3 = ""
                                data4 = ""
                                data5 = ""
                                data6 = ""
                                data7 = ""
                                data8 = ""
                                If sheet.Cells(tmpRow + idOffest, 1).value IsNot Nothing Then data1 = sheet.Cells(tmpRow + idOffest, 1).value
                                If sheet.Cells(tmpRow + idOffest, 2).value IsNot Nothing Then data2 = sheet.Cells(tmpRow + idOffest, 2).value
                                If sheet.Cells(tmpRow + idOffest, 3).value IsNot Nothing Then data3 = sheet.Cells(tmpRow + idOffest, 3).value
                                If sheet.Cells(tmpRow + idOffest, 4).value IsNot Nothing Then data4 = sheet.Cells(tmpRow + idOffest, 4).value
                                If sheet.Cells(tmpRow + idOffest, 5).value IsNot Nothing Then data5 = sheet.Cells(tmpRow + idOffest, 5).value
                                If sheet.Cells(tmpRow + idOffest, 6).value IsNot Nothing Then data6 = sheet.Cells(tmpRow + idOffest, 6).value
                                If sheet.Cells(tmpRow + idOffest, 7).value IsNot Nothing Then data7 = sheet.Cells(tmpRow + idOffest, 7).value
                                If sheet.Cells(tmpRow + idOffest, 8).value IsNot Nothing Then data8 = sheet.Cells(tmpRow + idOffest, 8).value

                                If idOffest > 0 And findList.Contains(tmpRow + idOffest) Then
                                    Exit Do
                                End If



                                If Not data2.Contains(findName) Or Len(data3) = 0 Then Exit Do

                                If Len(findUsing) > 0 And Not data8.Contains(findUsing) Then
                                    idOffest += 1
                                    Continue Do
                                End If

                                Try
                                    tmpSheet.Cells(nowLastID, 1).Value = data1.Split(" ")(0).Replace(data1.Split("/")(0), (Val(data1.Split("/")(0)) - 1911).ToString)
                                    If Len(data1.Split(".")(0)) = 3 Then tmpSheet.Cells(nowLastID, 1).Value = data1
                                Catch ex As Exception
                                    tmpSheet.Cells(nowLastID, 1).Value = data1
                                End Try
                                tmpSheet.Cells(nowLastID, 2).Value = data2
                                tmpSheet.Cells(nowLastID, 3).Value = data3
                                tmpSheet.Cells(nowLastID, 4).Value = data4
                                tmpSheet.Cells(nowLastID, 5).Value = data5
                                tmpSheet.Cells(nowLastID, 6).Value = data6
                                tmpSheet.Cells(nowLastID, 7).Value = data7
                                tmpSheet.Cells(nowLastID, 8).Value = data8

                                idOffest += 1
                                nowLastID += 1
                            Loop
                        Next

                        findList.Clear()

                    End If
                Next

            Catch ex As Exception
                MsgBox("出錯")
                Exit Try
            End Try

        Next
        Using GamePathf As New SaveFileDialog
            GamePathf.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm"
            GamePathf.Title = "儲存"
            GamePathf.ShowDialog()
            tmpBook.SaveAs(GamePathf.FileName)
        End Using

        tmpBook.Close()

        Book.Close()
        app.Quit()
        bg_busy = False
    End Sub

    Dim bg_busy As Boolean = False

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Button1.Enabled = False
        If thr Is Nothing Then thr = New Thread(New ThreadStart(AddressOf doOne))
        If bg_busy = False Then
            findName = TextBox1.Text
            findUsing = TextBox2.Text
            thr.Abort()
            thr = New Thread(New ThreadStart(AddressOf doOne))
            thr.ApartmentState = System.Threading.ApartmentState.STA
            thr.Start()
            bg_busy = True
        End If
        Timer1.Enabled = True
    End Sub

    Public Sub setShow(ByVal MyText As String)
        If Me.InvokeRequired() Then
            Dim cb As New updateX(AddressOf setShow)
            Me.Invoke(cb, MyText)
        Else
            Label2.Text = MyText
        End If
    End Sub

    Public Sub setShowT(ByVal MyText As String)
        If Me.InvokeRequired() Then
            Dim cb As New updateX(AddressOf setShowT)
            Me.Invoke(cb, MyText)
        Else
            TextBox3.Text = MyText
        End If
    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If thr IsNot Nothing Then thr.Abort()
        Environment.Exit(Environment.ExitCode)
    End Sub

    Private Delegate Sub updateX(ByVal MyText As String)

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If bg_busy = False Then
            Button1.Enabled = True
            Button2.Enabled = True
            Button3.Enabled = True
            Timer1.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Button2.Enabled = False
        If thr Is Nothing Then thr = New Thread(New ThreadStart(AddressOf serchStr))
        If bg_busy = False Then
            findName = TextBox1.Text
            thr.Abort()
            thr = New Thread(New ThreadStart(AddressOf serchStr))
            thr.ApartmentState = System.Threading.ApartmentState.STA
            thr.Start()
            bg_busy = True
        End If
        Timer1.Enabled = True


    End Sub

    Function serchStr()
        bg_busy = True
        Dim excelList As New List(Of String)
        Using GamePathf As New OpenFileDialog
            GamePathf.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
            GamePathf.Title = "請選擇檔案"
            GamePathf.Multiselect = True
            GamePathf.RestoreDirectory = True
            GamePathf.ShowDialog()
            excelList.AddRange(GamePathf.FileNames)

        End Using

        Dim findstr As String
        findstr = findName

        If excelList.Count = 0 Or Len(findstr) = 0 Then
            bg_busy = False
            Exit Function
        End If

        Dim app As New Excel.Application
        Dim Book As Excel.Workbook

        app.DisplayAlerts = False
        app.Visible = False

        Dim isFind As String = ""

        For Each tmpPath In excelList
            Book = app.Workbooks.Open(tmpPath, False, True)

            For Each sheet As Excel.Worksheet In Book.Worksheets
                setShow("進度：" & sheet.Name.ToString)
                sheet.AutoFilterMode = False

                If sheet.Cells.Find(What:=findstr) IsNot Nothing Then
                    isFind &= Book.Name & " __ " & sheet.Name & vbCrLf
                End If

            Next
        Next


        Book.Close()
        app.Quit()
        bg_busy = False
        setShowT(isFind)
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Button3.Enabled = False
        If thr Is Nothing Then thr = New Thread(New ThreadStart(AddressOf doOne))
        If bg_busy = False Then
            findName = TextBox1.Text
            findUsing = TextBox2.Text
            thr.Abort()
            thr = New Thread(New ThreadStart(AddressOf doOneA))
            thr.ApartmentState = System.Threading.ApartmentState.STA
            thr.Start()
            bg_busy = True
        End If
        Timer1.Enabled = True
    End Sub
End Class
