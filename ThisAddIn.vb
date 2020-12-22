
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports System.Data
Imports System.Text.RegularExpressions
Imports System.Collections
Imports System.IO
Public Class ThisAddIn
    ''' <summary>
    ''' 全局字典，从数据库读取存入
    ''' </summary>
    Public Shared PubDic As Dictionary(Of String, String)
    ''' <summary>
    ''' 合表
    ''' </summary>
    Sub CombinSheet()

        Dim col As String = "b"

        Try
            Select Case InputBox("按序号选择需要生成的表类型：" & vbCrLf &
                                           "1、标准件表" & vbCrLf &
                                           "2、免拼合计" & vbCrLf &
                                           "3、销售清单" & vbCrLf &
                                           "4、铁件销售")

                Case 1 : Exl.GetSheetType(Exl.DataSet.Tables("表信息").Select("headerName='" & "标准板明细清单" & "'")(0)) : col = "c"
                Case 2 : Exl.GetSheetType(Exl.DataSet.Tables("表信息").Select("headerName='" & "打包清单" & "'")(0))
                Case 3 : Exl.GetSheetType(Exl.DataSet.Tables("表信息").Select("headerName='" & "铝件销售明细清单" & "'")(0)) : col = "c"
                Case 4 : Exl.GetSheetType(Exl.DataSet.Tables("表信息").Select("headerName='" & "铁件销售明细清单" & "'")(0)) : col = "c"

                Case Else : Exit Try

            End Select

            Exl.InsertTH()

            Dim star As Integer = Exl.StaRow()

            Dim rep As Integer = 0

            Dim wsheetactive As Worksheet = Exl.Sheet

            Dim di As New OpenFileDialog()
            di.Multiselect = True
            di.Title = "选择需要合并的文件"

            If di.ShowDialog = System.Windows.Forms.DialogResult.OK Then

                Dim Flist As String() = di.FileNames
                Dim fPath As String = Path.GetDirectoryName(Flist(0))

                Dim SelRange As New Dictionary(Of String, LinkedList(Of Range))

                Dim wbook As Workbook

                Dim projectname As String = ""

                For Each fInfo As String In Flist

                    wbook = Globals.ThisAddIn.Application.Workbooks.Open(fInfo)

                    wbook.Windows.Item(1).Visible = True

                    For Each Sh As Worksheet In wbook.Sheets

                        Exl.Sheet = Sh

                        Exl.Sheet.Range("a" & Exl.StaRow & ":" & ChrW(Exl.ColNum + 64) & Exl.EndRow(col)).Copy()
                        wsheetactive.Range("a" & star).PasteSpecial()
                        Exl.Sheet = wsheetactive
                        star = Exl.EndRow + 1

                    Next

                    wbook.Windows.Item(1).Visible = True
                    wbook.Close()
                Next

                Exl.Sheet = wsheetactive

                SerialCombin()

                For index = Exl.StaRow To Exl.EndRow

                    If Exl.ColNum = 10 Then
                        Exl.Sheet.Range("i" & index).Value = "=g" & index & "*h" & index
                    ElseIf Exl.ColNum = 9 Then
                        Exl.Sheet.Range("h" & index).Value = "=f" & index & "*g" & index & "/1000"
                    End If

                Next

                Exl.SumTo()

                Exl.SheetFormat()

                MessageBox.Show("计算完成")

            End If

        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try

    End Sub

    Sub CombinSheet11()

        Dim col As String = "b"

        Try

            Dim ss As Dictionary(Of String, Dictionary(Of String, String)) = New Dictionary(Of String, Dictionary(Of String, String))

            Dim wsheetactive As Worksheet = Exl.WorkBook.ActiveSheet
            Dim star As Integer = wsheetactive.Range("h65536").End(XlDirection.xlUp).Row

            Dim di As New OpenFileDialog()
            di.Multiselect = True
            di.Title = "选择需要合并的文件"

            If di.ShowDialog = System.Windows.Forms.DialogResult.OK Then

                Dim Flist As String() = di.FileNames
                Dim fPath As String = Path.GetDirectoryName(Flist(0))

                Dim SelRange As New Dictionary(Of String, LinkedList(Of Range))

                Dim wbook As Workbook

                Dim projectname As String = ""

                For Each fInfo As String In Flist

                    wbook = Globals.ThisAddIn.Application.Workbooks.Open(fInfo)

                    wbook.Windows.Item(1).Visible = True

                    For Each Sh As Worksheet In wbook.Sheets

                        Dim wsheetactive11 As Worksheet = Sh
                        Dim star1 As Integer = wsheetactive11.Range("h65536").End(XlDirection.xlUp).Row

                        For index = 3 To star1
                            Dim sss As Dictionary(Of String, String) = New Dictionary(Of String, String)
                            For index1 = 1 To 12

                                Dim col11 As String = "A" & ChrW(index1 + 65)

                                If Not IsNothing(Sh.Range(col11 & index).Value) AndAlso Sh.Range(col11 & index).Value.ToString <> "" AndAlso
                                   Not IsNothing(Sh.Range(col11 & 2).Value) AndAlso Sh.Range(col11 & 2).Value.ToString <> "" Then

                                    sss.Add(Sh.Range(col11 & 2).Value, Sh.Range(col11 & index).Value.ToString)

                                End If

                            Next

                            If Not ss.ContainsKey(Sh.Range("f" & index).Value.ToString.Replace(" ", "") &
                                                  Sh.Range("h" & index).Value.ToString.Replace(" ", "") &
                                                  Sh.Range("i" & index).Value.ToString.Replace(" ", "")) AndAlso sss.Count > 0 Then
                                ss.Add(Sh.Range("h" & index).Value.ToString.Replace(" ", "") & Sh.Range("i" & index).Value.ToString.Replace(" ", ""), sss)
                            End If

                        Next

                    Next

                    wbook.Windows.Item(1).Visible = True

                    wbook.Close()
                Next

                For index = 3 To star

                    Dim key As String = wsheetactive.Range("f" & index).Value.ToString.Replace(" ", "") &
                        wsheetactive.Range("h" & index).Value.ToString.Replace(" ", "") &
                        wsheetactive.Range("i" & index).Value.ToString.Replace(" ", "")

                    If ss.ContainsKey(key) Then

                        For index1 = 1 To 12

                            Dim col11 As String = "A" & ChrW(index1 + 65)

                            If ss(key).ContainsKey(wsheetactive.Range(col11 & 2).Value) Then

                                wsheetactive.Range(col11 & index).Value = ss(key)(wsheetactive.Range(col11 & 2).Value)

                            End If

                        Next

                    End If

                Next

                MessageBox.Show("计算完成")

            End If

        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try

    End Sub
    ''' <summary>
    ''' 分表
    ''' </summary>
    Sub Fen()
        Try
            Dim path As String = Exl.WorkBook.Path

            Dim SheetDict As New Dictionary(Of String, Dictionary(Of String, Object)) : Dim pix As String = ""

            For i As Integer = Exl.StaRow To Exl.EndRow

                pix = Exl.Sheet.Range("a" & i).Value.Substring(0, Exl.Sheet.Range("a" & i).Value.IndexOf("-"))

                If Not SheetDict.ContainsKey(pix) Then

                    SheetDict.Add(pix, New Dictionary(Of String, Object))
                    SheetDict(pix).Add(Exl.Sheet.Range("a" & i).Value.Substring(Exl.Sheet.Range("a" & i).Value.IndexOf("-") + 1), Exl.Sheet.Range("b" & i).Value)

                Else

                    If Not SheetDict(pix).ContainsKey(Exl.Sheet.Range("a" & i).Value.Substring(Exl.Sheet.Range("a" & i).Value.IndexOf("-") + 1)) Then
                        SheetDict(pix).Add(Exl.Sheet.Range("a" & i).Value.Substring(Exl.Sheet.Range("a" & i).Value.IndexOf("-") + 1), Exl.Sheet.Range("b" & i).Value)
                    Else
                        SheetDict(pix)(Exl.Sheet.Range("a" & i).Value.Substring(Exl.Sheet.Range("a" & i).Value.IndexOf("-") + 1)) = SheetDict(pix)(Exl.Sheet.Range("a" & i).Value.Substring(Exl.Sheet.Range("a" & i).Value.IndexOf("-") + 1)) + Exl.Sheet.Range("b" & i).Value
                    End If
                End If
            Next

            Dim wb As Workbook : Dim index As Integer = 2

            For Each S As String In SheetDict.Keys

                wb = Exl.WorkBook.Application.Workbooks.Add()

                Exl.Sheet1 = wb.ActiveSheet

                Exl.Sheet.Range("a1").Value = "编号" : Exl.Sheet.Range("b1").Value = "数量"

                index = 2

                For Each se As String In SheetDict(S).Keys
                    Exl.Sheet.Range("a" & index).Value = se : Exl.Sheet.Range("b" & index).Value = SheetDict(S)(se)
                    index = index + 1
                Next

                wb.SaveCopyAs(path & "\" & S & ".xlsx")

                wb.Close()

            Next
        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try
    End Sub
    ''' <summary>
    ''' 编码合并
    ''' </summary>
    Public Sub SerialCombin(Optional ByVal keyCol_2 As String = "")
        Try
            If Exl.ColNum < 3 Then Exit Sub

            Dim keyCol, colN As String

            Select Case Exl.ColNum
                Case 3
                    keyCol = "b" : colN = "c"
                Case 9, 10
                    keyCol = "c" : colN = Exl.NumValue
                Case Else
                    Exit Sub
            End Select

            Exl.SortCol(keyCol, keyCol_2)

            Dim row As Integer = Exl.EndRow - Exl.StaRow

            Dim i As Integer

            While (i < row)

                If Exl.Sheet.Range(keyCol & i + Exl.StaRow).Value = Exl.Sheet.Range(keyCol & i + Exl.StaRow + 1).Value Then

                    Exl.Sheet.Range(colN & i + Exl.StaRow).Value = Exl.Sheet.Range(colN & i + Exl.StaRow).Value + Exl.Sheet.Range(colN & i + Exl.StaRow + 1).Value

                    Exl.Sheet.Range(colN & i + Exl.StaRow + 1).EntireRow.Delete()

                    row = row - 1 : Continue While

                End If

                i = i + 1

            End While

            Exl.SumTo()

            Exl.SheetFormat(keyCol)

            ' Exl.SetPrintFormat()
        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 生成生产清单
    ''' </summary>
    Sub CreatePBom()
        Try
            Dim SelRange As Range = GetSelRange()
            Dim Flist As New List(Of IO.FileInfo)
            Dim producSer As ProductionSer
            Exl.FliePath = Exl.WorkBook.Path
            ProductionSer.ClearPix()
            Method.wbookDic.Clear() : Method.psDic.Clear() : Method.serDic.Clear()

            If Not IsNothing(SelRange) Then

                Exl.ProjectName = Exl.WorkBook.Name.Substring(0, Exl.WorkBook.Name.IndexOf("."))

                For i As Short = 1 To SelRange.Rows.Count

                    If Not IsNothing(SelRange.Range("a" & i).Value) AndAlso Not IsNothing(SelRange.Range("b" & i).Value) Then

                        producSer = New ProductionSer(SelRange.Range("a" & i).Value, SelRange.Range("b" & i).Value)

                        If Not Method.serDic.ContainsKey(producSer.GetSer) Then producSer.SerialCal() : Method.serDic.Add(producSer.GetSer, producSer.SerDic)

                        Method.DicAdd(producSer)

                    Else
                        Continue For
                    End If

                Next

            ElseIf Not Method.SelFile(Flist, Exl.ProjectName, Exl.FliePath, ThisAddIn.PubDic("分表名")) Then

                Exit Sub

            Else

                Dim wbook As Workbook = Nothing
                Dim SelRangeFile As Range = Nothing

                For Each fInfo As FileInfo In Flist

                    wbook = Globals.ThisAddIn.Application.Workbooks.Open(fInfo.FullName)

                    SelRangeFile = GetSelRange(wbook.ActiveSheet)

                    If IsNothing(SelRangeFile) Then Continue For

                    ProductionSer.GetPix(fInfo.Name.Substring(0, fInfo.Name.LastIndexOf(".")))

                    For i As Short = 1 To SelRangeFile.Rows.Count

                        If Not IsNothing(SelRangeFile.Range("a" & i).Value) AndAlso Not IsNothing(SelRangeFile.Range("b" & i).Value) Then

                            producSer = New ProductionSer(SelRangeFile.Range("a" & i).Value, SelRangeFile.Range("b" & i).Value)

                            If Not Method.serDic.ContainsKey(producSer.GetSer) Then producSer.SerialCal() : Method.serDic.Add(producSer.GetSer, producSer.SerDic)

                            Method.DicAdd(producSer)

                        Else
                            Continue For
                        End If

                    Next

                    wbook.Close()

                Next

            End If

            Method.Eval()
            MessageBox.Show("计算完成")

        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try

    End Sub

    ''' <summary>
    ''' 打包清单
    ''' </summary>
    Sub PackBom()
        Try
            Dim SelRange As Range = GetSelRange()

            Dim Flist As New List(Of IO.FileInfo)

            Dim producSer As ProductionSer

            ProductionSer.ClearPix()

            Exl.FliePath = Exl.WorkBook.Path

            If Not IsNothing(SelRange) Then

                Dim drType As DataRow() = Nothing

                Dim dDir As New Dictionary(Of String, Range)

                Dim shName As String

                For i As Short = 1 To SelRange.Rows.Count

                    If Not IsNothing(SelRange.Range("a" & i).Value) AndAlso Not IsNothing(SelRange.Range("b" & i).Value) Then

                        producSer = New ProductionSer(SelRange.Range("a" & i).Value)

                        If producSer.GetSerType.Contains("铁") Then shName = PubDic("铁打包") Else shName = PubDic("铝打包")

                        If Not dDir.ContainsKey(shName) Then

                            dDir.Add(shName, SelRange.Range("a" & i & ": " & "b" & i))

                        Else

                            dDir(shName) = Application.Union(dDir(shName), SelRange.Range("a" & i & ":" & "b" & i))

                        End If

                    End If

                Next

                For Each sh As String In dDir.Keys

                    drType = Exl.DataSet.Tables("表信息").Select("编码类型" & "='" & sh & "'") ： If drType.Length = 0 Then Continue For

                    Exl.Sheet1 = Exl.WorkBook.Sheets.Add()

                    Exl.Sheet.Name = sh

                    Exl.GetSheetType(drType(0))

                    Exl.ProjectName = Exl.WorkBook.Name.Substring(0, Exl.WorkBook.Name.IndexOf("."))

                    Exl.InsertTH()

                    dDir(sh).Copy(Exl.Sheet.Range("b" & Exl.StaRow))

                    SerialCombin()

                Next

            ElseIf Not Method.SelFile(Flist, Exl.ProjectName, Exl.FliePath, PubDic("分表名")) Then

                Exit Sub

            Else

                Method.DwbookDic.Clear() : Method.dpDic.Clear()

                For Each fInfo As IO.FileInfo In Flist

                    Dim wbook As Workbook = Globals.ThisAddIn.Application.Workbooks.Open(fInfo.FullName) : Dim SelRangeFile As Range = GetSelRange(wbook.ActiveSheet)

                    If IsNothing(SelRangeFile) Then Continue For

                    Dim p As String = fInfo.Name.Substring(0, fInfo.Name.LastIndexOf(".")) : ProductionSer.GetPix(p)

                    For i As Short = 1 To SelRangeFile.Rows.Count

                        If Not IsNothing(SelRangeFile.Range("a" & i).Value) AndAlso Not IsNothing(SelRangeFile.Range("b" & i).Value) Then

                            producSer = New ProductionSer(SelRangeFile.Range("a" & i).Value, SelRangeFile.Range("b" & i).Value)

                            Method.DDicAdd(producSer, p)

                        Else
                            Continue For
                        End If

                    Next

                    wbook.Close()

                Next

                Method.DEval()

            End If
            MessageBox.Show("计算完成")
        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 二维清单
    ''' </summary>
    Sub AllBom()

        Try

            Dim SelRange As Range = GetSelRange() '获得单元格内容

            Dim pSet As New HashSet(Of String)

            Dim pix, ser As String

            Dim producSer As ProductionSer

            If IsNothing(SelRange) Then Exit Sub

            Exl.FliePath = Exl.WorkBook.Path

            Exl.ProjectName = Exl.WorkBook.Name.Substring(0, Exl.WorkBook.Name.IndexOf("."))

            Method.wbookDic.Clear() : Method.DwbookDic.Clear() : Method.psDic.Clear() : Method.dpDic.Clear() : Method.serDic.Clear() : Method.HillDic.Clear()

            For i As Short = 1 To SelRange.Rows.Count

                If Not IsNothing(SelRange.Range("a" & i).Value) AndAlso Not IsNothing(SelRange.Range("b" & i).Value) Then

                    pix = SelRange.Range("a" & i).Value.Substring(0, SelRange.Range("a" & i).Value.IndexOf("-"))

                    If pix.Contains("(") Then
                        pix = SelRange.Range("a" & i).Value.Substring(0, SelRange.Range("a" & i).Value.IndexOf(")") + 1)
                        ser = SelRange.Range("a" & i).Value.Substring(SelRange.Range("a" & i).Value.IndexOf(")") + 2)
                    Else
                        ser = SelRange.Range("a" & i).Value.Substring(SelRange.Range("a" & i).Value.IndexOf("-") + 1)
                    End If

                    ProductionSer.GetPix(pix)

                    If ProductionSer.PixCode = "" Then Exit For

                    producSer = New ProductionSer(ser, SelRange.Range("b" & i).Value)

                    If Not Method.serDic.ContainsKey(producSer.GetSer) Then producSer.SerialCal() : Method.serDic.Add(producSer.GetSer, producSer.SerDic)

                    Method.DicAdd(producSer) ' : Method.BCDicAdd(producSer) ： Method.BDicAdd(producSer)

                    Method.DDicAdd(producSer, pix) : Method.BDDicAdd(producSer)
                    ' Method.BDDicAdd(producSer)

                End If

            Next

            Method.Eval() '生成清单生产及单元格赋值
            Method.DEval() '打包清单生产及单元格赋值

            MessageBox.Show("计算完成")

        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try

    End Sub

    ''' <summary>
    ''' 二维清单
    ''' </summary>
    Sub AllBom2()

        ' Try

        Dim SelRange As Range = GetSelRange() '获得单元格内容

            Dim pSet As New HashSet(Of String)

            Dim pix, ser As String

            Dim producSer As ProductionSer

            If IsNothing(SelRange) Then Exit Sub

            Exl.FliePath = Exl.WorkBook.Path

            Exl.ProjectName = Exl.WorkBook.Name.Substring(0, Exl.WorkBook.Name.IndexOf("."))

            Method.wbookDic.Clear() : Method.DwbookDic.Clear() : Method.psDic.Clear() : Method.dpDic.Clear() : Method.serDic.Clear() : Method.HillDic.Clear()

            For i As Short = 1 To SelRange.Rows.Count

                If Not IsNothing(SelRange.Range("a" & i).Value) AndAlso Not IsNothing(SelRange.Range("b" & i).Value) Then

                    pix = SelRange.Range("a" & i).Value.Substring(0, SelRange.Range("a" & i).Value.IndexOf("-"))

                    If pix.Contains("(") Then
                        pix = SelRange.Range("a" & i).Value.Substring(0, SelRange.Range("a" & i).Value.IndexOf(")") + 1)
                        ser = SelRange.Range("a" & i).Value.Substring(SelRange.Range("a" & i).Value.IndexOf(")") + 2)
                    Else
                        ser = SelRange.Range("a" & i).Value.Substring(SelRange.Range("a" & i).Value.IndexOf("-") + 1)
                    End If

                    ProductionSer.GetPix(pix)

                    If ProductionSer.PixCode = "" Then Exit For

                    producSer = New ProductionSer(ser, SelRange.Range("b" & i).Value)

                    If Not Method.serDic.ContainsKey(producSer.GetSer) Then producSer.SerialCal() : Method.serDic.Add(producSer.GetSer, producSer.SerDic)

                    Method.DicAdd2(producSer)

                End If

            Next

            Method.Eval() '生成清单生产及单元格赋值

            MessageBox.Show("计算完成")

        '  Catch ex As Exception
        ' Method.ExceptionWrite(ex)
        '  End Try

    End Sub

    Sub SellBom1()
        Try

            Dim wbook As Workbook

            Dim FList1 As IO.FileInfo()

            Dim FolderDialog As New FolderBrowserDialog With {.Description = "选择清单所在的文件夹"}

            If DialogResult.OK = FolderDialog.ShowDialog Then

                Dim Path As String = FolderDialog.SelectedPath

                Dim Di As DirectoryInfo = New DirectoryInfo(Path)

                FList1 = Di.GetFiles("*.xl*")

            End If

            If IsNothing(FList1) Then Exit Sub

            For Each fInfo As IO.FileInfo In FList1

                wbook = Globals.ThisAddIn.Application.Workbooks.Open(fInfo.FullName)

                wbook.Windows.Item(1).Visible = True

                For Each Sh As Worksheet In wbook.Sheets

                    Exl.Sheet = Sh

                    If Exl.StaRow = 9 OrElse Exl.StaRow = 7 Then

                        For i As Integer = Exl.StaRow To Exl.EndRow

                            If Exl.Sheet.Range("c" & i).Value = "3C2100-600YDR-100" Then

                                Dim s As String = Exl.Sheet.Range("c" & i).Value
                                Dim ss As String = fInfo.Name

                                MsgBox(fInfo.Name & "+" & i)

                            End If

                        Next

                    End If

                Next

                wbook.Close()

            Next

            MessageBox.Show("计算完成")

        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try
    End Sub
    ''' <summary>
    ''' 销售清单
    ''' </summary>
    Sub SellBom()
        Try
            Dim Flist As New List(Of IO.FileInfo)

            Dim SelRange As New Dictionary(Of String, LinkedList(Of Range))

            Dim wbook As Workbook ： Dim key As String

            Dim w1 As Workbook = Application.Workbooks.Add : Dim shmin As Worksheet = w1.Sheets.Item(1)

            Exl.Sheet1 = shmin

            Exl.GetSheetType(Exl.DataSet.Tables("表信息").Select("编码类型='" & "铝制件销售清单" & "'")(0))

            Exl.InsertTH()

            Dim ro As Integer = Exl.StaRow

            If Not Method.SelFile(Flist, Exl.ProjectName, Exl.FliePath, "") Then Exit Sub

            Dim co As Integer = 0
            Dim na As String
            For Each fInfo As IO.FileInfo In Flist
                na = fInfo.Name
                co = co + 1
                '  If fInfo.Name.Contains("打包") OrElse fInfo.Name.Contains("销售") Then Continue For

                wbook = Globals.ThisAddIn.Application.Workbooks.Open(fInfo.FullName)

                wbook.Windows.Item(1).Visible = True

                '   If wbook.Name.Contains("铁制件") Then key = PubDic("铁销售清单") Else key = PubDic("铝销售清单")

                For Each Sh As Worksheet In wbook.Sheets

                    Exl.Sheet = Sh

                    If Exl.StaRow = 9 OrElse Exl.StaRow = 7 Then

                        'If Not SelRange.ContainsKey(key) Then SelRange.Add(key, New LinkedList(Of Range))

                        Sh.Range("B" & Exl.StaRow & ":" & ChrW(64 + Exl.ColNum) & Exl.EndRow("c")).Copy()
                        shmin.Range("b" & ro).PasteSpecial()
                        ro = shmin.Range("c65536").End(XlDirection.xlUp).Row + 1

                    End If

                Next


                wbook.Close()

            Next

            Exl.Sheet = shmin

            '    For Each Val As String In SelRange.Keys

            '        wbook = Application.Workbooks.Add

            '        wbook.Windows.Item(1).Visible = True

            '        Exl.Sheet1 = wbook.ActiveSheet

            '        Exl.Sheet1.Name = Val

            '        Exl.GetSheetType(Exl.DataSet.Tables("表信息").Select("编码类型='" & Val & "'")(0))

            '        Exl.InsertTH()

            '        For Each ra As Range In SelRange(Val)

            '            If Exl.EndRow("c") <= Exl.StaRow Then ra.Copy(Exl.Sheet.Range("B" & Exl.StaRow)) Else ra.Copy(Exl.Sheet.Range("B" & Exl.EndRow("c") + 1))

            '        Next

            SerialCombin("b")

            w1.SaveAs(Exl.FliePath & "\" & Exl.ProjectName, XlFileFormat.xlOpenXMLWorkbook)

            '    Next

            '    For Each w As Workbook In Globals.ThisAddIn.Application.Workbooks
            '        If Not w.Name.Contains("销售") Then w.Close()
            '    Next
            MessageBox.Show("计算完成")
        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 解锁共享文件薄
    ''' </summary>
    Private Function OpenShardWorkBook() As Boolean

        If Exl.WorkBook.MultiUserEditing Then

            MsgBox("共享文件薄无法计算，请取消共享")

        Else

            Return True

        End If

    End Function
    Sub DaringN()

        Dim wb As Workbook
        Dim ws As Worksheet
        Dim op As New OpenFileDialog
        '  op.Filter = "(*.xlsx)|*.xlsx"
        op.Title = "选择图纸编号对应表"
        Dim pa As String = ""
        If op.ShowDialog = DialogResult.OK Then
            pa = op.FileName
        End If

        Dim sta As Integer = 0

        If pa <> "" Then

            For Each wbit As Workbook In Application.Workbooks

                If wbit.Name = "图纸编号" Then
                    wb = wbit
                    ws = wb.Sheets("Sheet1")
                    sta = 1
                End If
            Next

            If IsNothing(wb) Then
                wb = Application.Workbooks.Open(pa)
                ws = wb.Sheets("Sheet1")
            End If

            Dim h As Integer = ws.Range("a65536").End(XlDirection.xlUp).Row
            Dim di As New Dictionary(Of String, String)
            Dim di1 As New Dictionary(Of String, String)
            If h > 1 Then

                For index = 2 To h

                    Dim val = ws.Range("b" & index).Value.ToString

                    Dim val0 = Regex.Match(val, pattern:="-T\d{1,2}").Value

                    Dim val1 = Regex.Replace(val, pattern:="-T\d{1,2}", replacement:="%")

                    Dim co1 As String = Regex.Replace(val1, pattern:="\([A-Za-z]+\d{0,2}\)|\d{1,4}|\s", replacement:="")

                    Dim co = Regex.Replace(co1, pattern:="%", replacement:=val0)

                    If Not di.ContainsKey(co) Then di.Add(co, ws.Range("a" & index).Value)
                    If Not di1.ContainsKey(val) Then di1.Add(val, ws.Range("a" & index).Value)

                Next

                If sta = 0 Then
                    wb.Save()
                    wb.Close()
                End If

                For Each wsh As Worksheet In Exl.WorkBook.Sheets

                    Exl.Sheet = wsh

                    For index = Exl.StaRow To Exl.EndRow()

                        If IsNothing(Exl.Sheet.Range(ChrW(Exl.ColNum + 64) & index).Value) OrElse
                            Regex.IsMatch(Exl.Sheet.Range(ChrW(Exl.ColNum + 64) & index).Value, "(^\s*$|见加工图)") Then

                            Dim inp As String = ProductionSer.MinusStaCode(Exl.Sheet.Range("C" & index).Value)

                            If Not di1.ContainsKey(inp) Then

                                inp = Regex.Replace(inp, pattern:="\d{1,4}\.?\d{0,4}|\s", replacement:="")
                                inp = Regex.Replace(inp, pattern:="\([ABCDE]\)", replacement:="(X)")

                                If di.ContainsKey(inp) Then

                                    Exl.Sheet.Range(ChrW(Exl.ColNum + 64) & index).Value = di(inp)

                                End If

                            Else

                                Exl.Sheet.Range(ChrW(Exl.ColNum + 64) & index).Value = di1(inp)

                            End If

                        End If

                    Next

                Next

            End If

        End If

    End Sub
    Sub DaringN1()

        Dim wb As Workbook = Application.Workbooks.Add()
        Dim ws As Worksheet = wb.Sheets("Sheet1")
        ws.Range("a1").Value = "没有生产图编码"
        Dim i As Integer = 2
        For Each wsh As Worksheet In Exl.WorkBook.Sheets

            Exl.Sheet = wsh

            For index = Exl.StaRow To Exl.EndRow()

                If IsNothing(Exl.Sheet.Range(ChrW(Exl.ColNum + 64) & index).Value) Then

                    ws.Range("a" & i).Value = Exl.Sheet.Range("c" & index).Value
                    i = i + 1

                End If

            Next

        Next

        Dim n As String = Path.GetFileNameWithoutExtension(Exl.WorkBook.FullName)

        wb.SaveAs(Exl.WorkBook.Path & "\" & n & "-没有生产图编码表", XlFileFormat.xlOpenXMLWorkbook)
        wb.Close()

    End Sub

    ''' <summary>
    ''' 变更清单,打包清单包含铝字按照铝件生产清单；包含铁字，按照铁件生产清单
    ''' </summary>
    Sub changeW()

        Dim op As New OpenFileDialog
        '  op.Filter = "(*.xlsx)|*.xlsx"
        op.Title = "选择打包清单"
        op.Multiselect = True

        Dim pa As String()
        If op.ShowDialog = DialogResult.OK Then
            pa = op.FileNames
        End If

        Dim pathwb As String

        If Not IsNothing(pa) AndAlso pa.Length >= 2 Then

            pathwb = Path.GetDirectoryName(pa(0))

            Dim dchange As Dictionary(Of String, Dictionary(Of String, Integer)) = New Dictionary(Of String, Dictionary(Of String, Integer))
            Dim dchange_suf As Dictionary(Of String, Dictionary(Of String, String)) = New Dictionary(Of String, Dictionary(Of String, String))

            Dim dchange_copy As Dictionary(Of String, Dictionary(Of String, Integer)) = New Dictionary(Of String, Dictionary(Of String, Integer))

            Dim dchange_copy_1 As Dictionary(Of String, Dictionary(Of String, String)) = New Dictionary(Of String, Dictionary(Of String, String))

            Dim addvalue As Dictionary(Of String, ArrayList) = New Dictionary(Of String, ArrayList)

            Dim reducevalue As Dictionary(Of String, ArrayList) = New Dictionary(Of String, ArrayList)

            Dim addvalue_t As Dictionary(Of String, ArrayList) = New Dictionary(Of String, ArrayList)

            Dim reducevalue_t As Dictionary(Of String, ArrayList) = New Dictionary(Of String, ArrayList)

            Dim changworks As List(Of String) = New List(Of String)
            Dim starworks As List(Of String) = New List(Of String)

            Dim filename_prf As String = ""

            For Each bname In pa

                Dim bnmid As String() = Path.GetFileNameWithoutExtension(bname).Split("-")

                If bnmid.Length = 2 AndAlso bnmid(1).Replace(" ", "") = "变更" Then

                    changworks.Add(bname)

                    filename_prf = bnmid(0).Replace("打包清单", "")

                Else

                    starworks.Add(bname)

                End If

            Next

            If changworks.Count > 0 AndAlso starworks.Count > 0 Then

                For Each changwork In changworks

                    Dim wbchange As Workbook = Application.Workbooks.Open(changwork)

                    Dim filename_s = Path.GetFileNameWithoutExtension(changwork)

                    For Each wshange As Worksheet In wbchange.Sheets

                        Dim h As Integer = wshange.Range("b65536").End(XlDirection.xlUp).Row

                        Dim schange As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)

                        Dim schange_1 As Dictionary(Of String, String) = New Dictionary(Of String, String)

                        For index = 3 To h

                            Dim changvaleu = wshange.Range("b" & index).Value

                            'If Regex.IsMatch(changvaleu, "-BG\d{1,2}$") Then

                            '    Dim changvaleu_nosuf = Regex.Replace(changvaleu, pattern:="-BG\d{1,2}$", replacement:="")

                            '    If schange.ContainsKey(changvaleu_nosuf) Then
                            '        schange(changvaleu_nosuf) = schange(changvaleu_nosuf) + wshange.Range("c" & index).Value
                            '    Else
                            '        schange.Add(changvaleu_nosuf, wshange.Range("c" & index).Value)
                            '    End If

                            '    If filename_s.Contains("铁") AndAlso Not schange_1.ContainsKey(changvaleu_nosuf) Then

                            '        schange_1.Add(changvaleu_nosuf, "铁")

                            '    ElseIf filename_s.Contains("铝") AndAlso Not schange_1.ContainsKey(changvaleu_nosuf) Then

                            '        schange_1.Add(changvaleu_nosuf, "铝")

                            '    End If

                            '    If Not dchange_suf.ContainsKey(wshange.Name) Then
                            '        Dim dchange_suf_code As Dictionary(Of String, String) = New Dictionary(Of String, String)
                            '        dchange_suf_code.Add(changvaleu_nosuf, changvaleu)
                            '        dchange_suf.Add(wshange.Name, dchange_suf_code)
                            '    Else
                            '        dchange_suf(wshange.Name).Add(changvaleu_nosuf, changvaleu)
                            '    End If

                            'Else

                            If filename_s.Contains("铁") AndAlso Not schange_1.ContainsKey(changvaleu) Then

                                    schange_1.Add(changvaleu, "铁")

                                ElseIf filename_s.Contains("铝") AndAlso Not schange_1.ContainsKey(changvaleu) Then

                                    schange_1.Add(changvaleu, "铝")

                                End If

                                If schange.ContainsKey(changvaleu) Then
                                    schange(changvaleu) = schange(changvaleu) + wshange.Range("c" & index).Value
                                Else
                                    schange.Add(changvaleu, wshange.Range("c" & index).Value)
                                End If

                           'End If

                        Next

                        dchange.Add(wshange.Name, schange)

                        Dim schange_2 As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)(schange)

                        dchange_copy.Add(wshange.Name, schange_2)
                        dchange_copy_1.Add(wshange.Name, schange_1)

                    Next

                    wbchange.Close()

                Next

                For Each starwork In starworks

                    Dim wbstar As Workbook = Application.Workbooks.Open(starwork)

                    Dim reducevalue_mid As Dictionary(Of String, ArrayList)

                    Dim filename = Path.GetFileName(starwork)

                    For Each wshange As Worksheet In wbstar.Sheets

                        Dim h As Integer = wshange.Range("b65536").End(XlDirection.xlUp).Row

                        For index = 3 To h

                            Dim value As String = wshange.Range("b" & index).Value

                            If dchange.ContainsKey(wshange.Name) AndAlso dchange(wshange.Name).ContainsKey(value) Then

                                dchange_copy(wshange.Name).Remove(value)

                                Dim count As Integer = dchange(wshange.Name)(value) - Convert.ToInt32(wshange.Range("c" & index).Value)

                                If count > 0 Then '数量不同说明存在变更

                                    If filename.Contains("铝") Then

                                        reducevalue_mid = addvalue

                                    ElseIf filename.Contains("铁") Then

                                        reducevalue_mid = addvalue_t

                                    Else

                                        Return

                                    End If

                                    If Not reducevalue_mid.ContainsKey(value) Then

                                        Dim addlist As ArrayList = New ArrayList

                                        addlist.Add(count)
                                        addlist.Add(wshange.Name & "(" & count & ")")

                                        ' If dchange_suf.ContainsKey(wshange.Name) AndAlso dchange_suf(wshange.Name).ContainsKey(value) Then
                                        '  addlist.Add(dchange_suf(wshange.Name)(value))
                                        'End If

                                        reducevalue_mid.Add(value, addlist)

                                    Else

                                        Dim addlist As ArrayList = reducevalue_mid(value)

                                        addlist(0) = addlist(0) + count

                                        addlist(1) = addlist(1) & "、" & wshange.Name & "(" & count & ")"

                                        'If addlist.Count = 3 Then

                                        ' If Not dchange_suf.ContainsKey(wshange.Name) OrElse Not dchange_suf(wshange.Name).ContainsKey(value) Then addlist.RemoveAt(2)

                                        'End If

                                    End If

                                ElseIf count < 0 Then

                                    If filename.Contains("铝") Then

                                        reducevalue_mid = reducevalue

                                    ElseIf filename.Contains("铁") Then

                                        reducevalue_mid = reducevalue_t

                                    Else

                                        Return

                                    End If

                                    If Not reducevalue_mid.ContainsKey(value) Then

                                        Dim reducelist As ArrayList = New ArrayList

                                        reducelist.Add(-count)
                                        reducelist.Add(wshange.Name & "(" & -count & ")")

                                        reducevalue_mid.Add(value, reducelist)

                                    Else

                                        Dim reducelist As ArrayList = reducevalue_mid(value)

                                        reducelist(0) = reducelist(0) - count

                                        If reducelist(1) <> wshange.Name Then reducelist(1) = reducelist(1) & "、" & wshange.Name & "(" & -count & ")"

                                    End If

                                End If

                            ElseIf Not dchange.ContainsKey(wshange.Name) OrElse Not dchange(wshange.Name).ContainsKey(value) Then '当模板完全替换时

                                If filename.Contains("铝") Then

                                    reducevalue_mid = reducevalue

                                ElseIf filename.Contains("铁") Then

                                    reducevalue_mid = reducevalue_t

                                Else

                                    Return

                                End If

                                If Not reducevalue_mid.ContainsKey(value) Then

                                    Dim reducelist As ArrayList = New ArrayList

                                    reducelist.Add(Convert.ToInt32(wshange.Range("c" & index).Value))
                                    reducelist.Add(wshange.Name & "(" & Convert.ToInt32(wshange.Range("c" & index).Value) & ")")

                                    reducevalue_mid.Add(value, reducelist)

                                Else

                                    Dim reducelist As ArrayList = reducevalue_mid(value)

                                    reducelist(0) = reducelist(0) + Convert.ToInt32(wshange.Range("c" & index).Value)

                                    reducelist(1) = reducelist(1) & "、" & wshange.Name & "(" & Convert.ToInt32(wshange.Range("c" & index).Value) & ")"

                                End If

                            End If

                        Next

                    Next

                    wbstar.Close()

                Next

                If dchange_copy.Count <> 0 Then '当新增件在对应部位没有模板时

                    Dim reducevalue_mid As Dictionary(Of String, ArrayList)

                    For Each ws In dchange_copy.Keys

                        For Each st In dchange_copy(ws).Keys

                            If dchange_copy_1(ws)(st) = "铁" Then
                                reducevalue_mid = addvalue_t
                            ElseIf dchange_copy_1(ws)(st) = "铝" Then
                                reducevalue_mid = addvalue
                            End If

                            If Not reducevalue_mid.ContainsKey(st) Then

                                Dim addlist As ArrayList = New ArrayList

                                addlist.Add(dchange_copy(ws)(st))
                                addlist.Add(ws & "(" & dchange_copy(ws)(st) & ")")

                                '  If dchange_suf.ContainsKey(ws) AndAlso dchange_suf(ws).ContainsKey(st) Then
                                ' addlist.Add(dchange_suf(ws)(st))
                                ' If

                                reducevalue_mid.Add(st, addlist)

                            Else

                                Dim addlist As ArrayList = reducevalue_mid(st)

                                addlist(0) = addlist(0) + dchange_copy(ws)(st)

                                addlist(1) = addlist(1) & "、" & ws & "(" & dchange_copy(ws)(st) & ")"

                                ' If addlist.Count = 3 Then
                                '
                                ' If Not dchange_suf.ContainsKey(ws) OrElse Not dchange_suf(ws).ContainsKey(st) Then addlist.RemoveAt(2)

                                'End If

                            End If

                        Next

                    Next

                End If

                '--------------------------------------------
                '移动
                Dim moveS As String

                Dim totlmove As Integer = 0

                Dim reducevalue_mid_l As Dictionary(Of String, ArrayList) = New Dictionary(Of String, ArrayList)(reducevalue)

                For Each rangevalue In reducevalue_mid_l.Keys

                    If addvalue.ContainsKey(rangevalue) Then

                        Dim addlist = addvalue(rangevalue)
                        Dim reducelist = reducevalue(rangevalue)

                        If addlist(0) = reducelist(0) Then

                            moveS = moveS & Chr(10) & rangevalue & "(" & addlist(0) & ")" & "件：" & addlist(0) & "件由" & reducelist(1) & "变动到" & addlist(1)

                            totlmove += addlist(0)

                            reducevalue.Remove(rangevalue)
                            addvalue.Remove(rangevalue)

                        ElseIf addlist(0) > reducelist(0) Then

                            Dim areavalues = addlist(1).ToString.Split("、")

                            Dim areavalues_cope As List(Of String) = New List(Of String)(areavalues)

                            Dim totl As Integer = 0

                            Dim movearea As String = ""

                            For i = 0 To areavalues.Length - 1

                                Dim areavalue_number As Integer = Regex.Match(areavalues(i), "\d{1,2}(?=\)$)").Value

                                totl += areavalue_number

                                If totl = reducelist(0) Then

                                    movearea = movearea & areavalues(i)

                                    areavalues_cope.RemoveAt(0)

                                    Exit For

                                ElseIf totl < reducelist(0) Then

                                    movearea = movearea & areavalues(i) & "、"

                                    areavalues_cope.RemoveAt(0)

                                Else

                                    totl = totl - areavalue_number

                                    Dim dv = (reducelist(0) - totl)

                                    Dim dv1 = areavalue_number - dv

                                    areavalues(i) = Regex.Replace(areavalues(i), "\d{1,2}(?=\)$)", replacement:=dv1.ToString)

                                    areavalues_cope(0) = areavalues(i)

                                    Dim area_end = Regex.Replace(areavalues(i), "\d{1,2}(?=\)$)", replacement:=dv.ToString)

                                    movearea = movearea & area_end

                                    Exit For

                                End If

                            Next

                            addlist(0） = addlist(0) - reducelist(0)

                            Dim addlist_1_value As String = ""

                            For i = 0 To areavalues_cope.Count - 1

                                If i < areavalues_cope.Count - 1 Then

                                    addlist_1_value = addlist_1_value & areavalues_cope(i) & "、"

                                Else

                                    addlist_1_value = addlist_1_value & areavalues_cope(i)

                                End If

                            Next

                            addlist(1） = addlist_1_value

                            reducevalue.Remove(rangevalue)
                            moveS = moveS & Chr(10) & rangevalue & "(" & reducelist(0) & ")" & "件：" & reducelist(0) & "件由" & reducelist(1) & "变动到" & movearea

                            totlmove += reducelist(0)

                        Else

                            Dim areavalues = reducelist(1).ToString.Split("、")

                            Dim areavalues_cope As List(Of String) = New List(Of String)(areavalues)

                            Dim totl As Integer = 0

                            Dim movearea As String = ""

                            For i = 0 To areavalues.Length - 1

                                Dim areavalue_number As Integer = Regex.Match(areavalues(i), "\d{1,2}(?=\)$)").Value

                                totl += areavalue_number

                                If totl = addlist(0) Then

                                    movearea = movearea & areavalues(i)

                                    areavalues_cope.RemoveAt(0)

                                    Exit For

                                ElseIf totl < addlist(0) Then

                                    movearea = movearea & areavalues(i) & "、"

                                    areavalues_cope.RemoveAt(0)

                                Else

                                    totl = totl - areavalue_number

                                    Dim dv = (addlist(0) - totl)

                                    Dim dv1 = areavalue_number - dv

                                    areavalues(i) = Regex.Replace(areavalues(i), "\d{1,2}(?=\)$)", replacement:=dv1.ToString)

                                    areavalues_cope(0) = areavalues(i)

                                    Dim area_end = Regex.Replace(areavalues(i), "\d{1,2}(?=\)$)", replacement:=dv.ToString)

                                    movearea = movearea & area_end

                                    Exit For

                                End If

                            Next

                            reducelist(0） = reducelist(0) - addlist(0)

                            Dim reducelist_1_value As String = ""

                            For i = 0 To areavalues_cope.Count - 1

                                If i < areavalues_cope.Count - 1 Then

                                    reducelist_1_value = reducelist_1_value & areavalues_cope(i) & "、"

                                Else

                                    reducelist_1_value = reducelist_1_value & areavalues_cope(i)

                                End If

                            Next

                            reducelist(1） = reducelist_1_value

                            addvalue.Remove(rangevalue)
                            moveS = moveS & Chr(10) & rangevalue & "(" & addlist(0) & ")" & "件：" & addlist(0) & "件由" & movearea & "变动到" & addlist(1)

                            totlmove += addlist(0)

                        End If

                    End If

                Next

                Dim reducevalue_mid_t As Dictionary(Of String, ArrayList) = New Dictionary(Of String, ArrayList)(reducevalue_t)

                For Each rangevalue In reducevalue_mid_t.Keys

                    If addvalue_t.ContainsKey(rangevalue) Then

                        Dim addlist = addvalue_t(rangevalue)
                        Dim reducelist = reducevalue_t(rangevalue)

                        If addlist(0) = reducelist(0) Then

                            moveS = moveS & Chr(10) & rangevalue & "(" & addlist(0) & ")" & "件：" & addlist(0) & "件由" & reducelist(1) & "变动到" & addlist(1)

                            totlmove += addlist(0)

                            reducevalue_t.Remove(rangevalue)
                            addvalue_t.Remove(rangevalue)

                        ElseIf addlist(0) > reducelist(0) Then

                            Dim areavalues = addlist(1).ToString.Split("、")

                            Dim areavalues_cope As List(Of String) = New List(Of String)(areavalues)

                            Dim totl As Integer = 0

                            Dim movearea As String = ""

                            For i = 0 To areavalues.Length - 1

                                Dim areavalue_number As Integer = Regex.Match(areavalues(i), "\d{1,2}(?=\)$)").Value

                                totl += areavalue_number

                                If totl = reducelist(0) Then

                                    movearea = movearea & areavalues(i)

                                    areavalues_cope.RemoveAt(0)

                                    Exit For

                                ElseIf totl < reducelist(0) Then

                                    movearea = movearea & areavalues(i) & "、"

                                    areavalues_cope.RemoveAt(0)

                                Else

                                    totl = totl - areavalue_number

                                    Dim dv = (reducelist(0) - totl)

                                    Dim dv1 = areavalue_number - dv

                                    areavalues(i) = Regex.Replace(areavalues(i), "\d{1,2}(?=\)$)", replacement:=dv1.ToString)

                                    areavalues_cope(0) = areavalues(i)

                                    Dim area_end = Regex.Replace(areavalues(i), "\d{1,2}(?=\)$)", replacement:=dv.ToString)

                                    movearea = movearea & area_end

                                    Exit For

                                End If

                            Next

                            addlist(0） = addlist(0) - reducelist(0)

                            Dim addlist_1_value As String = ""

                            For i = 0 To areavalues_cope.Count - 1

                                If i < areavalues_cope.Count - 1 Then

                                    addlist_1_value = addlist_1_value & areavalues_cope(i) & "、"

                                Else

                                    addlist_1_value = addlist_1_value & areavalues_cope(i)

                                End If

                            Next

                            addlist(1） = addlist_1_value

                            reducevalue_t.Remove(rangevalue)
                            moveS = moveS & Chr(10) & rangevalue & "(" & reducelist(0) & ")" & "件：" & reducelist(0) & "件由" & reducelist(1) & "变动到" & movearea

                            totlmove += reducelist(0)

                        Else

                            Dim areavalues = reducelist(1).ToString.Split("、")

                            Dim areavalues_cope As List(Of String) = New List(Of String)(areavalues)

                            Dim totl As Integer = 0

                            Dim movearea As String = ""

                            For i = 0 To areavalues.Length - 1

                                Dim areavalue_number As Integer = Regex.Match(areavalues(i), "\d{1,2}(?=\)$)").Value

                                totl += areavalue_number

                                If totl = addlist(0) Then

                                    movearea = movearea & areavalues(i)

                                    areavalues_cope.RemoveAt(0)

                                    Exit For

                                ElseIf totl < addlist(0) Then

                                    movearea = movearea & areavalues(i) & "、"

                                    areavalues_cope.RemoveAt(0)

                                Else

                                    totl = totl - areavalue_number

                                    Dim dv = (addlist(0) - totl)

                                    Dim dv1 = areavalue_number - dv

                                    areavalues(i) = Regex.Replace(areavalues(i), "\d{1,2}(?=\)$)", replacement:=dv1.ToString)

                                    areavalues_cope(0) = areavalues(i)

                                    Dim area_end = Regex.Replace(areavalues(i), "\d{1,2}(?=\)$)", replacement:=dv.ToString)

                                    movearea = movearea & area_end

                                    Exit For

                                End If

                            Next

                            reducelist(0） = reducelist(0) - addlist(0)

                            Dim reducelist_1_value As String = ""

                            For i = 0 To areavalues_cope.Count - 1

                                If i < areavalues_cope.Count - 1 Then

                                    reducelist_1_value = reducelist_1_value & areavalues_cope(i) & "、"

                                Else

                                    reducelist_1_value = reducelist_1_value & areavalues_cope(i)

                                End If

                            Next

                            reducelist(1） = reducelist_1_value

                            addvalue_t.Remove(rangevalue)
                            moveS = moveS & Chr(10) & rangevalue & "(" & addlist(0) & ")" & "件：" & addlist(0) & "件由" & movearea & "变动到" & addlist(1)

                            totlmove += addlist(0)

                        End If

                    End If

                Next
                '---------------------------------------------
                '新增铝件
                If addvalue.Count > 0 Then

                    Exl.WorkBook = Application.Workbooks.Add
                    Exl.Sheet1 = Exl.WorkBook.ActiveSheet

                    Exl.GetSheetType(Exl.DataSet.Tables("表信息").Select("headerName='" & "变更新增铝件明细清单" & "'")(0))

                    Exl.InsertTH()

                    Dim index_1 As Integer = Exl.StaRow

                    For Each value In addvalue.Keys

                        If addvalue(value).Count = 3 Then
                            Exl.Sheet1.Range("C" & index_1).Value = addvalue(value)(2)
                        Else
                            Exl.Sheet1.Range("C" & index_1).Value = value
                        End If

                        Exl.Sheet1.Range("h" & index_1).Value = addvalue(value)(0)
                        Exl.Sheet1.Range("k" & index_1).Value = addvalue(value)(1)
                        index_1 += 1

                    Next

                    SerialCal()

                    Exl.Sheet1.Name = "新增"
                    Exl.WorkBook.SaveAs(pathwb & "\" & filename_prf & "变更新增铝件明细清单")
                    Exl.WorkBook.Close()

                End If

                '--------------------------------------
                '替换铝件
                If reducevalue.Count > 0 Then

                    Exl.WorkBook = Application.Workbooks.Add
                    Exl.Sheet1 = Exl.WorkBook.ActiveSheet

                    Exl.GetSheetType(Exl.DataSet.Tables("表信息").Select("headerName='" & "变更替换铝件明细清单" & "'")(0))

                    Exl.InsertTH()

                    Dim index_2 As Integer = Exl.StaRow

                    For Each value In reducevalue.Keys

                        Exl.Sheet1.Range("C" & index_2).Value = value
                        Exl.Sheet1.Range("h" & index_2).Value = reducevalue(value)(0)
                        Exl.Sheet1.Range("k" & index_2).Value = reducevalue(value)(1)
                        index_2 += 1

                    Next

                    SerialCal()

                    Exl.Sheet1.Name = "替换"
                    Exl.WorkBook.SaveAs(pathwb & "\" & filename_prf & "变更替换铝件明细清单")
                    Exl.WorkBook.Close()

                End If

                '------------------------------------
                '替换铁件
                If reducevalue_t.Count > 0 Then

                    Exl.WorkBook = Application.Workbooks.Add
                    Exl.Sheet1 = Exl.WorkBook.ActiveSheet

                    Exl.GetSheetType(Exl.DataSet.Tables("表信息").Select("headerName='" & "变更替换铁件明细清单" & "'")(0))

                    Exl.InsertTH()

                    Dim index_2 As Integer = Exl.StaRow

                    For Each value In reducevalue_t.Keys

                        Exl.Sheet1.Range("C" & index_2).Value = value
                        Exl.Sheet1.Range("g" & index_2).Value = reducevalue_t(value)(0)
                        Exl.Sheet1.Range("j" & index_2).Value = reducevalue_t(value)(1)
                        index_2 += 1

                    Next

                    SerialCal()

                    Exl.Sheet1.Name = "替换"
                    Exl.WorkBook.SaveAs(pathwb & "\" & filename_prf & "变更替换铁件明细清单")
                    Exl.WorkBook.Close()

                End If

                '------------------------------------------------------
                '新增铁件
                If addvalue_t.Count > 0 Then

                    Exl.WorkBook = Application.Workbooks.Add
                    Exl.Sheet1 = Exl.WorkBook.ActiveSheet

                    Exl.GetSheetType(Exl.DataSet.Tables("表信息").Select("headerName='" & "变更新增铁件明细清单" & "'")(0))

                    Exl.InsertTH()

                    Dim index_1 As Integer = Exl.StaRow

                    For Each value In addvalue_t.Keys

                        If addvalue_t(value).Count = 3 Then
                            Exl.Sheet1.Range("C" & index_1).Value = addvalue_t(value)(2)
                        Else
                            Exl.Sheet1.Range("C" & index_1).Value = value
                        End If

                        Exl.Sheet1.Range("C" & index_1).Value = value
                        Exl.Sheet1.Range("g" & index_1).Value = addvalue_t(value)(0)
                        Exl.Sheet1.Range("j" & index_1).Value = addvalue_t(value)(1)
                        index_1 += 1

                    Next

                    SerialCal()

                    Exl.Sheet1.Name = "新增"
                    Exl.WorkBook.SaveAs(pathwb & "\" & filename_prf & "变更新增铁件明细清单")
                    Exl.WorkBook.Close()

                End If

                '------------------------------------------------
                '移动单
                If Not IsNothing(moveS) Then

                    Exl.WorkBook = Application.Workbooks.Add
                    Exl.Sheet1 = Exl.WorkBook.ActiveSheet

                    Exl.Sheet1.Range("a1").Value = moveS
                    Exl.Sheet1.Range("b1").Value = "总计移动：" & totlmove.ToString
                    Exl.WorkBook.SaveAs(pathwb & "\" & filename_prf & "移动单")
                    Exl.WorkBook.Close()

                End If

            End If

        End If

        MessageBox.Show("计算完成")

    End Sub

End Class
''' <summary>
''' 清单相关计算
''' </summary>
Module Cal
    ''' <summary>
    ''' 清单计算
    ''' </summary>
    Sub SerialCal()
        Try
            If Exl.ColNum < 4 Then Exit Sub

            For Exl.ActiveRow = Exl.StaRow To Exl.EndRow("c")

                Dim producSer As ProductionSer : If Not IsNothing(Exl.Sheet.Range("c" & Exl.ActiveRow).Value) Then

                    producSer = New ProductionSer(ProductionSer.MinusStaCode(Regex.Replace(Exl.Sheet.Range("c" & Exl.ActiveRow).Value, "-BG\d{1,2}$", "")))

                Else

                    Continue For

                End If

                producSer.SerialCal()

                For Each j In Exl.DicMap.Keys

                    If producSer.SerDic.ContainsKey(Exl.DicMap(j)) Then

                        If Not IsNothing(Exl.Sheet.Range(j & Exl.ActiveRow).Value) OrElse Exl.Sheet.Range(j & Exl.ActiveRow).Value <> "" Then Continue For

                        If producSer.SerDic(Exl.DicMap(j)).ToString.Contains("=") Then Exl.Sheet.Range(j & Exl.ActiveRow).Value = Method.GetValue5(producSer.SerDic(Exl.DicMap(j)), Exl.ActiveRow) : Continue For '值包含等号就将值替换当前单元格

                        Exl.Sheet.Range(j & Exl.ActiveRow).Value = producSer.SerDic(Exl.DicMap(j)) : Continue For '将具体值赋值到对应列

                    End If

                Next

            Next

            Exl.SumTo()

            Exl.SheetFormat("c")

            'Exl.SetPrintFormat()
        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try
    End Sub
    ''' <summary>
    ''' 清单分类
    ''' </summary>
    ''' <param name="DoCal">是否对分类清单进行计算</param>
    Sub SerialSort(ByVal DoCal As Boolean)
        Try
            If Exl.ColNum < 4 Then Exit Sub
            Dim ColNum As Integer = Exl.ColNum
            Dim producSer As ProductionSer

            Dim SheetDict As New Dictionary(Of String, Range)

            For i As Integer = Exl.StaRow To Exl.EndRow

                producSer = New ProductionSer(ProductionSer.MinusStaCode(Exl.Sheet.Range("c" & i).Value))

                If Not SheetDict.ContainsKey(producSer.GetSerType) Then

                    SheetDict.Add(producSer.GetSerType, Exl.Sheet.Range("B" & i & ":" & "J" & i))

                Else

                    SheetDict(producSer.GetSerType) = Globals.ThisAddIn.Application.Union(SheetDict(producSer.GetSerType), Exl.Sheet.Range("B" & i & ":" & "J" & i))

                End If

            Next

            For Each S As String In SheetDict.Keys

                Exl.WorkBook.Sheets.Add()

                Exl.Sheet1 = Exl.WorkBook.ActiveSheet : Exl.GetSheetType(Exl.DataSet.Tables("表信息").Select("编码类型='" & S & "'")(0))

                Exl.Sheet.Name = S

                Exl.InsertTH()

                SheetDict(S).Copy(Exl.Sheet.Range("B" & Exl.StaRow))

                If Exl.ColNum < ColNum Then

                    Exl.Sheet.Range("h" & Exl.StaRow & ":j" & Exl.EndRow("c")).Cut(Destination:=Exl.Sheet.Range("G" & Exl.StaRow))

                    Exl.Sheet.Range("h" & Exl.StaRow & ":j" & Exl.EndRow("c")).Clear()

                    SerialCal()

                ElseIf Exl.ColNum > ColNum Then

                    Exl.Sheet.Range("G" & Exl.StaRow & ":I" & Exl.EndRow("c")).Cut(Destination:=Exl.Sheet.Range("H" & Exl.StaRow))
                    Exl.Sheet.Range("I" & Exl.StaRow & ":j" & Exl.EndRow("c")).Clear()
                    SerialCal()

                Else

                    If DoCal Then SerialCal()

                    Exl.SortCol(col_2:="B")

                    Exl.SumTo()

                    Exl.SheetFormat("c")

                    '  Exl.SetPrintFormat()

                End If

            Next
        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try
    End Sub
    ''' <summary>
    ''' 编码检查
    ''' </summary>
    Sub CheckSerial()

        If Exl.ColNum < 4 Then Exit Sub

        For i As Integer = Exl.StaRow To Exl.EndRow

            If Exl.Sheet.Range("c" & i).Value.ToString.Length > 40 Then

                Exl.Sheet.Range("c" & i).Interior.ColorIndex = 4

            Else

                'Dim code As String

                'Dim DR As DataRow() = Exl.DataSet.Tables("判断配件").Select("Code='" & code & "'")

                'If DR.Length <> 0 AndAlso Not IsDBNull(DR(0)("CodeExamine")) Then

                '    If Not Regex.IsMatch(Exl.Sheet.Range("c" & i).Value, DR(0)("CodeExamine")) AndAlso
                '           Regex.IsMatch(Exl.Sheet.Range("c" & i).Value, "\d[LR]($|-)") Then

                '        Exl.Sheet.Range("c" & i).Interior.ColorIndex = 3

                '    End If

                'End If

            End If

        Next

    End Sub
    ''' <summary>
    ''' 表头正确则得到编码列和数量列
    ''' </summary>
    ''' <returns></returns>
    Function GetSelRange(Optional ByVal Sheet As Worksheet = Nothing) As Range

        If Sheet Is Nothing Then Sheet = Exl.Sheet

        Dim A1 As String = Sheet.Range("a1").Value
        Dim B1 As String = Sheet.Range("b1").Value

        If A1 <> "" And B1 <> "" Then

            If (A1.Contains("编") Or A1.Contains("号") Or A1.Contains("值")) AndAlso B1.Contains("数") Then

                Return Sheet.Range("A2" & ":" & "B" & Sheet.Range("A65536").End(XlDirection.xlUp).Row)

            End If

        End If

        Return Nothing

    End Function

End Module
''' <summary>
''' Excel参数
''' </summary>
Structure Exl

    ''' <summary>
    '''文件完整路径
    ''' </summary>
    Public Shared FliePath As String
    ''' <summary>
    ''' 项目名
    ''' </summary>
    Public Shared ProjectName As String
    ''' <summary>
    ''' 数据适配器
    ''' </summary>
    Public Shared DataSet As DataSet
    ''' <summary>
    ''' 激活行
    ''' </summary>
    Public Shared ActiveRow As Integer
    ''' <summary>
    ''' 工作簿
    ''' </summary>
    Private Shared WB As Workbook
    ''' <summary>
    ''' 工作表
    ''' </summary>
    Private Shared ST As Worksheet
    ''' <summary>
    ''' 表列数
    ''' </summary>
    Private Shared ColN As Integer
    ''' <summary>
    ''' 表起始行
    ''' </summary>
    Private Shared SRow As Integer
    ''' <summary>
    ''' 工作表格式数据库表名
    ''' </summary>
    Private Shared BomF As String
    ''' <summary>
    ''' 工作表打印格式数据库表名
    ''' </summary>
    Private Shared PrintF As String
    ''' <summary>
    ''' 工作表表头数据库表名
    ''' </summary>
    Private Shared headerName As String
    ''' <summary>
    ''' 表合计列
    ''' </summary>
    Private Shared Sum As String
    ''' <summary>
    ''' 替换内容
    ''' </summary>
    Private Shared Replace As String
    ''' <summary>
    ''' 列值对应关系字典
    ''' </summary>
    Private Shared _DicMap As New Dictionary(Of String, String)
    ''' <summary>
    ''' 当前激活的工作簿
    ''' </summary>
    ''' <returns></returns>
    Public Shared Property WorkBook() As Workbook

        Get

            Return WB

        End Get

        Set(value As Workbook)

            WB = value

        End Set

    End Property

    ''' <summary>
    ''' 当前激活的工作表，i指定激活表引索,0为当前激活表
    ''' </summary>      
    ''' <returns></returns>   
    Public Shared Property Sheet1() As Worksheet

        Get

            Return ST

        End Get

        Set(value As Worksheet)

            ST = value

        End Set

    End Property

    ''' <summary>
    ''' 当前激活的工作表，i指定激活表引索,0为当前激活表
    ''' </summary>      
    ''' <returns></returns>   
    Public Shared Property Sheet() As Worksheet

        Get

            Return ST

        End Get

        Set(value As Worksheet)
            Try
                ST = value : SRow = 1 : ColN = 1 : headerName = "" : BomF = "" : PrintF = "" : Sum = "" : Replace = "" : _DicMap = New Dictionary(Of String, String)

                Dim DR As DataRow()

                If Not IsNothing(ST.Range("a2").Value) AndAlso ST.Range("a2").Value.ToString.Contains("清单"） Then
                    DR = DataSet.Tables("表信息").Select("headerName='" & GetShHeaderT(ST.Range("a2").Value) & "'")
                    If DR.Length > 0 Then GetSheetType(DR(0))
                ElseIf Not IsNothing(ST.Range("a3").Value) AndAlso ST.Range("a3").Value.ToString.Contains("清单"） Then
                    DR = DataSet.Tables("表信息").Select("headerName='" & GetShHeaderT(ST.Range("a3").Value) & "'")
                    If DR.Length > 0 Then GetSheetType(DR(0))
                ElseIf Not IsNothing(ST.Range("a1").Value) AndAlso ST.Range("a1").Value.ToString.Contains("打包") Then
                    DR = DataSet.Tables("表信息").Select("headerName='" & GetShHeader(ST.Range("a1").Value) & "'")
                    If DR.Length > 0 Then GetSheetType(DR(0))
                ElseIf Not IsNothing(ST.Range("a1").Value) AndAlso (ST.Range("a1").Value.ToString.Contains("编") Or ST.Range("a1").Value.ToString.Contains("码")) Then
                    SRow = 2 : ColN = 2 : headerName = "" : BomF = "" : PrintF = "" : Sum = "" : Replace = "" : _DicMap = New Dictionary(Of String, String) : Exit Property
                End If

                If SRow = 1 Then

                    If Not IsNothing(ST.Range("a2").Value) AndAlso ST.Range("a2").Value.ToString() = "序号" Then

                        DR = DataSet.Tables("表信息").Select("headerName='" & "打包" & "'")
                        If DR.Length > 0 Then GetSheetType(DR(0))

                    ElseIf Not IsNothing(ST.Range("a7").Value) AndAlso ST.Range("a7").Value.ToString() = "序号" Then

                        DR = DataSet.Tables("表信息").Select("headerName='" & "铝件" & "'")
                        If DR.Length > 0 Then GetSheetType(DR(0))

                    ElseIf Not IsNothing(ST.Range("a5").Value) AndAlso ST.Range("a5").Value.ToString() = "序号" Then

                        DR = DataSet.Tables("表信息").Select("headerName='" & "铁件" & "'")
                        If DR.Length > 0 Then GetSheetType(DR(0))

                    End If

                End If

            Catch ex As Exception
                Method.ExceptionWrite(ex)
            End Try

        End Set

    End Property
    ''' <summary>
    '''  工作表数据赋值
    ''' </summary>
    ''' <returns></returns>
    Public Shared Sub GetSheetType(drType As DataRow)
        Try
            If Not IsDBNull(drType("headerName")) Then headerName = drType("headerName") Else headerName = ""
            If Not IsDBNull(drType("BomF")) Then BomF = drType("BomF") Else BomF = ""
            If Not IsDBNull(drType("PrintF")) Then PrintF = drType("PrintF") Else PrintF = ""
            If Not IsDBNull(drType("Srow")) Then SRow = drType("Srow") Else SRow = 1
            If Not IsDBNull(drType("ColN")) Then ColN = drType("ColN") Else ColN = 1
            If Not IsDBNull(drType("合计")) Then Sum = drType("合计") Else Sum = ""
            If Not IsDBNull(drType("替换")) Then Replace = drType("替换") Else Replace = ""
            If Not IsDBNull(drType("Colum")) Then _DicMap = GetColumHT(drType) Else _DicMap = New Dictionary(Of String, String)
        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try
    End Sub
    ''' <summary>
    ''' 得到对应的列值关系字典
    ''' </summary>
    ''' <param name="DR">对应关系数据行</param>
    ''' <returns></returns>
    Private Shared Function GetColumHT(ByVal DR As DataRow) As Dictionary(Of String, String)

        Dim Hat As New Dictionary(Of String, String)

        If Not IsDBNull(DR("Colum")) AndAlso Not IsDBNull(DR("Value")) Then

            Dim Str2() As String = DR("Colum").ToString.Split("-")

            Dim Str3() As String = DR("Value").ToString.Split("-")

            For i As Integer = 0 To Str2.Count - 1

                Hat.Add(Str2(i), Str3(i))

            Next

        End If

        Return Hat

    End Function

    ''' <summary>
    ''' 工作表其余部分计算，包括求和，序列等
    ''' </summary>
    ''' <param name="aSheet">需要计算的工作表</param>
    ''' <param name="SRow">起始行</param>
    ''' <param name="endRow">终止行</param>
    ''' <param name="Sum">求和的列</param>
    Shared Sub SumTo(Optional ByVal fill As Boolean = True)
        Try
            If IsNothing(ST) Then Exit Sub

            If fill Then ST.Range("a" & SRow).Value = 1

            Dim endR As Integer

            If SRow > 3 Then endR = EndRow("c") Else endR = EndRow

            If SRow < endR AndAlso fill Then

                ST.Range("a" & SRow).AutoFill(ST.Range("a" & SRow & ":a" & endR), XlAutoFillType.xlFillSeries)

            End If

            If Sum = "" Then Exit Sub

            ST.Range("a" & endR + 1).Value = "合计"

            Dim sumColum As String() = Sum.Split("-")

            For Each C As String In sumColum

                ST.Range(C & endR + 1).Value = "=SUM(" & C & SRow & ":" & C & endR & ")"

            Next
        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try
    End Sub
    ''' <summary>
    ''' 得到起始行
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property StaRow() As Integer

        Get

            Return SRow

        End Get

    End Property

    ''' <summary>
    ''' 得到列数
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property ColNum() As Integer

        Get

            Return ColN

        End Get

    End Property

    ''' <summary>
    ''' 数量列
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property NumValue() As String

        Get

            Return Sum.Split("-")(0)

        End Get

    End Property
    ''' <summary>
    ''' 得到行最大值
    ''' </summary>
    ''' <param name="Col">指定列</param>
    ''' <returns></returns>
    Public Shared ReadOnly Property EndRow(Optional ByVal Col As String = "b") As Integer

        Get

            Return ST.Range(Col & "65536").End(XlDirection.xlUp).Row

        End Get

    End Property
    ''' <summary>
    ''' 清单对应的清单格式表
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property BomFormat As String

        Get
            Return BomF
        End Get

    End Property
    ''' <summary>
    ''' 清单对应的赋值
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property DicMap As Dictionary(Of String, String)

        Get
            Return _DicMap
        End Get

    End Property
    ''' <summary>
    ''' 清单表头
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property Head As String

        Get
            Return headerName
        End Get

    End Property

    ''' <summary>
    ''' 得到表头信息
    ''' </summary>
    ''' <param name="value">表标题</param>
    ''' <returns></returns>
    Private Shared Function GetShHeader(ByVal value As String) As String
        Try
            If Not value.Contains("变层") Then

                If value.Contains("基本层铝件打包清单") Then Return "基本层铝件打包清单"
                If value.Contains("基本层铁件打包清单") Then Return "基本层铁件打包清单"

            Else

                If value.Contains("变层铝件打包清单") Then Return "变层铝件打包清单"
                If value.Contains("变层铁件打包清单") Then Return "变层铁件打包清单"

            End If

        Catch ex As Exception

            Method.ExceptionWrite(ex)

        End Try

        Return "打包清单"

    End Function

    Private Shared Function GetShHeaderT(ByVal value As String) As String

        Dim valRe As String = ""

        If value.Contains("销售") = False Then
            valRe = Regex.Replace(value, pattern:="\(.+\)", replacement:="")
        Else
            valRe = Regex.Replace(value, pattern:="\(\)", replacement:="")
        End If

        Return valRe

    End Function

    ''' <summary>
    ''' 排序
    ''' </summary>
    ''' <param name="col_1">排序第一关键列</param>
    ''' <param name="col_2">排序第二关键列</param>
    Shared Sub SortCol(Optional ByVal col_1 As String = "C", Optional ByVal col_2 As String = "")
        Try
            ST.Sort.SortFields.Clear()
            If col_2 <> "" Then ST.Sort.SortFields.Add(Key:=ST.Range(col_2 & SRow & ":" & col_2 & EndRow(col_1)), SortOn:=XlSortOn.xlSortOnValues, Order:=XlSortOrder.xlAscending, DataOption:=XlSortDataOption.xlSortNormal)
            ST.Sort.SortFields.Add(Key:=ST.Range(col_1 & SRow & ":" & col_1 & EndRow(col_1)), SortOn:=XlSortOn.xlSortOnValues, Order:=XlSortOrder.xlAscending, DataOption:=XlSortDataOption.xlSortNormal)
            With ST.Sort
                .SetRange(ST.Range("a" & SRow & ":" & ChrW(ColN + 64) & EndRow(col_1)))
                .Header = XlYesNoGuess.xlGuess
                .MatchCase = False
                .Orientation = XlSortOrientation.xlSortColumns
                .SortMethod = XlSortMethod.xlPinYin
                .Apply()
            End With
        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try
    End Sub
    ''' <summary>
    ''' 替换工作表中内容
    ''' </summary>
    Shared Sub RepValue(Optional ByVal reV As String = "")
        Try
            If IsNothing(ST) OrElse Replace = "" Then Exit Sub

            Dim gre As String() = Replace.Split(",")

            For Each re As String In gre

                Dim reStr As String() = re.Split("-")

                If reStr(1) = "%" AndAlso Not IsNothing(ST.Range(reStr(0)).Value) Then ST.Range(reStr(0)).Value = ST.Range(reStr(0)).Value.ToString.Replace("%", Exl.ProjectName) : Continue For

                If reV <> "" AndAlso reStr(1) = "!" AndAlso Not IsNothing(ST.Range(reStr(0)).Value) Then ST.Range(reStr(0)).Value = ST.Range(reStr(0)).Value.ToString.Replace("!", reV) : Continue For

                If reStr(1) = "****.**.**" AndAlso Not IsNothing(ST.Range(reStr(0)).Value) Then ST.Range(reStr(0)).Value = ST.Range(reStr(0)).Value.ToString.Replace("****.**.**", DateTime.Now.ToString("yyyy.MM.dd")) : Continue For

                Next
        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try
    End Sub
    ''' <summary>
    ''' 打包清单替换
    ''' </summary>
    ''' <param name="reV">替换的值</param>
    Shared Sub DRepValue(ByVal reV As String)
        Try
            If IsNothing(ST) AndAlso Replace = "" Then Exit Sub

            Dim gre As String() = Replace.Split(",")
            For Each re As String In gre

                Dim reStr As String() = re.Split("-")

            If reStr(1) = "%" Then ST.Range(reStr(0)).Value = ST.Range(reStr(0)).Value.ToString.Replace("%", reV)

                If reStr(1) = "!" Then ST.Range(reStr(0)).Value = ST.Range(reStr(0)).Value.ToString.Replace("!", Exl.ProjectName)

            Next
        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try
    End Sub
    ''' <summary>
    ''' 插入表头 
    ''' </summary>
    ''' <param name="TName">数据库中表名</param>
    Shared Sub InsertTH()
        Try
            If headerName = "" AndAlso IsNothing(ST) Then Exit Sub

            Dim Val As Object = Nothing
            Dim Range As String = ""
            Dim Value As New System.Action(Sub() ST.Range(Range).Value = Val)
            Dim Merge As New System.Action(Sub() ST.Range(Val).Merge())
            Dim Idiction As New Dictionary(Of String, System.Action) From {{"Value", Value}, {"Merge", Merge}}

            With DataSet.Tables(headerName).Rows

                Dim Hastable As New Hashtable

                For i As Integer = 0 To .Count - 1

                    Hastable = GetValue(.Item(i))

                    For Each j In Hastable.Keys

                        Range = .Item(i)("Range").ToString()

                        Val = Hastable(j)

                        Idiction(j).Invoke()

                    Next

                Next

            End With
        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 表格格式设置
    ''' </summary>
    ''' <param name="endRowOfCol">计算最大行的列</param>
    Shared Sub SheetFormat(Optional ByVal endRowOfCol As String = "")

        If BomF = "" AndAlso IsNothing(ST) Then Exit Sub

        Try

            Dim Val As Object = Nothing
            Dim Range As String = ""
            Dim NumberFormatLocal As New System.Action(Sub() ST.Range(Range).NumberFormatLocal = Val)
            Dim FontName As New System.Action(Sub() ST.Range(Range).Font.Name = Val)
            Dim FontSize As New System.Action(Sub() ST.Range(Range).Font.Size = CInt(Val))
            Dim Bold As New System.Action(Sub() ST.Range(Range).Font.Bold = True)
            Dim LineStyle As New System.Action(Sub()
                                                   Dim Value1 As String() = Val.ToString.Split(",")
                                                   Dim Value2 As String() = Val.ToString.Split("-")
                                                   If Value1.Length = 2 Then

                                                       ST.Range(Range).Borders.LineStyle = CInt(Value1(0))

                                                       ST.Range(Range).BorderAround2(Weight:=CInt(Value1(1)))

                                                   ElseIf Value2.Length = 2 Then

                                                       ST.Range(Range).Borders(Value2(1)).LineStyle = CInt(Value2(0))

                                                   Else

                                                       ST.Range(Range).Borders.LineStyle = CInt(Value1(0))

                                                   End If
                                               End Sub)
            Dim HAlignment As New System.Action(Sub() ST.Range(Range).HorizontalAlignment = CInt(Val))
            Dim VAlignment As New System.Action(Sub() ST.Range(Range).VerticalAlignment = CInt(Val))
            Dim Height As New System.Action(Sub() ST.Range(Range).RowHeight = CDbl(Val))
            Dim Width As New System.Action(Sub()
                                               Dim Str1(), Str2() As String
                                               Str1 = Range.Split("-")
                                               Str2 = Val.ToString.Split(",")
                                               For i As Integer = 0 To Str1.Length - 1
                                                   ST.Range(Str1(i) & ":" & Str1(i)).ColumnWidth = CDbl(Str2(i))
                                               Next
                                           End Sub)
            Dim Wraptext As New System.Action(Sub() ST.Range(Range).WrapText = True)
            Dim View As New System.Action(Sub()

                                              ST.PageSetup.PrintArea = ""

                                                  ST.Range(Range).Select()

                                                  ST.Application.ActiveWindow.View = XlWindowView.xlPageBreakPreview

                                                  ST.Application.ActiveWindow.Zoom = 100

                                          End Sub)
            Dim Idiction As New Dictionary(Of String, System.Action) From {{"FontName", FontName},
                {"FontSize", FontSize}, {"Bold", Bold}, {"LineStyle", LineStyle}, {"HAlignment", HAlignment},
                {"VAlignment", VAlignment}, {"Height", Height}, {"Width", Width},
                {"NumberFormatLocal", NumberFormatLocal}, {"Wraptext", Wraptext}, {"View", View}}
            Dim Hastable As New Hashtable

            With DataSet.Tables(BomF).Rows

                For i As Integer = 0 To .Count - 1

                    Hastable = GetValue(.Item(i))

                    Dim RangeValue As String = .Item(i)("Range").ToString()

                    If RangeValue.Contains("?") Then

                        Dim MV As Integer : If endRowOfCol = "" Then MV = EndRow Else MV = EndRow(endRowOfCol)
                        Dim V As Integer = RangeValue.Substring(1, RangeValue.IndexOf(":") - 1)

                        If MV >= V Then RangeValue = .Item(i)("Range").ToString.Replace("?", MV + 1)

                    End If

                    For Each j In Hastable.Keys

                        Range = RangeValue

                        Val = Hastable(j)

                        Idiction(j).Invoke()

                    Next

                Next

            End With

            ST.Range("a" & StaRow & ":j" & EndRow).Font.Bold = False

        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try

    End Sub

    ''' <summary>
    ''' 打印格式设置
    ''' </summary>
    Shared Sub SetPrintFormat()

        If PrintF = "" AndAlso IsNothing(ST) Then Exit Sub

        ST.Application.PrintCommunication = True

        ST.PageSetup.PrintArea = "$A$1:" & "$" & ChrW(Exl.ColNum + 64) & "$" & Exl.EndRow + 1

        ST.Application.PrintCommunication = False

        ' Exl.Sheet.PageSetup.PaperSize = XlPaperSize.xlPaperA4

        Try

            With DataSet.Tables(PrintF).Rows

                For i As Integer = 0 To .Count - 1

                    If Not IsDBNull(.Item(i).Item(1)) Then

                        Select Case .Item(i).Item(0)
                            Case "AlignMarginsHeaderFooter"
                                Try

                                    ST.PageSetup.AlignMarginsHeaderFooter = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try
                            Case "CenterFooter"

                                Try

                                    ST.PageSetup.CenterFooter = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "CenterHorizontally"

                                Try

                                    ST.PageSetup.CenterHorizontally = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "CenterVertically"

                                Try

                                    ST.PageSetup.CenterVertically = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "DifferentFirstPageHeaderFooter"

                                Try

                                    ST.PageSetup.AlignMarginsHeaderFooter = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "Draft"

                                Try

                                    ST.PageSetup.Draft = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "OddAndEvenPagesHeaderFooter"

                                Try

                                    ST.PageSetup.OddAndEvenPagesHeaderFooter = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "Orientation"

                                Try

                                    ST.PageSetup.Orientation = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "PaperSize"

                                Try

                                    ST.PageSetup.PaperSize = XlPaperSize.xlPaperA4

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "PrintGridlines"

                                Try

                                    ST.PageSetup.PrintGridlines = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "PrintHeadings"

                                Try

                                    ST.PageSetup.PrintHeadings = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "PrintQuality"

                                Try

                                    ST.PageSetup.PrintQuality = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "ScaleWithDocHeaderFooter"

                                Try

                                    ST.PageSetup.ScaleWithDocHeaderFooter = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "Zoom"

                                Try

                                    ST.PageSetup.Zoom = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "FitToPagesWide"

                                Try

                                    ST.PageSetup.FitToPagesWide = CInt(.Item(i).Item(1))

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "FitToPagesTall"

                                Try

                                    ST.PageSetup.FitToPagesTall = CInt(.Item(i).Item(1))

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "FooterMargin"

                                Try

                                    ST.PageSetup.FooterMargin = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "HeaderMargin"

                                Try

                                    ST.PageSetup.HeaderMargin = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "LeftMargin"

                                Try

                                    ST.PageSetup.LeftMargin = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "RightMargin"

                                Try

                                    ST.PageSetup.RightMargin = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "TopMargin"

                                Try

                                    ST.PageSetup.TopMargin = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "BottomMargin"

                                Try

                                    ST.PageSetup.BottomMargin = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "LeftHeader"

                                Try

                                    ST.PageSetup.LeftHeader = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "CenterHeader"

                                Try

                                    ST.PageSetup.CenterHeader = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "RightHeader"

                                Try

                                    ST.PageSetup.RightHeader = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "LeftFooter"

                                Try

                                    ST.PageSetup.LeftFooter = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                            Case "RightFooter"

                                Try

                                    ST.PageSetup.RightFooter = .Item(i).Item(1)

                                Catch ex As Exception
                                    Method.ExceptionWrite(ex)
                                End Try

                        End Select

                    End If

                Next

            End With

            ''----------------------------------------------


            'With ST.PageSetup
            '    .LeftHeader = ""
            '.CenterHeader = "第 &P 页，共 &N 页"
            '.RightHeader = ""
            '.LeftFooter = ""
            '.CenterFooter = ""
            '.RightFooter = ""
            '    .LeftMargin = ST.Application.InchesToPoints(0.31496062992126)
            '    .RightMargin = ST.Application.InchesToPoints(0.31496062992126)
            '    .TopMargin = ST.Application.InchesToPoints(0.590551181102362)
            '    .BottomMargin = ST.Application.InchesToPoints(0.708661417322835)
            '    .HeaderMargin = ST.Application.InchesToPoints(0.31496062992126)
            '    .FooterMargin = ST.Application.InchesToPoints(0.31496062992126)
            '    .PrintHeadings = False
            '.PrintGridlines = False
            '    .PrintComments = XlPrintLocation.xlPrintNoComments
            '    .PrintQuality = 200
            '.CenterHorizontally = True
            '.CenterVertically = False
            '    .Orientation = XlPageOrientation.xlPortrait
            '    .Draft = False
            '    .PaperSize = XlPaperSize.xlPaperA4
            '    ' .FirstPageNumber = xlAutomatic
            '    .Order = XlOrder.xlDownThenOver
            '    .BlackAndWhite = False
            '.Zoom = False
            '.FitToPagesWide = 1
            '.FitToPagesTall = False
            '    .PrintErrors = XlPrintErrors.xlPrintErrorsDisplayed
            '    .OddAndEvenPagesHeaderFooter = False
            '.DifferentFirstPageHeaderFooter = False
            '.ScaleWithDocHeaderFooter = True
            '.AlignMarginsHeaderFooter = True
            '.EvenPage.LeftHeader.Text = ""
            '.EvenPage.CenterHeader.Text = ""
            '.EvenPage.RightHeader.Text = ""
            '.EvenPage.LeftFooter.Text = ""
            '.EvenPage.CenterFooter.Text = ""
            '.EvenPage.RightFooter.Text = ""
            '.FirstPage.LeftHeader.Text = ""
            '.FirstPage.CenterHeader.Text = ""
            '.FirstPage.RightHeader.Text = ""
            '.FirstPage.LeftFooter.Text = ""
            '.FirstPage.CenterFooter.Text = ""
            '    .FirstPage.RightFooter.Text = ""
            'End With
            ST.Application.PrintCommunication = True
            ''----------------------------------------------
        Catch ex As Exception

            Method.ExceptionWrite(ex)

        End Try

    End Sub

    ''' <summary>
    ''' 得到数据库值
    ''' </summary>
    ''' <param name="DR"></param>
    ''' <returns></returns>
    Private Shared Function GetValue(ByVal DR As DataRow) As Hashtable

        Dim HasT As New Hashtable

        For i As Integer = 1 To DR.ItemArray.Length - 1

            If Not IsDBNull(DR.Item(i)) Then

                HasT.Add(DR.Table.Columns(i).ColumnName, DR.Item(i))

            End If

        Next

        Return HasT

    End Function

End Structure
''' <summary>
''' 公共方法
''' </summary>
Public Class Method
    ''' <summary>
    ''' 编码字典,列：AQ1:墙A区
    ''' </summary>
    Public Shared dpDic As New Dictionary(Of String, String)
    ''' <summary>
    ''' 编码字典
    ''' </summary>
    Public Shared serDic As New Dictionary(Of String, Dictionary(Of String, Object))
    ''' <summary>
    ''' 通过字典保存不同编码类型
    ''' </summary>
    Public Shared psDic As New Dictionary(Of String, String)
    ''' <summary>
    ''' 存储生产清单数据，结构为工作薄名称-多张工作表，每张工作表-工作表中编码，用于生产清单
    ''' </summary>
    Public Shared wbookDic As New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, Integer)))
    ''' <summary>
    ''' 存储打包清单数据，结构为工作薄名称-多张工作表，每张工作表-工作表中编码，用于打包清单
    ''' </summary>
    Public Shared DwbookDic As New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, Integer)))
    ''' <summary>
    ''' 保存填充信息
    ''' </summary>
    Public Shared HillDic As New Dictionary(Of String, String)
    ''' <summary>
    ''' 生产清单字典添加值
    ''' </summary>
    Public Shared Sub DicAdd(ByVal producSer As ProductionSer)

        Try

            Dim wb, sheetName As String

            If ProductionSer.PixName <> "" Then

                sheetName = ProductionSer.PixName
                If Not HillDic.ContainsKey(sheetName) Then HillDic.Add(sheetName, sheetName)

            Else
                sheetName = producSer.GetSerType
                If Not HillDic.ContainsKey(sheetName) Then HillDic.Add(sheetName, sheetName)
            End If

            If ProductionSer.PixName1 <> "" Then

                wb = ProductionSer.PixName1 & producSer.GetSerType
                If Not HillDic.ContainsKey(wb) Then HillDic.Add(wb, producSer.GetSerType)

            Else

                wb = producSer.GetSerType
                If Not HillDic.ContainsKey(wb) Then HillDic.Add(wb, producSer.GetSerType)
            End If

            If producSer.GetSerType.Contains("标准") Then
                sheetName = Exl.DataSet.Tables("表信息").Select("编码类型" & "='" & producSer.GetSerType & "'")(0)("SheetName").ToString
            End If

            If Not wbookDic.ContainsKey(wb) Then wbookDic.Add(wb, New Dictionary(Of String, Dictionary(Of String, Integer)))

            If Not wbookDic(wb).ContainsKey(sheetName) Then wbookDic(wb).Add(sheetName, New Dictionary(Of String, Integer))

            If Not wbookDic(wb)(sheetName).ContainsKey(producSer.GetSer) Then
                wbookDic(wb)(sheetName).Add(producSer.GetSer, producSer.ProNum)
            Else
                wbookDic(wb)(sheetName)(producSer.GetSer) = wbookDic(wb)(sheetName)(producSer.GetSer) + producSer.ProNum
            End If

        Catch ex As Exception
            ExceptionWrite(ex)
        End Try

    End Sub

    ''' <summary>
    ''' 标准率表
    ''' </summary>
    Public Shared Sub DicAdd2(ByVal producSer As ProductionSer)

        Try

            Dim wb, sheetName As String

            If Not producSer.GetSerType.Contains("铁") Then

                If ProductionSer.PixName1 <> "" Then

                    wb = ProductionSer.PixName1 & "标准非标对比表"
                    If Not HillDic.ContainsKey(wb) Then HillDic.Add(wb, "标准非标对比表")

                End If

                If producSer.GetSerType.Contains("标准") Then

                    sheetName = "标准板"

                Else
                    sheetName = "非标准板"

                End If

                If Not HillDic.ContainsKey(sheetName) Then HillDic.Add(sheetName, sheetName)

            Else

                wb = ProductionSer.PixName1 & "铁件清单"
                If Not HillDic.ContainsKey(wb) Then HillDic.Add(wb, "铁件清单")
                sheetName = "铁件"
                If Not HillDic.ContainsKey(sheetName) Then HillDic.Add(sheetName, sheetName)

            End If

            If Not wbookDic.ContainsKey(wb) Then wbookDic.Add(wb, New Dictionary(Of String, Dictionary(Of String, Integer)))

                If Not wbookDic(wb).ContainsKey(sheetName) Then wbookDic(wb).Add(sheetName, New Dictionary(Of String, Integer))

                If Not wbookDic(wb)(sheetName).ContainsKey(producSer.GetSer) Then
                    wbookDic(wb)(sheetName).Add(producSer.GetSer, producSer.ProNum)
                Else
                    wbookDic(wb)(sheetName)(producSer.GetSer) = wbookDic(wb)(sheetName)(producSer.GetSer) + producSer.ProNum
                End If


        Catch ex As Exception
            ExceptionWrite(ex)
        End Try

    End Sub
    ''' <summary>

    ''' <summary>
    ''' 根据编码得来的字典给工作表赋值
    ''' </summary>
    Public Shared Sub Eval()

        Try

            Dim index As Integer

            For Each wbookName As String In wbookDic.Keys

                Dim dr As DataRow()

                If HillDic.ContainsKey(wbookName) Then

                    dr = Exl.DataSet.Tables("表信息").Select("编码类型='" & HillDic(wbookName) & "'")

                End If

                If IsNothing(dr) Then Exit For

                index = 1

                Dim WBook As Workbook = Globals.ThisAddIn.Application.Workbooks.Add()

                While (WBook.Sheets.Count <wbookDic(wbookName).Keys.Count)
                    WBook.Sheets.Add()
                End While

                For Each ShName As String In wbookDic(wbookName).Keys

                    Try

                        Exl.Sheet1 = WBook.Sheets.Item(index)

                        Exl.Sheet.Name = ShName

                        Exl.GetSheetType(dr(0))

                        Exl.InsertTH()

                        If HillDic.ContainsKey(ShName) Then

                            Exl.RepValue(HillDic(ShName))

                        End If

                        If ShName.Contains("标准") Then Exl.RepValue()

                        Dim sRow As Integer = Exl.StaRow

                        For Each seriInfo As String In wbookDic(wbookName)(ShName).Keys

                            Exl.Sheet.Range("c" & sRow).Value = seriInfo

                            Exl.Sheet.Range(Exl.NumValue & sRow).Value = wbookDic(wbookName)(ShName)(seriInfo) '工作表赋值数量

                            If IsNothing(serDic(seriInfo)) Then sRow = sRow + 1 : Continue For

                            For Each colum As String In Exl.DicMap.Keys

                                If serDic(seriInfo).ContainsKey(Exl.DicMap(colum)) Then

                                    If Not serDic(seriInfo)(Exl.DicMap(colum)).ToString.Contains("=") Then

                                        Exl.Sheet.Range(colum & sRow).Value = serDic(seriInfo)(Exl.DicMap(colum))

                                    Else

                                        Exl.Sheet.Range(colum & sRow).Value = GetValue5(serDic(seriInfo)(Exl.DicMap(colum)), sRow) '值包含等号就将值替换当前单元格

                                    End If

                                End If

                            Next

                            sRow = sRow + 1

                        Next

                        Exl.SortCol(col_2:="B")

                        Exl.SumTo()

                        Exl.SheetFormat("c")

                        Exl.SetPrintFormat()

                        index = index + 1

                    Catch ex As Exception
                        ExceptionWrite(ex)
                    End Try

                Next

                For Each she As Worksheet In WBook.Sheets

                    If she.Name.Contains("Sheet") Then WBook.Sheets(she.Name).Delete

                Next

                WBook.SaveAs(Exl.FliePath & "\" & Exl.ProjectName & wbookName, XlFileFormat.xlOpenXMLWorkbook)

                WBook.Close()

            Next

        Catch ex As Exception
            ExceptionWrite(ex)
        End Try

    End Sub
    ''' <summary>
    ''' 打包清单字典添加值
    ''' </summary>
    Public Shared Sub DDicAdd(ByVal producSer As ProductionSer, ByVal p As String)

        Try

            Dim wb, sh As String : Dim nu As Integer

            If Not ProductionSer.PixCode.Contains("B") Then

                nu = producSer.DDProNum

                If producSer.GetSerType.Contains("铁") Then

                    If Regex.IsMatch(ProductionSer.PixName, "(节点|楼梯)") Then
                        wb = ProductionSer.PixName1 & ThisAddIn.PubDic("铁打包")
                        sh = ProductionSer.PixName
                        If Not HillDic.ContainsKey(sh) Then HillDic.Add(sh, sh)
                    Else
                        wb = ProductionSer.PixName1 & ThisAddIn.PubDic("铁打包")
                        sh = ProductionSer.PixName & "铁"
                        If Not HillDic.ContainsKey(sh) Then HillDic.Add(sh, ProductionSer.PixName)
                        If Not dpDic.ContainsKey(p) Then dpDic.Add(p, ProductionSer.PixName)
                    End If
                    If Not HillDic.ContainsKey(wb) Then HillDic.Add(wb, ThisAddIn.PubDic("铁打包"))

                Else

                    If Regex.IsMatch(ProductionSer.PixName, "(节点|楼梯)") Then

                        If Regex.IsMatch(producSer.GetSer, "LB[LH]\d{1,4}") Then '对应铝背楞情况

                            wb = ProductionSer.PixName1 & ThisAddIn.PubDic("铝背楞件免拼打包清单")
                            If Not HillDic.ContainsKey(wb) Then HillDic.Add(wb, ThisAddIn.PubDic("铝背楞件免拼打包清单"))
                        Else
                            wb = ProductionSer.PixName1 & ThisAddIn.PubDic("铝打包")
                            If Not HillDic.ContainsKey(wb) Then HillDic.Add(wb, ThisAddIn.PubDic("铝打包"))
                        End If
                        sh = ProductionSer.PixName
                        If Not HillDic.ContainsKey(sh) Then HillDic.Add(sh, sh)

                    ElseIf Regex.IsMatch(ProductionSer.PixName, "吊模") Then

                        If Regex.IsMatch(producSer.GetSer, "LB[LH]\d{1,4}") Then

                            wb = ProductionSer.PixName1 & ThisAddIn.PubDic("铝背楞件免拼打包清单")
                            If Not HillDic.ContainsKey(wb) Then HillDic.Add(wb, ThisAddIn.PubDic("铝背楞件免拼打包清单"))
                        Else

                            wb = ProductionSer.PixName1 & ThisAddIn.PubDic("铝打包")
                            If Not HillDic.ContainsKey(wb) Then HillDic.Add(wb, ThisAddIn.PubDic("铝打包"))
                        End If
                        sh = ProductionSer.PixName & "铝"
                        If Not HillDic.ContainsKey(sh) Then HillDic.Add(sh, ProductionSer.PixName)

                    Else

                        If Regex.IsMatch(producSer.GetSer, "LB[LH]\d{1,4}") Then

                            wb = ProductionSer.PixName1 & ThisAddIn.PubDic("铝背楞件免拼打包清单")
                            If Not HillDic.ContainsKey(wb) Then HillDic.Add(wb, ThisAddIn.PubDic("铝背楞件免拼打包清单"))
                        Else

                            wb = ProductionSer.PixName & ThisAddIn.PubDic("铝打包")
                            If Not HillDic.ContainsKey(wb) Then HillDic.Add(wb, ThisAddIn.PubDic("铝打包"))
                        End If
                        sh = p
                        If Not HillDic.ContainsKey(sh) Then HillDic.Add(sh, ProductionSer.PixName & p)
                        If Not dpDic.ContainsKey(p) Then dpDic.Add(p, ProductionSer.PixName)

                    End If

                End If

            Else

                nu = producSer.ProNum

                Dim shh As String

                If producSer.GetSerType.Contains("铁") Then

                    wb = ProductionSer.PixName1 & ThisAddIn.PubDic("铁变层打包")

                    If Not HillDic.ContainsKey(wb) Then HillDic.Add(wb, ThisAddIn.PubDic("铁变层打包"))

                    shh = ProductionSer.PixName.Replace(ProductionSer.PixName1, "")

                    sh = ProductionSer.PixName1 & "(" & shh & ")"
                    '   sh = ProductionSer.PixName1 & "(" & shh.Replace("F-", "~") & ")"

                    If Not HillDic.ContainsKey(sh) Then HillDic.Add(sh, sh)

                Else

                    If Regex.IsMatch(producSer.GetSer, "LB[LH]\d{1,4}") Then

                        wb = ProductionSer.PixName1 & ThisAddIn.PubDic("铝背楞件变层打包清单")
                        If Not HillDic.ContainsKey(wb) Then HillDic.Add(wb, ThisAddIn.PubDic("铝背楞件变层打包清单"))
                    Else

                        wb = ProductionSer.PixName1 & ThisAddIn.PubDic("铝变层打包")
                        If Not HillDic.ContainsKey(wb) Then HillDic.Add(wb, ThisAddIn.PubDic("铝变层打包"))

                    End If

                    shh = ProductionSer.PixName.Replace(ProductionSer.PixName1, "")

                    sh = ProductionSer.PixName1 & "(" & shh & ")"

                    '  sh = ProductionSer.PixName1 & "(" & shh.Replace("F-", "~") & ")"

                    If Not HillDic.ContainsKey(sh) Then HillDic.Add(sh, sh)

                End If

            End If

            If Not DwbookDic.ContainsKey(wb) Then DwbookDic.Add(wb, New Dictionary(Of String, Dictionary(Of String, Integer)))

            If Not DwbookDic(wb).ContainsKey(sh) Then DwbookDic(wb).Add(sh, New Dictionary(Of String, Integer))

            Dim DSer As String = producSer.GetSer

            If Not DwbookDic(wb)(sh).ContainsKey(DSer) Then
                DwbookDic(wb)(sh).Add(DSer, nu)
            Else
                DwbookDic(wb)(sh)(DSer) = DwbookDic(wb)(sh)(DSer) + nu
            End If

        Catch ex As Exception
            ExceptionWrite(ex)
        End Try

    End Sub
    ''' <summary>
    ''' 打包备用清单字典添加值
    ''' </summary>
    Public Shared Sub BDDicAdd(ByVal producSer As ProductionSer)

        Try
            If producSer.ProNum <> producSer.DProNum Then

                Dim sheetName As String
                Dim wb As String

                wb = ProductionSer.PixName1 & ThisAddIn.PubDic("备用件免拼打包清单")
                sheetName = ProductionSer.PixName1
                If Not HillDic.ContainsKey(wb) Then HillDic.Add(wb, ThisAddIn.PubDic("备用件免拼打包清单"))
                If Not HillDic.ContainsKey(sheetName) Then HillDic.Add(sheetName, sheetName)

                If Not DwbookDic.ContainsKey(wb) Then DwbookDic.Add(wb, New Dictionary(Of String, Dictionary(Of String, Integer)))

                If Not DwbookDic(wb).ContainsKey(sheetName) Then DwbookDic(wb).Add(sheetName, New Dictionary(Of String, Integer))

                If Not DwbookDic(wb)(sheetName).ContainsKey(producSer.GetSer) Then
                    DwbookDic(wb)(sheetName).Add(producSer.GetSer, producSer.DProNum)
                Else
                    DwbookDic(wb)(sheetName)(producSer.GetSer) = DwbookDic(wb)(sheetName)(producSer.GetSer) + producSer.DProNum
                End If

            End If
        Catch ex As Exception
            ExceptionWrite(ex)
        End Try

    End Sub
    ''' <summary>
    ''' 将字典中值赋值到工作表
    ''' </summary>
    Public Shared Sub DEval()

        Try

            Dim index As Integer

            For Each wName As String In DwbookDic.Keys

                Dim dr As DataRow()

                If HillDic.ContainsKey(wName) Then
                    dr = Exl.DataSet.Tables("表信息").Select("编码类型='" & HillDic(wName) & "'")
                End If

                If IsNothing(dr) Then Exit For

                index = 1

                Dim Dwbook As Workbook = Globals.ThisAddIn.Application.Workbooks.Add()

                While (Dwbook.Sheets.Count < DwbookDic(wName).Keys.Count)
                    Dwbook.Sheets.Add()
                End While

                For Each DInfo As String In DwbookDic(wName).Keys

                    Try

                        Exl.Sheet1 = Dwbook.Sheets.Item(index)

                        Exl.Sheet.Name = DInfo

                        Exl.GetSheetType(dr(0))

                        Exl.InsertTH()

                        If HillDic.ContainsKey(DInfo） Then
                            Exl.DRepValue(HillDic(DInfo))
                        End If

                        Dim DRow As Integer = Exl.StaRow

                        For Each Dseri As String In DwbookDic(wName)(DInfo).Keys

                            Exl.Sheet.Range("B" & DRow).Value = Dseri '工作表赋值

                            Exl.Sheet.Range("c" & DRow).Value = DwbookDic(wName)(DInfo)(Dseri) '工作表赋值数量

                            DRow = DRow + 1

                        Next

                        Exl.SortCol(col_1:="b")

                        Exl.SumTo()

                        Exl.SheetFormat()

                        Exl.SetPrintFormat()

                        index = index + 1

                    Catch ex As Exception
                        ExceptionWrite(ex)
                    End Try

                Next

                Dwbook.SaveAs(Exl.FliePath & "\" & Exl.ProjectName & wName, XlFileFormat.xlOpenXMLWorkbook)

                Dwbook.Close()

            Next

        Catch ex As Exception
            ExceptionWrite(ex)
        End Try

    End Sub

    ''' <summary>
    ''' 得到选择的文件夹
    ''' </summary>
    ''' <param name="FList">返回文件地址信息</param>
    ''' <param name="projectName">返回项目名称</param>
    ''' <param name="path">返回文件夹路径</param>
    ''' <param name="pattern">筛选文件的匹配式</param>
    ''' <returns></returns>
    Shared Function SelFile(ByRef FList As List(Of IO.FileInfo), ByRef projectName As String, ByRef path As String, ByVal pattern As String) As Boolean

        Try

            Dim FolderDialog As New FolderBrowserDialog With {.Description = "选择清单所在的文件夹"}

        If DialogResult.OK = FolderDialog.ShowDialog Then

            path = FolderDialog.SelectedPath

            Dim Di As DirectoryInfo = New DirectoryInfo(Exl.FliePath) : projectName = Di.Name

            Dim AdsList As IO.FileInfo() = Di.GetFiles("*.xl*")

            For Each Address As IO.FileInfo In AdsList

                If Regex.IsMatch(Address.Name.Substring(0, Address.Name.LastIndexOf(".")), pattern:=pattern) Then FList.Add(Address)

            Next

            If Not IsNothing(FList) AndAlso FList.Count > 0 Then Return True

        End If

        Return False

        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try

    End Function
    ''' <summary>
    ''' 当数据表对应列值带=时计算，如=G*F/1000形式，将字母替换成对应单元格
    ''' </summary>
    ''' <param name="Value">对应数据表列值</param>
    ''' <param name="Row">单元格行</param>
    ''' <returns></returns>
    Public Shared Function GetValue5(ByVal Value As Object, ByVal Row As Integer) As String

        Dim MCollection As MatchCollection = Regex.Matches(Value, "[A-Z]{1,2}")

        Dim Arr As New HashSet(Of String)

        For j As Integer = 0 To MCollection.Count - 1

            Arr.Add(MCollection(j).Value)

        Next

        For i As Integer = 0 To Arr.Count - 1

            Value = Regex.Replace(Value, pattern:=Arr(i), replacement:=Arr(i) & Row)

        Next

        Return Value

    End Function
    ''' <summary>
    ''' 将异常信息写入文本
    ''' </summary>
    ''' <param name="ex">异常</param>
    Public Shared Sub ExceptionWrite(ByVal ex As Exception)

        Dim path As String = "c:\ExceptionLog.txt"
        Dim fs As FileStream
        Dim sw As StreamWriter
        Dim fl As Long

        If File.Exists(path) Then

            fs = New FileStream(path, FileMode.Open, FileAccess.Write)
            sw = New StreamWriter(fs, UTF8Encoding.Unicode)
            fl = fs.Length
            fs.Seek(fl, SeekOrigin.Begin)

        Else

            fs = New FileStream(path, FileMode.Create, FileAccess.Write)
            sw = New StreamWriter(fs, UTF8Encoding.Unicode)
            fl = fs.Length
            fs.Seek(fl, SeekOrigin.End)

        End If

        sw.WriteLine(DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss") + ":")
        sw.WriteLine(ex.Message + ex.StackTrace)
        sw.WriteLine()

        sw.Close()

        fs.Close()

    End Sub
    ''' <summary>
    ''' 加密
    ''' </summary>
    ''' <param name="SourceStr">需要加密的字符串</param>
    ''' <param name="myKey">加密的8个任意字符</param>
    ''' <param name="myIV">密码本8位数字</param>
    ''' <returns></returns>
    Shared Function EncryptDes1(ByVal SourceStr As String, ByVal myKey As String, ByVal myIV As String) As String '使用的DES对称加密  

        Dim des As New System.Security.Cryptography.DESCryptoServiceProvider 'DES算法  
        'Dim DES As New System.Security.Cryptography.TripleDESCryptoServiceProvider'TripleDES算法  
        Dim inputByteArray As Byte()
        inputByteArray = System.Text.Encoding.Default.GetBytes(SourceStr)
        des.Key = System.Text.Encoding.UTF8.GetBytes(myKey) 'myKey DES用8个字符，TripleDES要24个字符  
        des.IV = System.Text.Encoding.UTF8.GetBytes(myIV) 'myIV DES用8个字符，TripleDES要24个字符  
        Dim ms As New System.IO.MemoryStream
        Dim cs As New System.Security.Cryptography.CryptoStream(ms, des.CreateEncryptor(), System.Security.Cryptography.CryptoStreamMode.Write)
        Dim sw As New System.IO.StreamWriter(cs)
        sw.Write(SourceStr)
        sw.Flush()
        cs.FlushFinalBlock()
        ms.Flush()
        EncryptDes1 = Convert.ToBase64String(ms.GetBuffer(), 0, ms.Length)

    End Function
End Class
''' <summary>
''' 编码类
''' </summary>
Public Class ProductionSer
    ''' <summary>
    ''' 编码
    ''' </summary>
    Private _s As String
    ''' <summary>
    ''' 去掉前缀的编码
    ''' </summary>
    Private s As String
    ''' <summary>
    ''' 编码数量
    ''' </summary>
    Private _num As Integer
    ''' <summary>
    ''' 编码数量,等于数量列值,如数量列为1,则为1
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property DDProNum() As String
        Get
            Return _num
        End Get
    End Property
    ''' <summary>
    ''' 生产编码数量
    ''' </summary>
    Private pCount As Integer
    ''' <summary>
    ''' 生产编码数量,乘以编码特殊标识中的值,如Z[3],则为：数量列X3
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property ProNum() As String
        Get
            Return pCount
        End Get
    End Property
    ''' <summary>
    ''' 打包编码数量
    ''' </summary>
    Private dCount As Integer
    ''' <summary>
    ''' 打包编码数量,乘以编码特殊标识中的值减去1,如Z[3],则为：数量列X(3-1)
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property DProNum() As String
        Get
            Return dCount
        End Get
    End Property
    ''' <summary>
    '''编码中提取的字符编码
    ''' </summary>
    Private cd As String
    ''' <summary>
    '''编码类型
    ''' </summary>
    Private sType As String
    ''' <summary>
    '''编码配件pattern
    ''' </summary>
    Private sAc As String
    ''' <param name="initiSer">初始编码</param>
    ''' <summary>
    ''' 编码格式
    ''' </summary>
    Private ft As String
    ''' <summary>
    ''' 部位分区，如墙
    ''' </summary>
    Private Shared pName1 As String
    ''' <summary>
    ''' 部位分区，如墙
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property PixName1() As String
        Get
            Return pName1
        End Get
    End Property
    ''' <summary>
    ''' 部位分区，如墙A区
    ''' </summary>
    Private Shared pName As String
    ''' <summary>
    ''' 部位分区，如墙A区
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property PixName() As String
        Get
            Return pName
        End Get
    End Property
    ''' <summary>
    ''' 部位分区，如L、Q等
    ''' </summary>
    Private Shared pCode As String
    ''' <summary>
    ''' 部位分区，如L、Q等
    ''' </summary>
    Public Shared ReadOnly Property PixCode() As String
        Get
            Return pCode
        End Get
    End Property
    ''' <summary>
    ''' 编码中对应的值
    ''' </summary>
    Private sDic As New Dictionary(Of String, Object)
    ''' <summary>
    ''' 编码中对应的值
    ''' </summary>
    Public ReadOnly Property SerDic() As Dictionary(Of String, Object)
        Get
            Return sDic
        End Get
    End Property
    ''' <summary>
    '''编码类型
    ''' </summary>
    Public ReadOnly Property GetSerType() As String
        Get
            Return sType
        End Get
    End Property
    ''' <summary>
    '''编码
    ''' </summary>
    Public ReadOnly Property GetSer() As String
        Get
            Return _s
        End Get
    End Property
    ''' <summary>
    ''' 初始化主编码
    ''' </summary>
    ''' <param name="seri">主编码</param>
    ''' <param name="num">编码数量</param>
    Sub New(ByVal seri As String, Optional ByVal num As Integer = 0)

        _num = num

        SerProcess(seri, num)

        SerType()

    End Sub
    ''' <summary>
    ''' 编码加工
    ''' </summary>
    ''' <param name="initiSer">初始编码</param>
    Private Sub SerProcess(ByVal initiSer As String, Optional ByVal num As Integer = 0)
        Try

            s = Regex.Replace(initiSer, "\s", "").ToUpper

            _s = s

            Dim tao As String = "" : Dim tao1 As String = "" : Dim increaseNum As Integer = 0

            If num <> 0 Then tao = Regex.Match(s, pattern:=ThisAddIn.PubDic("备用标识")).Value
            If num <> 0 Then tao1 = Regex.Match(s, "\[2\]$").Value

            If tao = "" Then

                pCount = num : dCount = num

            Else

                s = Regex.Replace(s, pattern:=ThisAddIn.PubDic("备用标识"), replacement:="")

                _s = s

                increaseNum = Regex.Match(tao, pattern:="\d{1,2}").Value

                If increaseNum > 2 Then
                    pCount = num * increaseNum : dCount = num * (increaseNum - 1)
                ElseIf increaseNum = 2 AndAlso tao1 <> "" Then
                    pCount = num * increaseNum : dCount = num * (increaseNum - 1)
                Else
                    pCount = num * increaseNum : dCount = num * increaseNum : _num = num * increaseNum
                End If

            End If

            ft = Regex.Replace(s, pattern:="\d{1,4}\.?\d*", replacement:="#")

            cd = Regex.Match(ft, pattern:="[A-Z]+").Value

        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try
    End Sub
    ''' <summary>
    ''' 得到部位分区信息
    ''' </summary>
    Public Shared Sub GetPix(ByVal p As String)

        Try

            ClearPix()

            Dim pT As String = Regex.Replace(p, "\s", "").ToUpper

            Dim Pi As String = Regex.Match(pT, "[DJLMTQ](?=\d|\()").Value

            Dim DR As DataRow() = Exl.DataSet.Tables("部位信息").Select("P='" & Pi & "'")

            For i As Integer = 0 To DR.Length - 1

                If Regex.IsMatch(p, pattern:=DR(i)("pattern")) Then

                    pCode = DR(i)("Pcode").ToString

                    If Not pCode.Contains("B") Then

                        Dim gr As GroupCollection = Regex.Match(p, pattern:=DR(i)("pattern")).Groups

                        pName1 = DR(i)("Pname1").ToString

                        pName = DR(i)("Pname").ToString.Replace("#", gr(1).Value)

                    Else

                        Dim pN As Integer = Regex.Match(pT, "\d{1,2}(?=\()").Value

                        pCode = pCode & pN

                        pName1 = DR(i)("Pname1").ToString

                        Dim pp As String = Regex.Replace(pT, "[DJLMTQ]\d{0,2}(?=\()", replacement:="")

                        pName = pp.Replace("(", "(" & pName1)

                        pName = Regex.Replace(pName, pattern:="[\(\)]", replacement:="")

                    End If

                End If

            Next

        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try

    End Sub
    ''' <summary>
    ''' 编码值计算
    ''' </summary>
    Public Sub SerialCal()

        Try

            Dim BDo As Boolean = False '编码特殊处理方式
            Dim DR() As DataRow = Nothing

            Dim ser As String = s

            Dim mFt As String = Regex.Replace(ser, pattern:="\d{1,4}\.?\d*", replacement:="#")

            DR = Exl.DataSet.Tables("Format").Select("Format" & "='" & mFt & "'")

            If sAc <> "" AndAlso DR.Length = 0 Then
                ser = Regex.Replace(s, pattern:=sAc, replacement:="") '当存在配件时，去掉配件部分
                mFt = Regex.Replace(ser, pattern:="\d{1,4}\.?\d*", replacement:="#")
            End If

            Do

                DR = Exl.DataSet.Tables("Format").Select("Format" & "='" & mFt & "'")

                For i As Int16 = 0 To DR.Length - 1

                    If Regex.IsMatch(ser, DR(i)("Pattern").ToString) Then

                        GetValue(DR(i), Regex.Match(ser, DR(i)("Pattern").ToString).Groups, mFt)

                        Exit Do

                    End If

                Next

                If cd = "BL" AndAlso Regex.IsMatch(ser.Replace("BL", ""), "[A-Z]") Then '如果除背楞外存在其他字母

                    ser = Regex.Replace(ser, "BL", replacement:="%")
                    ser = Regex.Replace(ser, pattern:="[A-Z]+", replacement:="")
                    ser = Regex.Replace(ser, pattern:="%", replacement:="BL")

                ElseIf cd <> "BL" AndAlso ser.Contains("-") Then

                    ser = ser.Remove(ser.LastIndexOf("-"))

                Else

                    Exit Do

                End If

                mFt = Regex.Replace(ser, pattern:="\d{1,4}\.?\d*", replacement:="#")

            Loop

        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try

    End Sub
    ''' <summary>
    ''' 编码包含的数据添加入字典中
    ''' </summary>
    ''' <param name="DR">数据表中对应行</param>
    ''' <param name="mFt">处理后得到的格式，处理后得到的格式和未处理的格式比较得出正确图号</param>
    Private Sub GetValue(ByVal DR As DataRow, ByVal Group As GroupCollection, ByVal mFt As String)

        Try

            For i As Integer = 3 To DR.ItemArray.Length - 2  '数据库从第三列开始加入需要计算的列

                If Not IsDBNull(DR.Item(i)) Then

                    With DR.Table.Columns(i)

                        If Not .ColumnName.Contains("+") Then

                            If Not Regex.IsMatch(DR(i).ToString, "[!\+]") OrElse DR(i).ToString.Contains("=") Then

                                Dim Value As Object = Nothing

                                If Regex.IsMatch(DR(i).ToString, pattern:="^\d+$") AndAlso DR.Item(i) < Group.Count Then

                                    Value = Group(CInt(DR.Item(i))).Value

                                Else

                                    Value = DR(i)

                                End If

                                sDic.Add(.ColumnName, Value)

                                Continue For

                            ElseIf DR(i).ToString.Contains("!") Then

                                sDic.Add(.ColumnName, Exl.DataSet.Tables("CalSize").
                                             Select("Code" & "='" & Group(CInt(DR.Item(i).ToString.TrimEnd("!"))).
                                             Value & "'")(0)("Value"))

                                Continue For

                            ElseIf DR(i).ToString.Contains("+") AndAlso Not DR(i).ToString.Contains("=") Then

                                Dim MidValue As Single = 0

                                For Each Str As String In DR.Item(i).ToString.Split("+")

                                    MidValue = Group(CInt(Str)).Value + MidValue

                                Next

                                sDic.Add(.ColumnName, MidValue)

                                Continue For

                            End If

                        Else

                            sDic(.ColumnName.TrimStart("+")) = sDic(.ColumnName.TrimStart("+")) + DR(i) : Continue For

                        End If

                    End With

                End If

            Next

            If sType = ThisAddIn.PubDic("标准件清单") Then
                sDic.Add("Mark", "标准件")
            ElseIf ft = mFt AndAlso Not IsDBNull(DR("Mark")) Then
                sDic.Add("Mark", DR("Mark"))
            End If

        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try

    End Sub
    ''' <summary>
    ''' 得到编码类型
    ''' </summary>
    Private Sub SerType()

        Try

            Dim midS As String

            If Regex.IsMatch(cd, "LB[LH]") Then

                If Not IsNothing(PixCode) AndAlso PixCode.Contains("B") Then

                    sType = ThisAddIn.PubDic("铝背楞变层清单")

                Else

                    sType = ThisAddIn.PubDic("铝背楞清单")

                End If

                Dim aa As String = Regex.Replace(s, "[A-Z]+", replacement:="")

                If Nor(cd, aa) = False AndAlso PixCode <> "" Then
                    _s = PixCode & "-" & s
                End If

                Exit Sub

            End If

            If Nor(cd, s) Then

                sType = ThisAddIn.PubDic("标准件清单")

            ElseIf Exl.DataSet.Tables("铝铁件").Select("铁制件" & "='" & cd & "'").Count > 0 Then

                If PixCode <> "" Then _s = PixCode & "-" & s

                If Not IsNothing(PixCode) AndAlso PixCode.Contains("B") Then

                    sType = ThisAddIn.PubDic("铁制件变层清单")

                Else

                    sType = ThisAddIn.PubDic("铁制件清单")

                End If

            Else

                Dim DR As DataRow() = Exl.DataSet.Tables("判断配件").Select("Code='" & cd & "'")

                If DR.Length > 0 Then
#Region “穿墙孔范围”
                    Dim Kpar As Integer = 0
                    If cd <> "P" Then Kpar = 1
                    If cd = "P" AndAlso Regex.IsMatch(s, pattern:=DR(0)("PatternK")) Then
                        Dim wit As Integer = Regex.Match(s, pattern:="^\d{1,3}").Value

                        If ThisAddIn.PubDic.ContainsKey(wit) Then

                            Dim mis As String = Regex.Match(s, "-\d{1,3}K").Value

                            Dim w1 As Integer = Regex.Match(mis, pattern:="\d{1,3}").Value

                            If ThisAddIn.PubDic(wit).Contains(",") Then

                                Dim fd As String() = ThisAddIn.PubDic(wit).Split(",")

                                For Each f As String In fd

                                    Dim fw As String() = f.Split("-")

                                    If w1 > fw(0) AndAlso w1 < fw(1) Then

                                        Kpar = 1

                                    End If

                                Next
                            Else
                                Dim fd As String() = ThisAddIn.PubDic(wit).Split("-")
                                If w1 > fd(0) AndAlso w1 < fd(1) Then

                                    Kpar = 1

                                End If
                            End If

                        End If

                    End If
#End Region
                    If Not IsDBNull(DR(0)("Pattern1")) AndAlso Regex.IsMatch(s, pattern:=DR(0)("Pattern1")) Then '确定带配件

                        sAc = DR(0)("Pattern1") : midS = Regex.Replace(s, pattern:=sAc, replacement:="") '将配件匹配式保留，并将编码中配件部分替换为空

                        If s.Contains("-") And Nor(cd, midS) Then '替换后编码标准件判断

                            sType = ThisAddIn.PubDic("带配件标准件清单")

                        ElseIf Not IsDBNull(DR(0)("PatternK")) AndAlso Regex.IsMatch(midS, pattern:=DR(0)("PatternK")) AndAlso
                                     Nor(cd, Regex.Replace(midS, pattern:=DR(0)("PatternK"), replacement:="")) AndAlso Kpar = 1 Then '判断是否存在穿墙孔，将穿墙孔部分替换为空，再判断标准件

                            sType = ThisAddIn.PubDic("标准板带配件及穿墙孔清单")

                        Else

                            If PixCode <> "" Then _s = PixCode & "-" & s

                            If Not IsNothing(PixCode) AndAlso PixCode.Contains("B") Then

                                sType = ThisAddIn.PubDic("铝制件带配件变层清单")

                            Else

                                sType = ThisAddIn.PubDic("铝制件带配件清单")

                            End If

                        End If

                    ElseIf Not IsDBNull(DR(0)("PatternK")) AndAlso Regex.IsMatch(s, pattern:=DR(0)("PatternK")) AndAlso
                    Nor(cd, Regex.Replace(s, pattern:=DR(0)("PatternK"), replacement:="")) AndAlso Kpar = 1 Then '判断是否存在穿墙孔，将穿墙孔部分替换为空，再判断标准件

                        sType = ThisAddIn.PubDic("标准板带穿墙孔清单")

                    Else

                        If PixCode <> "" Then _s = PixCode & "-" & s

                        If Not IsNothing(PixCode) AndAlso PixCode.Contains("B") Then

                            sType = ThisAddIn.PubDic("铝制件变层清单")

                        Else

                            sType = ThisAddIn.PubDic("铝制件清单")

                        End If

                    End If

                Else

                    If PixCode <> "" Then _s = PixCode & "-" & s

                    If Not IsNothing(PixCode) AndAlso PixCode.Contains("B") Then

                        sType = ThisAddIn.PubDic("铝制件变层清单")

                    Else

                        sType = ThisAddIn.PubDic("铝制件清单")

                    End If

                End If

            End If

        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try

    End Sub
    ''' <summary>
    ''' 标准件判断
    ''' </summary>
    ''' <returns></returns>
    Private Function Nor(ByVal cd As String, ByVal ser As String) As Boolean

        If Exl.DataSet.Tables("标准件表").Columns.Contains(cd) AndAlso
           Exl.DataSet.Tables("标准件表").Select(cd & "='" & ser & "'").Count > 0 Then Return True

    End Function
    ''' <summary>
    ''' 初始化静态变量
    ''' </summary>
    Public Shared Sub ClearPix()

        pName = ""
        pName1 = ""
        pCode = ""

    End Sub
    ''' <summary>
    ''' 去掉前缀
    ''' </summary>
    ''' <param name="S">编码</param>
    ''' <returns></returns>
    Shared Function MinusStaCode(ByVal S As String) As String

        If S <> "" AndAlso S.Contains("-") Then

            Dim CodeID As String = Regex.Match(S, pattern:="[A-Z]+").Value

            If FirstCodeVerdict(CodeID, S.Substring(0, S.IndexOf("-"))) Then   '''如果第一段编码不属于铝铁件编码，则去掉前缀；如果属于铝铁件编码，但前缀不包含任何数字，则当前缀去掉。
                Return S.Substring(S.IndexOf("-") + 1)
            Else
                Return S
            End If

        Else

            Return S

        End If

    End Function
    ''' <summary>
    ''' 前缀判断
    ''' </summary>
    ''' <param name="firstCode">完整编码中第一段编码的字母</param>
    ''' <param name="firstCompleteCode">完整编码的第一段编码</param>
    ''' <returns></returns>
    Shared Function FirstCodeVerdict(ByVal firstCode As String, ByVal firstCompleteCode As String) As Boolean

        If firstCode = "" Then Return False

        Dim lm As Boolean = Exl.DataSet.Tables("铝铁件").Select("铝制件='" & firstCode & "'").Length = 0
        Dim Tm As Boolean = Exl.DataSet.Tables("铝铁件").Select("铁制件='" & firstCode & "'").Length = 0

        If lm AndAlso Tm Then Return True

        If (Not lm OrElse Not Tm) AndAlso Regex.IsMatch(firstCompleteCode, ThisAddIn.PubDic("前缀标识")) Then Return True

    End Function
End Class