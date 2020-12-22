Imports Microsoft.Office.Tools.Ribbon
Imports System.Data
Imports System.Data.OleDb
Imports System.Text.RegularExpressions
Public Class Ribbon2

    Private Sub Ribbon2_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

        Try

            '  If GetIsRegedit() Then

            Dim CN As New OleDbConnection

                Dim PateDB As String = "d:\3DConfiguration\ZWDB.accdb"

                Try

                    CN.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & PateDB
                    CN.Open()

                Catch ex As Exception

                    MsgBox("数据库未链接成功" & vbCr &
                           "1、请确认已安装好数据库驱动：AccessDatabaseEngine_X64" & vbCr &
                           "2、请确认已安装数据库文件" & PateDB)

                    Exit Sub

                End Try

                Dim DT As DataTable = CN.GetSchema("Tables")

                Exl.DataSet = New DataSet

                For Each Row As DataRow In DT.Rows

                    If Row(3).ToString() = "TABLE" Then

                        Dim Adapter As OleDbDataAdapter = New OleDbDataAdapter("select * from " & Row(2).ToString(), CN)

                        Adapter.Fill(Exl.DataSet, Row(2).ToString())

                    End If

                Next

                CN.Close()

            '   End If

            ThisAddIn.PubDic = New Dictionary(Of String, String)

            For r As Integer = 0 To Exl.DataSet.Tables("全局参数").Rows.Count - 1

                ThisAddIn.PubDic.Add(Exl.DataSet.Tables("全局参数").Rows.Item(r).Item(0), Exl.DataSet.Tables("全局参数").Rows.Item(r).Item(1))

            Next

        Catch ex As Exception
            Method.ExceptionWrite(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 插件注册
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Regist_Click(sender As Object, e As RibbonControlEventArgs) Handles Regist.Click

        Dim F As New Regedit

        F.Show()

    End Sub

    ''' <summary>
    ''' 插件说明
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Descripe_Click(sender As Object, e As RibbonControlEventArgs) Handles Descripe.Click

        '  If Not IsRegedit AndAlso Not GetIsRegedit() Then

        '     MsgBox("请先注册软件") : Exit Sub

        '  Else

        Dim V As Version = Reflection.Assembly.GetExecutingAssembly().GetName().Version
            Dim S As String = "软件版本：" & V.Major & "." & V.Minor & "." & V.Build & "." & V.Revision

            MsgBox(S & ",软件使用说明：" & Chr(13) & Chr(10) &
        "1、 在使用插件前，先备份原始表格，谨防出错；" & Chr(13) & Chr(10) &
        "2、 每个表单里面仅能有一张表，如果同时存在两张表，计算将会出错；" & Chr(13) & Chr(10) &
        "3、 插件所匹配的铝件大样图为：1491523013铝件大样图-7.19；" & Chr(13) & Chr(10) &
        "4、 插件所匹配的铁件大样图为：1491523013铁件大样图-7.11；" & Chr(13) & Chr(10) &
        "5、 编码列中，不能存在空单元格，否则将导致程序停止；" & Chr(13) & Chr(10) &
        "6、 导出的原始清单需严格按照程序识别的名称命名；" & Chr(13) & Chr(10) &
        "7、 原始清单A1单元格为‘编码’B1单元格为‘数量’；" & Chr(13) & Chr(10) &
        "8、 多表自动·二维清单 选择原始清单，编码格式：部位分区—编码，生成‘生产清单’和‘打包清单’； " & Chr(13) & Chr(10) &
        "9、 多表自动·生产清单 选择放已分区原始清单文件夹，生成‘生产清单’； " & Chr(13) & Chr(10) &
        "10、多表自动·打包清单 选择放已分区原始清单文件夹，生成‘打包清单’；" & Chr(13) & Chr(10) &
        "11、多表自动·销售清单 选择放生产清单的文件夹，文件夹名为项目名，生成‘销售清单’；" & Chr(13) & Chr(10) &
        "12、功能·清单计算 对当前清单进行计算（铝件生产清单和铁件生产清单）；" & Chr(13) & Chr(10) &
        "13、功能·清单分类 对当前清单的编码进行分类，按：铝件（标准和非标准）和铁件；" & Chr(13) & Chr(10) &
        "14、功能·编码合并 对当前清单相同的编码进行合并；" & Chr(13) & Chr(10) &
        "15、功能·插入表头 在当前表中自动插入生产清单表头或者打包清单表头；" & Chr(13) & Chr(10) &
        "16、功能·编码检查 对当前清单的编码进行规范检查；" & Chr(13) & Chr(10) &
        "17、功能·格式设置 对当前清单进行打印格式设置；" & Chr(13) & Chr(10) &
        "18、功能·添加前缀 可以仅对当前表单添加，或者对整个文件添加；" & Chr(13) & Chr(10) &
        "19、功能·去除前缀 可以仅对当前表单去除，或者对整个文件去除。"）

        ' End If

    End Sub

    ''' <summary>
    ''' 添加表头
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub InsertTable_Click(sender As Object, e As RibbonControlEventArgs) Handles InsertTable.Click

        '  If Not IsRegedit AndAlso Not GetIsRegedit() Then

        '    MsgBox("请先注册软件") : Exit Sub

        '  Else

        Try

                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.DisplayAlerts = False
                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook
                Exl.Sheet1 = Exl.WorkBook.ActiveSheet

                If Not IsNothing(Exl.Sheet1.Range("a1").Value) Then Exit Sub

                Select Case InputBox("按序号选择需要添加的表头：" & vbCrLf &
                                           "1、标准件清单表头" & vbCrLf &
                                           "2、标准件带配件清单表头" & vbCrLf &
                                           "3、非标件清单表头" & vbCrLf &
                                           "4、非标件带配件清单表头" & vbCrLf &
                                           "5、铁件清单表头" & vbCrLf &
                                           "6、打包清单表头")

                    Case 1 : Exl.GetSheetType(Exl.DataSet.Tables("表信息").Select("headerName='" & "标准板明细清单" & "'")(0))
                    Case 2 : Exl.GetSheetType(Exl.DataSet.Tables("表信息").Select("headerName='" & "标准板带配件明细清单" & "'")(0))
                    Case 3 : Exl.GetSheetType(Exl.DataSet.Tables("表信息").Select("headerName='" & "非标准板明细清单" & "'")(0))
                    Case 4 : Exl.GetSheetType(Exl.DataSet.Tables("表信息").Select("headerName='" & "非标准板带配件明细清单" & "'")(0))
                    Case 5 : Exl.GetSheetType(Exl.DataSet.Tables("表信息").Select("headerName='" & "铁件明细清单" & "'")(0))
                Case 6 : Exl.GetSheetType(Exl.DataSet.Tables("表信息").Select("headerName='" & "打包清单" & "'")(0))

                Case Else : Exit Try

                End Select

                Exl.InsertTH()

                Exl.SheetFormat()

                Exl.SetPrintFormat()
            Catch ex As Exception
                Method.ExceptionWrite(ex)
            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True

            End Try

     '   End If

    End Sub

    ''' <summary>
    ''' 编码检查
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CodeCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles CodeCheck.Click

        ' If Not IsRegedit AndAlso Not GetIsRegedit() Then

        '      MsgBox("请先注册软件")

        '  Else

        Dim Bool As MsgBoxResult = MsgBox("是否检查所有工作表编码？", MsgBoxStyle.YesNoCancel)
            Dim Count As Short = 1

            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.DisplayAlerts = False

            Try

                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook

                Select Case Bool

                    Case MsgBoxResult.Yes : Count = Exl.WorkBook.Sheets.Count

                    Case MsgBoxResult.Cancel : Exit Sub

                End Select

                For i As Integer = 1 To Count

                    If Count <> 1 Then

                        Exl.Sheet = Exl.WorkBook.Sheets.Item(i)

                    Else

                        Exl.Sheet = Exl.WorkBook.ActiveSheet

                    End If

                    CheckSerial()

                Next

            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True

            End Try

        'End If

    End Sub

    ''' <summary>
    ''' 打印格式设置
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub PrintFormat_Click(sender As Object, e As RibbonControlEventArgs) Handles PrintFormat.Click

        '  If Not IsRegedit AndAlso Not GetIsRegedit() Then

        '     MsgBox("请先注册软件") : Exit Sub

        '   Else

        Dim Bool As MsgBoxResult = MsgBox("是否所有工作表设置打印格式？", MsgBoxStyle.YesNoCancel)
            Dim Count As Short = 1

            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.DisplayAlerts = False

            Try

                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook

                Select Case Bool

                    Case MsgBoxResult.Yes : Count = Exl.WorkBook.Sheets.Count

                    Case MsgBoxResult.Cancel : Exit Sub

                End Select

                For i As Integer = 1 To Count

                    If Count <> 1 Then

                        Exl.Sheet = Exl.WorkBook.Sheets.Item(i)

                    Else

                        Exl.Sheet = Exl.WorkBook.ActiveSheet

                    End If

                    If IsNothing(Exl.Sheet.Range("a1").Value) Then Continue For

                    Exl.SumTo()

                    If Exl.BomFormat.Contains("打包") Then Exl.SheetFormat() Else Exl.SheetFormat("c")

                    Exl.SetPrintFormat()

                Next
            Catch ex As Exception
                Method.ExceptionWrite(ex)
            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True
                ' Globals.ThisAddIn.Application.ScreenUpdating = True
            End Try

     '   End If

    End Sub

    ''' <summary>
    ''' 添加前缀
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AddItemCode_Click(sender As Object, e As RibbonControlEventArgs) Handles AddItemCode.Click

        '  If Not IsRegedit AndAlso Not GetIsRegedit() Then

        '      MsgBox("请先注册软件")

        '  Else

        Dim Input As String = InputBox("请输入：", "输入框")
            Dim Bool As MsgBoxResult = MsgBox("是否所有工作表添加前缀？" & vbCr & "输入的前缀为：" & Input, MsgBoxStyle.YesNoCancel)
            Dim Count As Short = 1

            If Input = "" Then Exit Sub

            Input = Input.ToUpper

            If Not Input.EndsWith("-") Then Input &= "-"

            Try

                Globals.ThisAddIn.Application.ScreenUpdating = False
                Globals.ThisAddIn.Application.DisplayAlerts = False

                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook

                Select Case Bool

                    Case MsgBoxResult.Yes : Count = Exl.WorkBook.Sheets.Count

                    Case MsgBoxResult.Cancel : Exit Sub

                End Select

                For i As Integer = 1 To Count

                    If Count <> 1 Then

                        Exl.Sheet = Exl.WorkBook.Sheets.Item(i)

                    Else

                        Exl.Sheet = Exl.WorkBook.ActiveSheet

                    End If

                    If Exl.StaRow > 0 Then

                        Dim Col As String = "C"

                        If Exl.StaRow = 3 Then Col = "B"

                        Dim producSer As ProductionSer

                        Dim value As String

                        For j As Integer = Exl.StaRow To Exl.EndRow(Col)

                            value = Exl.Sheet.Range(Col & j).Value

                            If IsNothing(value) Then Continue For

                            If value.Contains("-") Then

                                Dim pix As String = value.Substring(0, value.IndexOf("-")) : Dim pixCode As String = Regex.Match(pix, pattern:="[A-Z]+").Value

                                If ProductionSer.FirstCodeVerdict(pixCode, pix) Then

                                    Continue For

                                Else

                                    producSer = New ProductionSer(value)

                                    If Not producSer.GetSerType.Contains("标准") Then Exl.Sheet.Range(Col & j).Value = Input & value

                                End If

                            Else

                                producSer = New ProductionSer(value)

                                If Not producSer.GetSerType.Contains("标准") Then Exl.Sheet.Range(Col & j).Value = Input & value

                            End If

                        Next

                        CheckSerial()

                    End If

                Next
            Catch ex As Exception
                Method.ExceptionWrite(ex)
            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True

            End Try

       ' End If

    End Sub

    ''' <summary>
    ''' 去除前缀
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CancelItemCode_Click(sender As Object, e As RibbonControlEventArgs) Handles CancelItemCode.Click

        '    If Not IsRegedit AndAlso Not GetIsRegedit() Then

        '      MsgBox("请先注册软件")

        '  Else

        Dim Bool As MsgBoxResult = MsgBox("是否所有工作表取消前缀？", MsgBoxStyle.YesNoCancel)
            Dim Count As Short = 1

            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.DisplayAlerts = False

            Try

                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook

                Select Case Bool

                    Case MsgBoxResult.Yes : Count = Exl.WorkBook.Sheets.Count

                    Case MsgBoxResult.Cancel : Exit Sub

                End Select

                For i As Integer = 1 To Count

                    If Count <> 1 Then

                        Exl.Sheet = Exl.WorkBook.Sheets.Item(i)

                    Else

                        Exl.Sheet = Exl.WorkBook.ActiveSheet

                    End If

                    If Exl.StaRow > 0 Then

                        Dim Col As String = "C"

                        If Exl.StaRow = 3 Then Col = "B"

                        For j As Integer = Exl.StaRow To Exl.EndRow(Col)

                            Exl.Sheet.Range(Col & j).Value = ProductionSer.MinusStaCode(Exl.Sheet.Range(Col & j).Value)

                        Next

                    End If

                Next

            Catch ex As Exception
                Method.ExceptionWrite(ex)
            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True

            End Try

      '  End If

    End Sub

    ''' <summary>
    ''' 清单计算
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CodeCal_Click(sender As Object, e As RibbonControlEventArgs) Handles CodeCal.Click

        '   If Not IsRegedit AndAlso Not GetIsRegedit() Then

        '      MsgBox("请先注册软件")

        '  Else

        Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.DisplayAlerts = False

            Try

                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook
                Exl.Sheet = Exl.WorkBook.ActiveSheet

                SerialCal()

            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True

            End Try

       ' End If

    End Sub

    ''' <summary>
    ''' 清单分类
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BomSort_Click(sender As Object, e As RibbonControlEventArgs) Handles BomSort.Click

        '  If Not IsRegedit AndAlso Not GetIsRegedit() Then

        '       MsgBox("请先注册软件")

        '   Else

        Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.DisplayAlerts = False

            Try
                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook
                Exl.Sheet = Exl.WorkBook.ActiveSheet

                SerialSort(False)

            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True

            End Try

      '  End If

    End Sub

    ''' <summary>
    ''' 编码合并
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub IDCombin_Click(sender As Object, e As RibbonControlEventArgs) Handles IDCombin.Click

        ' If Not IsRegedit AndAlso Not GetIsRegedit() Then

        '     MsgBox("请先注册软件")

        ' Else

        Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.DisplayAlerts = False

            Try

                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook
                Exl.Sheet = Exl.WorkBook.ActiveSheet

                If Exl.ColNum > 4 Then Globals.ThisAddIn.SerialCombin("b") Else Globals.ThisAddIn.SerialCombin()

            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True

            End Try

     '   End If

    End Sub

    ''' <summary>
    ''' 生产清单
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub PBomSingle_Click(sender As Object, e As RibbonControlEventArgs) Handles PBomSingle.Click

        '  If Not IsRegedit AndAlso Not GetIsRegedit() Then

        '    MsgBox("请先注册软件")

        '  Else

        Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.DisplayAlerts = False

            Try

                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook
                Exl.Sheet = Exl.WorkBook.ActiveSheet

                Globals.ThisAddIn.CreatePBom()

            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True

            End Try

      '  End If

    End Sub

    ''' <summary>
    ''' 打包清单
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub PackBom_Click(sender As Object, e As RibbonControlEventArgs) Handles PackBom.Click

        '    If Not IsRegedit AndAlso Not GetIsRegedit() Then

        '      MsgBox("请先注册软件")

        '   Else

        Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.DisplayAlerts = False

            Try

                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook
                Exl.Sheet = Exl.WorkBook.ActiveSheet

                Globals.ThisAddIn.PackBom()

            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True

            End Try

     '   End If

    End Sub

    ''' <summary>
    ''' 销售清单
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub SellBom_Click(sender As Object, e As RibbonControlEventArgs)

        ' If Not IsRegedit AndAlso Not GetIsRegedit() Then

        '      MsgBox("请先注册软件")

        ' Else

        Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.DisplayAlerts = False

            Try

                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook

                Globals.ThisAddIn.SellBom()

            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True

            End Try

     '   End If

    End Sub

    ''' <summary>
    ''' 二维清单
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AutoAllBom_Click(sender As Object, e As RibbonControlEventArgs) Handles AutoAllBom.Click

        ' If Not IsRegedit AndAlso Not GetIsRegedit() Then

        ' MsgBox("请先注册软件")

        '  Else

        Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.DisplayAlerts = False
            '   Globals.ThisAddIn.Application.ActivePrinter = "Fax 在 Ne00:"
            'Globals.ThisAddIn.Application.PrintCommunication = False
            Try

                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook
                Exl.Sheet = Exl.WorkBook.ActiveSheet

                Globals.ThisAddIn.AllBom

            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True
                '    Globals.ThisAddIn.Application.PrintCommunication = True
                '  Globals.ThisAddIn.Application.ActivePrinter = "Fax 在 Ne00:"
            End Try

      '  End If

    End Sub
    ''' <summary>
    ''' 分表
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click

        ' If Not IsRegedit AndAlso Not GetIsRegedit() Then

        ' MsgBox("请先注册软件")

        ' Else

        Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.DisplayAlerts = False

            Try

                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook
                Exl.Sheet = Exl.WorkBook.ActiveSheet

                Globals.ThisAddIn.Fen()

            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True

            End Try

        '  End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click

        '  If Not IsRegedit AndAlso Not GetIsRegedit() Then

        '      MsgBox("请先注册软件")

        '  Else

        Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.DisplayAlerts = False

            Try

                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook
                Exl.Sheet = Exl.WorkBook.ActiveSheet

                Globals.ThisAddIn.CombinSheet()

            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True

            End Try

        '   End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs)
        ' If Not IsRegedit AndAlso Not GetIsRegedit() Then

        'MsgBox("请先注册软件")

        '    Else

        Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.DisplayAlerts = False

            Try

                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook

                Globals.ThisAddIn.SellBom1()

            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True

            End Try

        ' End If
    End Sub

    Private Sub 图纸编号_Click(sender As Object, e As RibbonControlEventArgs) Handles 图纸编号.Click
        ' If Not IsRegedit AndAlso Not GetIsRegedit() Then

        'MsgBox("请先注册软件")

        '  Else

        Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.DisplayAlerts = False

            Try

                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook

                Globals.ThisAddIn.DaringN()

            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True

            End Try

        ' End If
    End Sub

    Private Sub Button5_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click

        '  If Not IsRegedit AndAlso Not GetIsRegedit() Then

        'MsgBox("请先注册软件")

        ' Else

        Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.DisplayAlerts = False

            Try

                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook

                Globals.ThisAddIn.DaringN1()

            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True

            End Try

        ' End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As RibbonControlEventArgs)

        '   If Not IsRegedit AndAlso Not GetIsRegedit() Then

        ' MsgBox("请先注册软件")

        '  Else

        Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.DisplayAlerts = False

            Try

                Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook
                Exl.Sheet = Exl.WorkBook.ActiveSheet
                Globals.ThisAddIn.AllBom2()

            Finally

                Globals.ThisAddIn.Application.ScreenUpdating = True
                Globals.ThisAddIn.Application.DisplayAlerts = True

            End Try

        ' End If

    End Sub

    Private Sub Button6_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click

        Globals.ThisAddIn.Application.ScreenUpdating = False
        Globals.ThisAddIn.Application.DisplayAlerts = False

        Try

            Globals.ThisAddIn.changeW()

        Finally

            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.DisplayAlerts = True

        End Try

    End Sub

    Private Sub Button7_Click(sender As Object, e As RibbonControlEventArgs) 

        Globals.ThisAddIn.Application.ScreenUpdating = False
        Globals.ThisAddIn.Application.DisplayAlerts = False

        Try

            Exl.WorkBook = Globals.ThisAddIn.Application.ActiveWorkbook

            Globals.ThisAddIn.CombinSheet11()

        Finally

            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.DisplayAlerts = True

        End Try
    End Sub
End Class