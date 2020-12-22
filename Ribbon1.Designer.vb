Partial Class Ribbon2

    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 类撰写设计器支持所必需的
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '组件设计器需要此调用。
        InitializeComponent()

    End Sub

    '组件重写释放以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '组件设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是组件设计器所必需的
    '可使用组件设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.AutoAllBom = Me.Factory.CreateRibbonButton
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.PBomSingle = Me.Factory.CreateRibbonButton
        Me.PackBom = Me.Factory.CreateRibbonButton
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Start = Me.Factory.CreateRibbonGroup
        Me.CodeCal = Me.Factory.CreateRibbonButton
        Me.BomSort = Me.Factory.CreateRibbonButton
        Me.IDCombin = Me.Factory.CreateRibbonButton
        Me.InsertTable = Me.Factory.CreateRibbonButton
        Me.CodeCheck = Me.Factory.CreateRibbonButton
        Me.图纸编号 = Me.Factory.CreateRibbonButton
        Me.PrintFormat = Me.Factory.CreateRibbonButton
        Me.CancelItemCode = Me.Factory.CreateRibbonButton
        Me.AddItemCode = Me.Factory.CreateRibbonButton
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Regist = Me.Factory.CreateRibbonButton
        Me.Descripe = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Start.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Start)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Label = "忠旺清单"
        Me.Tab1.Name = "Tab1"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.AutoAllBom)
        Me.Group3.Items.Add(Me.Button4)
        Me.Group3.Items.Add(Me.Button6)
        Me.Group3.Items.Add(Me.PBomSingle)
        Me.Group3.Items.Add(Me.PackBom)
        Me.Group3.Items.Add(Me.Button2)
        Me.Group3.Label = "自动表单"
        Me.Group3.Name = "Group3"
        '
        'AutoAllBom
        '
        Me.AutoAllBom.Label = "二维清单"
        Me.AutoAllBom.Name = "AutoAllBom"
        '
        'Button4
        '
        Me.Button4.Label = "合表"
        Me.Button4.Name = "Button4"
        '
        'Button6
        '
        Me.Button6.Label = "变更清单"
        Me.Button6.Name = "Button6"
        '
        'PBomSingle
        '
        Me.PBomSingle.Label = "生产清单"
        Me.PBomSingle.Name = "PBomSingle"
        '
        'PackBom
        '
        Me.PackBom.Label = "打包清单"
        Me.PackBom.Name = "PackBom"
        '
        'Button2
        '
        Me.Button2.Label = " "
        Me.Button2.Name = "Button2"
        '
        'Start
        '
        Me.Start.Items.Add(Me.CodeCal)
        Me.Start.Items.Add(Me.BomSort)
        Me.Start.Items.Add(Me.IDCombin)
        Me.Start.Items.Add(Me.InsertTable)
        Me.Start.Items.Add(Me.CodeCheck)
        Me.Start.Items.Add(Me.图纸编号)
        Me.Start.Items.Add(Me.PrintFormat)
        Me.Start.Items.Add(Me.CancelItemCode)
        Me.Start.Items.Add(Me.AddItemCode)
        Me.Start.Items.Add(Me.Button1)
        Me.Start.Label = "功能"
        Me.Start.Name = "Start"
        '
        'CodeCal
        '
        Me.CodeCal.Label = "清单计算"
        Me.CodeCal.Name = "CodeCal"
        '
        'BomSort
        '
        Me.BomSort.Label = "清单分类"
        Me.BomSort.Name = "BomSort"
        '
        'IDCombin
        '
        Me.IDCombin.Label = "编码合并"
        Me.IDCombin.Name = "IDCombin"
        '
        'InsertTable
        '
        Me.InsertTable.Label = "插入表头"
        Me.InsertTable.Name = "InsertTable"
        '
        'CodeCheck
        '
        Me.CodeCheck.Label = "编码检查"
        Me.CodeCheck.Name = "CodeCheck"
        '
        '图纸编号
        '
        Me.图纸编号.Label = "图纸编号"
        Me.图纸编号.Name = "图纸编号"
        '
        'PrintFormat
        '
        Me.PrintFormat.Label = "打印格式"
        Me.PrintFormat.Name = "PrintFormat"
        '
        'CancelItemCode
        '
        Me.CancelItemCode.Label = "去除前缀"
        Me.CancelItemCode.Name = "CancelItemCode"
        '
        'AddItemCode
        '
        Me.AddItemCode.Label = "添加前缀"
        Me.AddItemCode.Name = "AddItemCode"
        '
        'Button1
        '
        Me.Button1.Label = " "
        Me.Button1.Name = "Button1"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Regist)
        Me.Group2.Items.Add(Me.Descripe)
        Me.Group2.Items.Add(Me.Button3)
        Me.Group2.Items.Add(Me.Button5)
        Me.Group2.Label = "其他"
        Me.Group2.Name = "Group2"
        '
        'Regist
        '
        Me.Regist.Label = "注册"
        Me.Regist.Name = "Regist"
        '
        'Descripe
        '
        Me.Descripe.Label = "说明"
        Me.Descripe.Name = "Descripe"
        '
        'Button3
        '
        Me.Button3.Label = "分表"
        Me.Button3.Name = "Button3"
        '
        'Button5
        '
        Me.Button5.Label = "未出图"
        Me.Button5.Name = "Button5"
        '
        'Ribbon2
        '
        Me.Name = "Ribbon2"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Start.ResumeLayout(False)
        Me.Start.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents CodeCal As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AddItemCode As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents CancelItemCode As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents PrintFormat As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BomSort As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Regist As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Descripe As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Start As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents CodeCheck As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents InsertTable As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents PBomSingle As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents PackBom As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents IDCombin As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AutoAllBom As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents 图纸编号 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()>
    Friend ReadOnly Property Ribbon1() As Ribbon2
        Get
            Return Me.GetRibbon(Of Ribbon2)()
        End Get
    End Property

End Class
