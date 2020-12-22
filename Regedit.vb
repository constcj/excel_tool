
Option Strict Off

Imports System.Management

Public Class Regedit
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TextBox1.Text = Str1

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

        Str2 = TextBox2.Text

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Me.Close()

        If EncryptDes(Str1, "ZjMjjWEi", "95135746") = Str2 Then

            CreateKey(Str2)

            MsgBox("注册成功")

        Else

            MsgBox("注册失败")

        End If

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub
End Class

Module RegeditCal

    Public IsRegedit As Boolean = False

    ''' <summary>
    ''' 加密
    ''' </summary>
    ''' <param name="SourceStr">需要加密的字符串</param>
    ''' <param name="myKey">加密的8个任意字符</param>
    ''' <param name="myIV">密码本8位数字</param>
    ''' <returns></returns>
    Public Function EncryptDes(ByVal SourceStr As String, ByVal myKey As String, ByVal myIV As String) As String '使用的DES对称加密  
        '  Dim SourceStr1 As String = "BFEBFBFF000906E9WD-WMC6Y"
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
        EncryptDes = Convert.ToBase64String(ms.GetBuffer(), 0, ms.Length)

    End Function
    ''' <summary>
    ''' 解密
    ''' </summary>
    ''' <param name="SourceStr">加密完成的字符串</param>
    ''' <param name="myKey">加密的8个任意字符</param>
    ''' <param name="myIV">加密的密码本8位数字</param>
    ''' <returns></returns>
    Public Function DecryptDes(ByVal SourceStr As String, ByVal myKey As String, ByVal myIV As String) As String    '使用标准DES对称解密  
        Dim des As New System.Security.Cryptography.DESCryptoServiceProvider With
            {.Key = Encoding.UTF8.GetBytes(myKey), .IV = Encoding.UTF8.GetBytes(myIV)} 'DES算法  
        'Dim DES As New System.Security.Cryptography.TripleDESCryptoServiceProvider'TripleDES算法  

        Dim buffer As Byte() = Convert.FromBase64String(SourceStr)
        Dim ms As New System.IO.MemoryStream(buffer)
        Dim cs As New System.Security.Cryptography.CryptoStream(ms, des.CreateDecryptor(), System.Security.Cryptography.CryptoStreamMode.Read)
        Dim sr As New System.IO.StreamReader(cs)
        DecryptDes = sr.ReadToEnd()
    End Function

    ''' <summary>
    ''' 新建键值
    ''' </summary>
    ''' <param name="Key"></param>
    Public Sub CreateKey(ByVal Key As String)

        '返回当前用户键
        Dim Key1 As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser

        '返回当前用户键下的northsnow键,如果想创建项，必须指定第二个参数为true
        Dim Key2 As Microsoft.Win32.RegistryKey = Key1.OpenSubKey("3DConfiguraKey", True)

        '如果键不存在就创建它
        If Key2 Is Nothing Then Key2 = Key1.CreateSubKey("3DConfiguraKey")

        '创建项，如果不存在就创建，如果存在则覆盖
        Key2.SetValue("ExcelKey", Key)

    End Sub

    ''' <summary>
    ''' 得到键值
    ''' </summary>
    ''' <returns></returns>
    Public Function GetKey() As String

        Dim Key1 As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser

        Dim Key2 As Microsoft.Win32.RegistryKey = Key1.OpenSubKey("3DConfiguraKey", True)

        Dim Key As String = ""

        Try

            Key = Key2.GetValue("ExcelKey").ToString

        Catch

        End Try

        Return Key

    End Function

    ''' <summary>
    ''' 网卡 MacAddress 硬盘序列号 C 取第一块硬盘编号
    ''' </summary>
    ''' <returns></returns>
    Public Function GetDiskVolumeSerialNumber() As String

        Dim mc As New ManagementClass("Win32_NetworkAdapterConfiguration")
        Dim moc As ManagementObjectCollection = mc.GetInstances()

        For Each mo As ManagementObject In moc

            If CBool(mo("IPEnabled")) = True Then Exit For

        Next

        Dim disk As New ManagementObject("win32_logicaldisk.deviceid=""c:""")

        disk.[Get]()

        Dim strHardDiskID As String = disk.GetPropertyValue("VolumeSerialNumber").ToString()
        Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMedia")

        For Each mo As ManagementObject In searcher.[Get]()

            strHardDiskID = mo("SerialNumber").ToString().Trim()

            Exit For

        Next

        Return strHardDiskID

    End Function

    ''' <summary>
    ''' 获得CPU的序列号
    ''' </summary>
    ''' <returns></returns>
    Public Function GetCpu() As String
        Dim strCpu As String = Nothing
        Dim myCpu As New ManagementClass("win32_Processor")
        Dim myCpuConnection As ManagementObjectCollection = myCpu.GetInstances()
        For Each myObject As ManagementObject In myCpuConnection
            strCpu = myObject.Properties("Processorid").Value.ToString()
            Exit For
        Next
        Return strCpu
    End Function

    ''' <summary>
    ''' 生成机器码
    ''' </summary>
    ''' <returns></returns>
    Public Function GetMNum() As String

        Dim strNum As String = GetCpu() & GetDiskVolumeSerialNumber()

        '获得24位Cpu和硬盘序列号

        Dim strMNum As String = strNum.Substring(0, 24)

        '从生成的字符串中取出前24个字符做为机器码
        Return strMNum

    End Function

    ''' <summary>
    ''' 得到软件是否注册
    ''' </summary>
    ''' <returns></returns>
    Public Function GetIsRegedit() As Boolean

        Dim Str1 As String = EncryptDes(GetMNum(), "ZjMjjWEi", "95135746")

        Dim str2 As String = GetKey()

        If Str1 = str2 Then

            IsRegedit = True

        Else

            IsRegedit = False

        End If

        Return IsRegedit

    End Function

End Module