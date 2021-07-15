Public Class DocumentInfo

    ''' <summary>
    ''' 交易日期
    ''' </summary>
    Public JYRQ As Date

    ''' <summary>
    ''' 交易对象
    ''' </summary>
    Public JYDX As Integer

    Private Shared JYDXIDToNameItems As New Dictionary(Of Integer, String) From {
        {1, "客户"},
        {2, "供应商"},
        {3, "人员"},
        {9, "其它"}
        }

    ''' <summary>
    ''' 交易对象名称
    ''' </summary>
    Public ReadOnly Property JYDXStr As String
        Get

            If Not JYDXIDToNameItems.ContainsKey(JYDX) Then
                Return "未定义"
            End If

            Return JYDXIDToNameItems(JYDX)
        End Get
    End Property

    ''' <summary>
    ''' 对象编号
    ''' </summary>
    Public DXBH As String

    ''' <summary>
    ''' 对象简称
    ''' </summary>
    Public DXJC As String

    ''' <summary>
    ''' 对象全称
    ''' </summary>
    Public DXQC As String

    ''' <summary>
    ''' 员工编号
    ''' </summary>
    Public YGBH As String

    ''' <summary>
    ''' 员工姓名
    ''' </summary>
    Public YGXM As String

    ''' <summary>
    ''' 交易单别
    ''' </summary>
    Public JYDB As String

    ''' <summary>
    ''' 交易单号
    ''' </summary>
    Public JYDH As String

    ''' <summary>
    ''' 需归还物品种数
    ''' </summary>
    Public XGHWPZS As Integer

    ''' <summary>
    ''' 最近需归还日期
    ''' </summary>
    Public ZJXGHRQ As Date

End Class
