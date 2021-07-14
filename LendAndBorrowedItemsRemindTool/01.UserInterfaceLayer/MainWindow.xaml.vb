Imports System.Data.SqlClient
Imports DingTalk.Api
Imports DingTalk.Api.Request
Imports DingTalk.Api.Response
Imports OfficeOpenXml

Class MainWindow

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

        Me.Title = $"{My.Application.Info.Title} V{AppSettingHelper.Instance.ProductVersion}"

        '设置使用方式为个人使用
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim tmpAppCenterSparkle As New AppCenterSparkle(AppSettingHelper.AppKey, Me)
        tmpAppCenterSparkle.CheckUpdateAsync()

        StartAutoRun.IsChecked = AppSettingHelper.Instance.StartAutoRun

    End Sub

    Public Sub Shutdown()

        AppSettingHelper.SaveToLocaltion()

        System.Windows.Application.Current.Shutdown()

        End

    End Sub

    Private Sub UpdateInfoMenuItem_Click(sender As Object, e As RoutedEventArgs)

        FileHelper.Open("https://install.appcenter.ms/users/707wk/apps/jie4-chu1-jie4-ru4-wu4-pin3-ti2-xing3-gong1-ju4/distribution_groups/public")

    End Sub

    Private Sub AboutMenuItem_Click(sender As Object, e As RoutedEventArgs)

        Dim tmpWindow As New AboutWindow With {
          .Owner = Me
        }
        tmpWindow.ShowDialog()

    End Sub

    Private Sub Window_Closing(sender As Object, e As ComponentModel.CancelEventArgs)

        e.Cancel = True

        Me.WindowState = WindowState.Minimized

    End Sub

    Private Sub Test(sender As Object, e As RoutedEventArgs)

        Dim tmpWindow As New Wangk.ResourceWPF.BackgroundWork(Me) With {
            .Title = "初始化"
        }

        tmpWindow.Run(Sub(uie)
                          Dim stepCount = 5

#Region "获取未结束表单"
                          uie.Write("获取未结束表单", 0 * 100 / stepCount)

                          AppSettingHelper.Instance.DocumentItems.Clear()

                          Using SqlConn As New SqlConnection(AppSettingHelper.Instance.ERPSqlServerConnStr)
                              SqlConn.Open()

                              Using tmpSqlCommand = SqlConn.CreateCommand
                                  tmpSqlCommand.CommandText = $"select

INVTF.TF003 as 交易日期,
INVTF.TF004 as 交易对象,
INVTF.TF005 as 对象编号,
INVTF.TF006 as 对象简称,
INVTF.TF015 as 对象全称,
INVTF.TF008 as 员工编号,
CMSMV.MV002 as 员工姓名,
rtrim(CMSMQ.MQ002)+'('+TempINVTG.TG001+')' as 交易单别,
TempINVTG.TG002 as 交易单号,
TempINVTG.MaterialCount as 需归还物品种数,
TempINVTG1.MinDate as 最近需归还日期

from

-- 合并单据信息
(select

TG001,
TG002,
COUNT(1) as MaterialCount
--MIN(TG027) as MinDate

from INVTG

where TG022='Y'
and TG024='N'
and TG009>0

group by TG001,TG002

) as TempINVTG

-- 计算有效日期
left join(select

TG001,
TG002,
MIN(TG027) as MinDate

from INVTG

where TG022='Y'
and TG024='N'
and TG009>0
and TG027<>''

group by TG001,TG002

) as TempINVTG1
on TempINVTG1.TG001=TempINVTG.TG001
and TempINVTG1.TG002=TempINVTG.TG002

-- 关联员工及对象信息
left join INVTF
on INVTF.TF001=TempINVTG.TG001
and INVTF.TF002=TempINVTG.TG002

-- 关联员工基本信息
left join CMSMV
on CMSMV.MV001=INVTF.TF008

left join CMSMQ
on CMSMQ.MQ001=TempINVTG.TG001

order by 交易日期
"

                                  Using tmpSqlDataReader = tmpSqlCommand.ExecuteReader

                                      While tmpSqlDataReader.Read

                                          Dim tmpDocumentInfo = New DocumentInfo With {
                                              .JYRQ = Date.ParseExact($"{tmpSqlDataReader(0)}", "yyyyMMdd", Nothing),
                                              .JYDX = Val($"{tmpSqlDataReader(1)}"),
                                              .DXBH = $"{tmpSqlDataReader(2)}".Trim,
                                              .DXJC = $"{tmpSqlDataReader(3)}".Trim,
                                              .DXQC = $"{tmpSqlDataReader(4)}".Trim,
                                              .YGBH = $"{tmpSqlDataReader(5)}".Trim,
                                              .YGXM = $"{tmpSqlDataReader(6)}".Trim,
                                              .JYDB = $"{tmpSqlDataReader(7)}".Trim,
                                              .JYDH = $"{tmpSqlDataReader(8)}".Trim,
                                              .XGHWPZS = tmpSqlDataReader(9),
                                              .ZJXGHRQ = If(String.IsNullOrWhiteSpace($"{tmpSqlDataReader(10)}"),
                                              .JYRQ.AddDays(AppSettingHelper.Instance.DefaultUsageDays),
                                              Date.ParseExact($"{tmpSqlDataReader(10)}", "yyyyMMdd", Nothing))
                                          }

                                          AppSettingHelper.Instance.DocumentItems.Add(tmpDocumentInfo)

                                      End While

                                  End Using

                              End Using

                              SqlConn.Close()
                          End Using

                          Console.WriteLine($"表单数 : {AppSettingHelper.Instance.DocumentItems.Count}")
#End Region

#Region "获取钉钉AccessToken"
                          uie.Write("获取钉钉AccessToken", 1 * 100 / stepCount)

                          GetDingTalkAccessToken()
#End Region

#Region "获取钉钉部门信息"
                          uie.Write("获取钉钉部门信息", 2 * 100 / stepCount)

                          AppSettingHelper.Instance.DingTalkDepartmentIDItems.Clear()

                          GetDingTalkDepartmentIDItems(1)

                          Console.WriteLine($"部门数 : {AppSettingHelper.Instance.DingTalkDepartmentIDItems.Count}")
#End Region

#Region "获取钉钉员工信息"
                          uie.Write("获取钉钉员工信息", 3 * 100 / stepCount)

                          AppSettingHelper.Instance.DingTalkUserJobNumberItems.Clear()

                          Dim tmpID = 1
                          For Each item In AppSettingHelper.Instance.DingTalkDepartmentIDItems

                              uie.Write($"获取钉钉员工信息 {tmpID}/{AppSettingHelper.Instance.DingTalkDepartmentIDItems.Count}")
                              tmpID += 1

                              GetDingTalkUserItems(item)
                          Next

                          Console.WriteLine($"有工号的员工数 : {AppSettingHelper.Instance.DingTalkUserJobNumberItems.Count}")
#End Region

#Region "发送工作通知消息"
                          uie.Write("获取钉钉员工信息", 4 * 100 / stepCount)

                          ' 无对应的钉钉账号的ERP用户
                          Dim NotHaveJobNumberUserItems As New Dictionary(Of String, String)

                          For Each item In AppSettingHelper.Instance.DocumentItems

                              ' 判断是否有对应的钉钉账号
                              If Not AppSettingHelper.Instance.DingTalkUserJobNumberItems.ContainsKey(item.YGBH) Then

                                  If Not NotHaveJobNumberUserItems.ContainsKey(item.YGBH) Then
                                      NotHaveJobNumberUserItems.Add(item.YGBH, item.YGXM)

                                  End If

                                  Continue For
                              End If

                              ' 发送消息
                              SendDingTalkMessage(AppSettingHelper.Instance.DingTalkUserJobNumberItems(item.YGBH), item)

                          Next

                          For Each item In NotHaveJobNumberUserItems
                              Console.WriteLine($"{item.Key} {item.Value}")
                          Next

#End Region

                      End Sub)

        If tmpWindow.Error IsNot Nothing Then
            MsgBox(tmpWindow.Error.Message, MsgBoxStyle.Information, tmpWindow.Title)
            Exit Sub
        End If

        If tmpWindow.IsCancel Then
            Wangk.ResourceWPF.Toast.ShowInfo(Me, $"操作已取消")
        Else
            Wangk.ResourceWPF.Toast.ShowSuccess(Me, $"操作完毕")
        End If

    End Sub

#Region "获取钉钉调用企业接口凭证"
    ''' <summary>
    ''' 获取钉钉调用企业接口凭证
    ''' </summary>
    Private Sub GetDingTalkAccessToken()

        Dim client As IDingTalkClient = New DefaultDingTalkClient("https://oapi.dingtalk.com/gettoken")
        Dim req As OapiGettokenRequest = New OapiGettokenRequest()
        req.Appkey = AppSettingHelper.Instance.DingTalkAppKey
        req.Appsecret = AppSettingHelper.Instance.DingTalkAppSecret
        req.SetHttpMethod("GET")
        Dim rsp As OapiGettokenResponse = client.Execute(req, Nothing)
        AppSettingHelper.Instance.DingTalkAccessToken = rsp.AccessToken

    End Sub
#End Region

#Region "获取钉钉部门信息"
    ''' <summary>
    ''' 获取钉钉部门信息
    ''' </summary>
    Private Sub GetDingTalkDepartmentIDItems(parentDepartmentID As Long)

        Dim client As IDingTalkClient = New DefaultDingTalkClient("https://oapi.dingtalk.com/topapi/v2/department/listsub")
        Dim req As OapiV2DepartmentListsubRequest = New OapiV2DepartmentListsubRequest()
        req.DeptId = parentDepartmentID
        Dim rsp As OapiV2DepartmentListsubResponse = client.Execute(req, AppSettingHelper.Instance.DingTalkAccessToken)

        If rsp.Result Is Nothing Then
            Exit Sub
        End If

        For Each item In rsp.Result
            AppSettingHelper.Instance.DingTalkDepartmentIDItems.Add(item.DeptId)

            GetDingTalkDepartmentIDItems(item.DeptId)
        Next

    End Sub
#End Region

#Region "获取钉钉部门用户信息"
    ''' <summary>
    ''' 获取钉钉部门用户信息
    ''' </summary>
    Private Sub GetDingTalkUserItems(parentDepartmentID As Long)

        Dim client As IDingTalkClient = New DefaultDingTalkClient("https://oapi.dingtalk.com/topapi/v2/user/list")

        Dim Cursor As Long = 0

        Do

            Dim req As OapiV2UserListRequest = New OapiV2UserListRequest()
            req.DeptId = parentDepartmentID
            req.Cursor = Cursor
            req.Size = 100L
            Dim rsp As OapiV2UserListResponse = client.Execute(req, AppSettingHelper.Instance.DingTalkAccessToken)

            If rsp.Result.List Is Nothing Then
                Exit Sub
            End If

            For Each item In rsp.Result.List

                If String.IsNullOrWhiteSpace(item.JobNumber) Then
                    Continue For
                End If

                If Not AppSettingHelper.Instance.DingTalkUserJobNumberItems.ContainsKey(item.JobNumber) Then
                    AppSettingHelper.Instance.DingTalkUserJobNumberItems.Add(item.JobNumber, item.Userid)
                End If

            Next

            Cursor += 1
        Loop

    End Sub
#End Region

#Region "向钉钉用户发送工作通知消息"
    ''' <summary>
    ''' 向钉钉用户发送工作通知消息
    ''' </summary>
    Private Sub SendDingTalkMessage(dingTalkUserid As String,
                                    doc As DocumentInfo)

        Dim client As IDingTalkClient = New DefaultDingTalkClient("https://oapi.dingtalk.com/topapi/message/corpconversation/asyncsend_v2")
        Dim req As OapiMessageCorpconversationAsyncsendV2Request = New OapiMessageCorpconversationAsyncsendV2Request()
        req.AgentId = AppSettingHelper.Instance.DingTalkAgentId
        req.UseridList = dingTalkUserid
        Dim obj1 As OapiMessageCorpconversationAsyncsendV2Request.MsgDomain = New OapiMessageCorpconversationAsyncsendV2Request.MsgDomain()
        obj1.Msgtype = "markdown"
        Dim obj2 As OapiMessageCorpconversationAsyncsendV2Request.MarkdownDomain = New OapiMessageCorpconversationAsyncsendV2Request.MarkdownDomain()
        obj2.Text = "002"
        obj2.Title = "001"
        obj1.Markdown = obj2
        req.Msg_ = obj1
        Dim rsp As OapiMessageCorpconversationAsyncsendV2Response = client.Execute(req, AppSettingHelper.Instance.DingTalkAccessToken)
        Console.WriteLine(rsp.Body)

    End Sub
#End Region

    Private Sub SaveChange(sender As Object, e As RoutedEventArgs)

        Try

            AppSettingHelper.Instance.StartAutoRun = StartAutoRun.IsChecked

            AppSettingHelper.Instance.SendMsgTime = TimeSpan.Parse(SendMsgTime.Value)
            AppSettingHelper.Instance.AdvanceNoticeDays = Val(AdvanceNoticeDays.Value)
            AppSettingHelper.Instance.DefaultUsageDays = Val(DefaultUsageDays.Value)

            AppSettingHelper.Instance.ERPSqlServerConnStr = ERPSqlServerConnStr.Value

            AppSettingHelper.Instance.DingTalkAgentId = CLng(DingTalkAgentIdStr.Value)
            AppSettingHelper.Instance.DingTalkAppKey = DingTalkAppKey.Value
            AppSettingHelper.Instance.DingTalkAppSecret = DingTalkAppSecret.Value

        Catch ex As Exception
            Wangk.ResourceWPF.Toast.ShowError(Me, ex.Message)
            Exit Sub
        End Try

        SendMsgTime.AddHistoryValue()
        AdvanceNoticeDays.AddHistoryValue()
        DefaultUsageDays.AddHistoryValue()
        ERPSqlServerConnStr.AddHistoryValue()
        DingTalkAgentIdStr.AddHistoryValue()
        DingTalkAppKey.AddHistoryValue()
        DingTalkAppSecret.AddHistoryValue()

        AppSettingHelper.SaveToLocaltion()

    End Sub

    Private Sub NotSaveChange(sender As Object, e As RoutedEventArgs)

    End Sub

End Class
