Imports System.Data.SqlClient
Imports System.IO
Imports System.Timers
Imports DingTalk.Api
Imports DingTalk.Api.Request
Imports DingTalk.Api.Response
Imports Microsoft.AppCenter.Analytics
Imports OfficeOpenXml

Class MainWindow

    Private SendTimer As Timer

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

        Me.Title = $"{My.Application.Info.Title} V{AppSettingHelper.Instance.ProductVersion}"

        '设置使用方式为个人使用
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim tmpAppCenterSparkle As New AppCenterSparkle(AppSettingHelper.AppKey, Me)
        tmpAppCenterSparkle.CheckUpdateAsync()

        StartAutoRun.IsChecked = AppSettingHelper.Instance.StartAutoRun

        SendTimer = New Timer With {
            .Interval = 40 * 1000
        }
        AddHandler SendTimer.Elapsed, AddressOf SendTimerElapsed

        SendTimer.Start()

        ' 开机自启后最小化
        If AppSettingHelper.Instance.StartAutoRun Then
            Me.WindowState = WindowState.Minimized
        End If

    End Sub

    ''' <summary>
    ''' 定时处理
    ''' </summary>
    Private Sub SendTimerElapsed(sender As Object, e As ElapsedEventArgs)
        Console.WriteLine("定时处理")

        ' 心跳包
        If Now.Minute Mod 30 = 0 AndAlso
            (Now - AppSettingHelper.Instance.LastSendHeartBeatDate).TotalMinutes > 1 Then

            AppSettingHelper.Instance.LastSendHeartBeatDate = Now
            AppSettingHelper.SaveToLocaltion()

            Analytics.TrackEvent("心跳包")
            AppSettingHelper.Instance.Logger.Info("心跳包")
        End If

        ' 周末不发送
        If Now.DayOfWeek = DayOfWeek.Saturday OrElse
            Now.DayOfWeek = DayOfWeek.Sunday Then
            Exit Sub
        End If

        ' 定时发送
        If Now.Hour <> AppSettingHelper.Instance.SendMsgTime.Hours OrElse
            Now.Minute <> AppSettingHelper.Instance.SendMsgTime.Minutes Then
            Exit Sub
        End If

        ' 一天只自动发送一次
        If AppSettingHelper.Instance.LastSendDate = Now.Date Then
            Exit Sub
        End If

        AppSettingHelper.Instance.LastSendDate = Now.Date
        AppSettingHelper.SaveToLocaltion()

        Analytics.TrackEvent("自动发送通知")
        AppSettingHelper.Instance.Logger.Info("自动发送通知")

        Me.Dispatcher.Invoke(Sub()
                                 SendMessage(Nothing, Nothing)
                             End Sub)

    End Sub

    Public Sub Shutdown()

        SendTimer.Stop()
        RemoveHandler SendTimer.Elapsed, AddressOf SendTimerElapsed

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

    Private Sub SendMessage(sender As Object, e As RoutedEventArgs)

        If e IsNot Nothing Then
            Analytics.TrackEvent("手动发送通知")
            AppSettingHelper.Instance.Logger.Info("手动发送通知")
        End If

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
                          uie.Write("发送工作通知消息", 4 * 100 / stepCount)

                          ' 无对应的钉钉账号的ERP用户
                          Dim NotHaveJobNumberUserItems As New Dictionary(Of String, String)

                          tmpID = 1
                          For Each item In AppSettingHelper.Instance.DocumentItems

                              uie.Write($"发送工作通知消息 {tmpID}/{AppSettingHelper.Instance.DocumentItems.Count}")
                              tmpID += 1

                              ' 判断是否需要提醒
                              If (item.ZJXGHRQ - Now).TotalDays > AppSettingHelper.Instance.AdvanceNoticeDays Then
                                  Continue For
                              End If

                              ' 判断是否有对应的钉钉账号
                              If Not AppSettingHelper.Instance.DingTalkUserJobNumberItems.ContainsKey(item.YGBH) Then

                                  If Not NotHaveJobNumberUserItems.ContainsKey(item.YGBH) Then
                                      NotHaveJobNumberUserItems.Add(item.YGBH, item.YGXM)

                                  End If

                                  Continue For
                              End If

                              ' 发送消息
                              SendDingTalkWorkMessage(AppSettingHelper.Instance.DingTalkUserJobNumberItems(item.YGBH), item)

                              'SendDingTalkMessage("3349644230885065", item)
                              'Exit For

                          Next

                          'Using tmpStreamWriter As New StreamWriter("无对应的钉钉账号的ERP用户.txt", False)

                          '    tmpStreamWriter.WriteLine($"工号{vbTab}姓名")

                          '    For Each item In NotHaveJobNumberUserItems
                          '        tmpStreamWriter.WriteLine($"{item.Key}{vbTab}{item.Value}")
                          '    Next

                          'End Using

                          ' 通知管理员更新账号信息
                          If NotHaveJobNumberUserItems.Count > 0 Then

                              Dim tempAdminMessage = $"无对应的钉钉账号的ERP用户  
> 工号{vbTab}姓名  
{String.Join(vbCrLf, From item In NotHaveJobNumberUserItems
                     Select $"> {item.Key}{vbTab}{item.Value}  ")}"

                              SendDingTalkMessageToAdmin(tempAdminMessage)

                          End If

#End Region

                      End Sub)

        If tmpWindow.Error IsNot Nothing Then
            Wangk.ResourceWPF.Toast.ShowError(Me, tmpWindow.Error.Message)
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
    Private Sub SendDingTalkWorkMessage(dingTalkUserid As String,
                                        doc As DocumentInfo)

        Dim client As IDingTalkClient = New DefaultDingTalkClient("https://oapi.dingtalk.com/topapi/message/corpconversation/asyncsend_v2")
        Dim req As New OapiMessageCorpconversationAsyncsendV2Request With {
            .AgentId = AppSettingHelper.Instance.DingTalkAgentId,
            .UseridList = dingTalkUserid
        }
        Dim obj1 As New OapiMessageCorpconversationAsyncsendV2Request.MsgDomain With {
            .Msgtype = "markdown"
        }
        Dim obj2 As New OapiMessageCorpconversationAsyncsendV2Request.MarkdownDomain With {
            .Text = $"**<font color=#1296DB>{doc.JYDB} - {doc.JYDH}</font>**

------
通知时间 : {Now:G}  
交易日期 : {doc.JYRQ:d}  
交易对象 : {doc.JYDXStr}  
借出/入对象 : {doc.DXQC}({doc.DXBH})  
需归还物品种数 : {doc.XGHWPZS}  
最近需归还日期 : {doc.ZJXGHRQ:d}",
            .Title = $"{doc.JYDB} - {doc.JYDH}"
        }
        obj1.Markdown = obj2
        req.Msg_ = obj1
        Dim rsp As OapiMessageCorpconversationAsyncsendV2Response = client.Execute(req, AppSettingHelper.Instance.DingTalkAccessToken)

    End Sub
#End Region

#Region "发送消息给所有主管理员"
    ''' <summary>
    ''' 发送消息给所有主管理员
    ''' </summary>
    Private Sub SendDingTalkMessageToAdmin(msg As String)

        For Each item In GetDingTalkAdminItems()
            SendDingTalkAdminMessage(item, msg)
        Next

    End Sub
#End Region

#Region "获取主管理员列表"
    ''' <summary>
    ''' 获取主管理员列表
    ''' </summary>
    Private Function GetDingTalkAdminItems() As List(Of String)

        Dim client As New DefaultDingTalkClient("https://oapi.dingtalk.com/topapi/user/listadmin")
        Dim req As New OapiUserListadminRequest()
        Dim rsp As OapiUserListadminResponse = client.Execute(req, AppSettingHelper.Instance.DingTalkAccessToken)

        Return (From item In rsp.Result
                Where item.SysLevel = 1
                Select item.Userid).ToList

    End Function
#End Region

#Region "发送消息给主管理员"
    ''' <summary>
    ''' 发送消息给主管理员
    ''' </summary>
    Private Sub SendDingTalkAdminMessage(dingTalkUserid As String,
                                         msg As String)

        Dim client As IDingTalkClient = New DefaultDingTalkClient("https://oapi.dingtalk.com/topapi/message/corpconversation/asyncsend_v2")
        Dim req As New OapiMessageCorpconversationAsyncsendV2Request With {
            .AgentId = AppSettingHelper.Instance.DingTalkAgentId,
            .UseridList = dingTalkUserid
        }
        Dim obj1 As New OapiMessageCorpconversationAsyncsendV2Request.MsgDomain With {
            .Msgtype = "markdown"
        }
        Dim obj2 As New OapiMessageCorpconversationAsyncsendV2Request.MarkdownDomain With {
            .Text = msg,
            .Title = "管理员消息"
        }
        obj1.Markdown = obj2
        req.Msg_ = obj1
        Dim rsp As OapiMessageCorpconversationAsyncsendV2Response = client.Execute(req, AppSettingHelper.Instance.DingTalkAccessToken)


    End Sub
#End Region

    Private Sub SaveChange(sender As Object, e As RoutedEventArgs)

        Try

            If AppSettingHelper.Instance.StartAutoRun <> StartAutoRun.IsChecked Then

                If StartAutoRun.IsChecked Then

                    Dim shortcutPath As String = $"{System.Environment.GetFolderPath(Environment.SpecialFolder.Startup) }\{My.Application.Info.ProductName}.lnk"
                    Dim tmpWshShell = New IWshRuntimeLibrary.WshShell()
                    Dim tmpIWshShortcut As IWshRuntimeLibrary.IWshShortcut = tmpWshShell.CreateShortcut(shortcutPath)
                    With tmpIWshShortcut
                        .TargetPath = System.Reflection.Assembly.GetExecutingAssembly().Location
                        .WorkingDirectory = IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
                        .WindowStyle = 1
                        .Description = My.Application.Info.ProductName
                        .IconLocation = .TargetPath
                        .Save()
                    End With

                Else
                    Dim shortcutPath As String = $"{System.Environment.GetFolderPath(Environment.SpecialFolder.Startup) }\{My.Application.Info.ProductName}.lnk"
                    Try
                        IO.File.Delete(shortcutPath)
#Disable Warning CA1031 ' Do not catch general exception types
                    Catch ex As Exception
#Enable Warning CA1031 ' Do not catch general exception types
                    End Try

                End If
            End If

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

        Wangk.ResourceWPF.Toast.ShowSuccess(Me, "修改成功")

    End Sub

    Private Sub NotSaveChange(sender As Object, e As RoutedEventArgs)

    End Sub

End Class
