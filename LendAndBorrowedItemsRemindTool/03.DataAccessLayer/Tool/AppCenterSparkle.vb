Imports System.Net.Http
Imports System.Xml

Public Class AppCenterSparkle

    Private appCastUrlStr As String

    Private currentVersion As String

    Private MainUI As Window

    ''' <summary>
    ''' 查询间隔,默认60s
    ''' </summary>
    Public LoopIntervalSec As Integer = 60

    Public Sub New(appKey As String, UIControl As Window)

        If String.IsNullOrWhiteSpace(appKey) Then Throw New Exception("appKey 不能为空")
        appCastUrlStr = $"https://api.appcenter.ms/v0.1/public/sparkle/apps/{appKey}"

        If String.IsNullOrWhiteSpace(appKey) Then Throw New Exception("UIControl 不能为空")
        MainUI = UIControl

        Dim assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location
        currentVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(assemblyLocation).ProductVersion

    End Sub

    ''' <summary>
    ''' 检查更新
    ''' </summary>
    Public Sub CheckUpdateAsync(Optional enableLoop As Boolean = True)

        Task.Run(Sub()

                     Do
                         Console.WriteLine($"{Now:G} : 检查更新")

                         Try

                             Dim tmpHttpClient As New HttpClient

                             Dim resultStr = tmpHttpClient.GetStringAsync(appCastUrlStr).GetAwaiter.GetResult

                             Dim tmpXmlDocument As New XmlDocument
                             tmpXmlDocument.LoadXml(resultStr)

                             Dim ReleasesItems = tmpXmlDocument.SelectNodes("//item")

                             If ReleasesItems.Count = 0 Then
                                 Console.WriteLine("无版本信息")
                                 Exit Try
                             End If

                             For Each ReleasesItem As XmlNode In ReleasesItems

                                 Dim enclosureNode = ReleasesItem.SelectSingleNode("enclosure")
                                 Dim pubDateNode = ReleasesItem.SelectSingleNode("pubDate")
                                 Dim descriptionNode = ReleasesItem.SelectSingleNode("description")

                                 ' 发布版本
                                 Dim ReleasesVersion = enclosureNode.Attributes("sparkle:version").Value

                                 ' 发布时间
                                 Dim tmpubDate As Date = DateTime.Parse(pubDateNode.InnerText)

                                 ' 发布说明
                                 Dim tmpHtmlDocument As New HtmlAgilityPack.HtmlDocument
                                 tmpHtmlDocument.LoadHtml(descriptionNode.InnerText)

                                 ' 比较版本号大小
                                 If StringHelper.StrCmpLogical(currentVersion, ReleasesVersion) >= 0 Then
                                     Console.WriteLine("不需升级")

                                 Else

                                     If CType(MainUI.Dispatcher.Invoke(Function()

                                                                           Return MsgBox($"有新版本发布
当前版本 : {currentVersion}
最新版本 : {enclosureNode.Attributes("sparkle:version").Value}
发布日期 : {tmpubDate:d}
更新说明 :
{tmpHtmlDocument.DocumentNode.InnerText()}
是否更新 ?", MsgBoxStyle.YesNo Or MsgBoxStyle.Information, "升级提醒")

                                                                       End Function), MsgBoxResult) = MsgBoxResult.Yes Then

                                         Process.Start("DownloadUpdate.exe", $"""{enclosureNode.Attributes("url").Value}"" ""{System.Reflection.Assembly.GetExecutingAssembly().Location}""")

                                         MainUI.Dispatcher.Invoke(Sub()
                                                                      MainUI.Close()
                                                                  End Sub)

                                     Else
                                         '退出检测
                                         Console.WriteLine("退出升级检测")
                                         Exit Do
                                     End If

                                 End If

                                 Exit For

                             Next

                         Catch ex As Exception
                             Console.WriteLine(ex)
                         End Try

                         Threading.Thread.Sleep(LoopIntervalSec * 1000)

                     Loop While enableLoop

                 End Sub)

    End Sub

End Class
