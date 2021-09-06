Imports System.Globalization
Imports System.Windows.Threading
Imports Microsoft.AppCenter
Imports Microsoft.AppCenter.Analytics
Imports Microsoft.AppCenter.Crashes

Class Application
    Private Sub Application_DispatcherUnhandledException(sender As Object, e As DispatcherUnhandledExceptionEventArgs) Handles Me.DispatcherUnhandledException

        Application_Exit(Nothing, Nothing)

        AppSettingHelper.Instance.Logger.Error(e.Exception)

        MsgBox($"应用程序中发生了未处理的异常 :
{e.Exception.Message}

点击""确定"", 应用程序将立即关闭, 具体异常信息可在 \Logs\Error 文件夹内查看",
               MsgBoxStyle.Critical)

    End Sub

    Private Sub Application_Exit(sender As Object, e As ExitEventArgs) Handles Me.[Exit]

        AppSettingHelper.Instance.ClearTempFiles()

    End Sub

    Private Sub Application_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup

        Dim countryCode = RegionInfo.CurrentRegion.TwoLetterISORegionName
        AppCenter.SetCountryCode(countryCode)

        '使用调试器时不记录数据
        Analytics.SetEnabledAsync(Not Debugger.IsAttached)

        AppCenter.Start(AppSettingHelper.AppKey,
                        GetType(Analytics),
                        GetType(Crashes))

        Analytics.TrackEvent("程序启动")
        AppSettingHelper.Instance.Logger.Info("程序启动")

    End Sub

End Class
