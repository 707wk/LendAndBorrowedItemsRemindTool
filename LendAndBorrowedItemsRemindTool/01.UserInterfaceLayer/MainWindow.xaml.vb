Imports OfficeOpenXml

Class MainWindow

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

        Me.Title = $"{My.Application.Info.Title} V{AppSettingHelper.Instance.ProductVersion}"

        '设置使用方式为个人使用
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim tmpAppCenterSparkle As New AppCenterSparkle(AppSettingHelper.AppKey, Me)
        tmpAppCenterSparkle.CheckUpdateAsync()

    End Sub

    Private Sub Shutdown()

        System.Windows.Application.Current.Shutdown()

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

End Class
