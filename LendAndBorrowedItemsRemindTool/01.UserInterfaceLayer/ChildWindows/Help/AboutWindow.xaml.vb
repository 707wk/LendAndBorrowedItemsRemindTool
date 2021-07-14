Imports System.IO

Public Class AboutWindow
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)

        UIHelper.InitChildWindowStyle(Me)

        TitleText.Text = My.Application.Info.Title

#If DEBUG Then
        ProductVersion.Text = $"版本 {AppSettingHelper.Instance.ProductVersion}_{If(Environment.Is64BitProcess, "64", "32")}Bit_Debug"

#Else
        ProductVersion.Text = $"版本 {AppSettingHelper.Instance.ProductVersion}_{If(Environment.Is64BitProcess, "64", "32")}Bit_Release"

#End If

        Copyright.Text = My.Application.Info.Copyright

    End Sub
End Class
