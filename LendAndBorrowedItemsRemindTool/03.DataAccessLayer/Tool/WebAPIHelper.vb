Imports System.Net.Http
Imports ERPInfoServiceLib
Imports Newtonsoft.Json

''' <summary>
''' 请求辅助模块
''' </summary>
Public NotInheritable Class WebAPIHelper

    Private Shared tmpHttpClient As New HttpClient With {
        .Timeout = New TimeSpan(0, 0, 10)
    }

    Public Shared Function GetData(Of T)(url As String) As ReceiveMsg(Of T)

        Try

            tmpHttpClient.DefaultRequestHeaders.TryAddWithoutValidation("ContentType", "application/json")

            Dim UnixTimeMilliseconds = DateTimeOffset.Now.ToUnixTimeMilliseconds
            tmpHttpClient.DefaultRequestHeaders.Remove(DingTalkAuthorizationHelper.RequestHeader)
            tmpHttpClient.DefaultRequestHeaders.Add(DingTalkAuthorizationHelper.RequestHeader, DingTalkAuthorizationHelper.CalcSignature("K4E997XGgxjAfCGfC232Yf3HZk8WwJhF", UnixTimeMilliseconds))

            tmpHttpClient.DefaultRequestHeaders.Remove(DingTalkAuthorizationHelper.RequestHeaderTimestamp)
            tmpHttpClient.DefaultRequestHeaders.Add(DingTalkAuthorizationHelper.RequestHeaderTimestamp, UnixTimeMilliseconds)

            Dim tmpResponse = tmpHttpClient.GetAsync(url).GetAwaiter.GetResult
            tmpResponse.EnsureSuccessStatusCode()

            '接收数据
            Dim contentStr = tmpResponse.Content.ReadAsStringAsync().GetAwaiter.GetResult
            Return JsonConvert.DeserializeObject(Of ReceiveMsg(Of T))(contentStr)

        Catch ex As Exception
            Return New ReceiveMsg(Of T) With {.Code = 404, .Message = ex.Message}
        End Try

    End Function

End Class
