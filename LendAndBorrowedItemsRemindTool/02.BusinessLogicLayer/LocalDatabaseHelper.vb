Imports System.Data
Imports System.Data.SQLite
Imports Newtonsoft.Json
''' <summary>
''' 本地数据库辅助模块
''' </summary>
Public NotInheritable Class LocalDatabaseHelper

    Private Shared _DatabaseConnection As SQLite.SQLiteConnection
    Private Shared ReadOnly Property DatabaseConnection() As SQLiteConnection
        Get
            If _DatabaseConnection Is Nothing Then
                _DatabaseConnection = New SQLite.SQLiteConnection With {
                    .ConnectionString = AppSettingHelper.SQLiteConnection
                }
                _DatabaseConnection.Open()

                'Init()

            End If

            Return _DatabaseConnection
        End Get
    End Property

    Public Shared Sub Close()
        Try
            _DatabaseConnection?.Close()

#Disable Warning CA1031 ' Do not catch general exception types
        Catch ex As Exception
#Enable Warning CA1031 ' Do not catch general exception types
        End Try

    End Sub

#Region "初始化数据库"
    ''' <summary>
    ''' 初始化数据库
    ''' </summary>
    Public Shared Sub Init()

        Using cmd As New SQLite.SQLiteCommand(DatabaseConnection)
            cmd.CommandText = "
--关闭同步
PRAGMA synchronous = OFF;
--不记录日志
PRAGMA journal_mode = OFF;"

            cmd.ExecuteNonQuery()
        End Using

    End Sub
#End Region

#Region "配置值是否存在"
    ''' <summary>
    ''' 配置值是否存在
    ''' </summary>
    Public Shared Function OptionExists(key As String) As Boolean

        Using cmd As New SQLiteCommand(DatabaseConnection) With {
            .CommandText = "select
*

from AppSettingInfo
where Key=@Key"
        }
            cmd.Parameters.Add(New SQLiteParameter("@Key", DbType.String) With {.Value = key})

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                Return reader.Read
            End Using
        End Using

    End Function
#End Region

#Region "设置配置值"
    ''' <summary>
    ''' 设置配置值
    ''' </summary>
    Public Shared Sub SetOption(key As String, value As Object)

        If OptionExists(key) Then
            UpdateOption(key, JsonConvert.SerializeObject(value))
        Else
            AddOption(key, JsonConvert.SerializeObject(value))
        End If

    End Sub

#Region "新增配置值"
    ''' <summary>
    ''' 新增配置值
    ''' </summary>
    Private Shared Sub AddOption(key As String, value As String)

        Using cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "insert into
AppSettingInfo 
values(
@Key,
@Value,
@LastModifyTime
)"
        }
            cmd.Parameters.Add(New SQLiteParameter("@Key", DbType.String) With {.Value = key})
            cmd.Parameters.Add(New SQLiteParameter("@Value", DbType.String) With {.Value = value})
            cmd.Parameters.Add(New SQLiteParameter("@LastModifyTime", DbType.DateTime) With {.Value = Now})

            cmd.ExecuteNonQuery()
        End Using

    End Sub
#End Region

#Region "更新配置值"
    ''' <summary>
    ''' 更新配置值
    ''' </summary>
    Private Shared Sub UpdateOption(key As String, value As String)

        Using cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "update
AppSettingInfo 
set 
Value=@Value,
LastModifyTime=@LastModifyTime
where Key=@Key"
        }

            cmd.Parameters.Add(New SQLiteParameter("@Key", DbType.String) With {.Value = key})
            cmd.Parameters.Add(New SQLiteParameter("@Value", DbType.String) With {.Value = value})
            cmd.Parameters.Add(New SQLiteParameter("@LastModifyTime", DbType.DateTime) With {.Value = Now})

            cmd.ExecuteNonQuery()
        End Using

    End Sub
#End Region

#End Region

#Region "获取配置值"
    ''' <summary>
    ''' 获取配置值
    ''' </summary>
    Public Shared Function GetOption(Of T)(key As String) As T

        Using cmd As New SQLiteCommand(DatabaseConnection) With {
            .CommandText = "select
Value

from AppSettingInfo
where Key=@Key"
        }
            cmd.Parameters.Add(New SQLiteParameter("@Key", DbType.String) With {.Value = key})

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                If reader.Read Then
                    Return JsonConvert.DeserializeObject(Of T)($"{reader(0)}")

                Else
                    Return Nothing

                End If

            End Using
        End Using

    End Function
#End Region

End Class
