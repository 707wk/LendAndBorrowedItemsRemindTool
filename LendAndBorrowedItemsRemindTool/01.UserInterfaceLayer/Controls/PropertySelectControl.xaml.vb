Public Class PropertySelectControl

    Private _historyItemKey As String
    Public Property HistoryItemKey As String
        Get
            Return _historyItemKey
        End Get
        Set(value As String)
            If String.IsNullOrWhiteSpace(value) Then
                Exit Property
            End If

            _historyItemKey = value

            If AppSettingHelper.Instance.InputHistoryItems.ContainsKey(HistoryItemKey) Then
                '有记录
                If IsEditable Then
                    '可编辑状态
                    Dim values = AppSettingHelper.Instance.InputHistoryItems(HistoryItemKey)
                    For Each item In values
                        PropertySelectComboBox.Items.Add(item)
                    Next

                    '显示第一个值
                    If PropertySelectComboBox.Items.Count > 0 AndAlso
                        Not String.IsNullOrWhiteSpace(values(0)) Then

                        PropertySelectComboBox.SelectedIndex = 0
                    End If

                Else
                    '不可编辑状态
                    Dim values = AppSettingHelper.Instance.InputHistoryItems(HistoryItemKey)
                    If values.Count = 0 Then
                        Exit Property
                    End If

                    Dim index = PropertySelectComboBox.Items.IndexOf(values(0))
                    PropertySelectComboBox.SelectedIndex = index

                End If

            Else
                '无记录
                AppSettingHelper.Instance.InputHistoryItems.Add(HistoryItemKey, New List(Of String))

            End If

        End Set
    End Property

    Public Property IsEditable As Boolean
        Get
            Return PropertySelectComboBox.IsEditable
        End Get
        Set(value As Boolean)
            PropertySelectComboBox.IsEditable = value
        End Set
    End Property

    Private _dataType As Integer
    ''' <summary>
    ''' 数据类型 0字符串,1整数,2小数
    ''' </summary>
    Public Property DataType As Integer
        Get
            Return _dataType
        End Get
        Set(value As Integer)
            _dataType = value
        End Set
    End Property

    ''' <summary>
    ''' 输入值
    ''' </summary>
    Public ReadOnly Property Value As String
        Get
            Return PropertySelectComboBox.Text
        End Get
    End Property

    Private Sub PropertySelectComboBox_TextChanged(sender As Object, e As TextChangedEventArgs)
        ClearButton.Visibility = If(Not String.IsNullOrWhiteSpace(PropertySelectComboBox.Text) OrElse
            PropertySelectComboBox.SelectedIndex >= 0,
            Visibility.Visible,
            Visibility.Hidden)
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As RoutedEventArgs)
        PropertySelectComboBox.Text = String.Empty
        PropertySelectComboBox.SelectedIndex = -1
    End Sub

    Private Sub PropertySelectComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        ClearButton.Visibility = If(Not String.IsNullOrWhiteSpace(PropertySelectComboBox.Text) OrElse
            PropertySelectComboBox.SelectedIndex >= 0,
            Visibility.Visible,
            Visibility.Hidden)
    End Sub

    ''' <summary>
    ''' 增加历史记录(保留前15个),只支持向Items添加数据
    ''' </summary>
    Public Sub AddHistoryValue()

        If String.IsNullOrWhiteSpace(_historyItemKey) Then
            Throw New NullReferenceException
        End If

        Dim tmpValue = Value

        If Not PropertySelectComboBox.IsEditable Then
            AppSettingHelper.Instance.InputHistoryItems(HistoryItemKey) = {tmpValue}.ToList
            Exit Sub
        End If

        If PropertySelectComboBox.Items.Contains(tmpValue) Then
            Dim valueIndex = PropertySelectComboBox.Items.IndexOf(tmpValue)
            PropertySelectComboBox.Items.Insert(0, tmpValue)
            PropertySelectComboBox.SelectedIndex = 0
            PropertySelectComboBox.Items.RemoveAt(valueIndex + 1)

        Else
            PropertySelectComboBox.Items.Insert(0, tmpValue)

        End If

        If PropertySelectComboBox.Items.Count > 15 Then
            PropertySelectComboBox.Items.RemoveAt(15)
        End If

        Dim values = (From item In PropertySelectComboBox.Items
                      Select $"{item}").ToList

        AppSettingHelper.Instance.InputHistoryItems(HistoryItemKey) = values

    End Sub

    ''' <summary>
    ''' 显示历史值
    ''' </summary>
    Public Sub ShowHistoryValueWhenNotEditable()

        If Not IsEditable Then

            '只读状态
            Dim values = AppSettingHelper.Instance.InputHistoryItems(HistoryItemKey)
            If values.Count = 0 Then
                Exit Sub
            End If

            Dim index = PropertySelectComboBox.Items.IndexOf(values(0))
            PropertySelectComboBox.SelectedIndex = index

        End If

    End Sub

End Class
