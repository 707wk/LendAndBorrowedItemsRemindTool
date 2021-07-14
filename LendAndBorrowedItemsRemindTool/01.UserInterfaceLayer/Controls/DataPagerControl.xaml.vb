Public Class DataPagerControl

    Private _recordCount As Integer
    ''' <summary>
    ''' 总记录数
    ''' </summary>
    Public ReadOnly Property RecordCount As Integer
        Get
            Return _recordCount
        End Get
    End Property

    Private _pageSize As Integer
    ''' <summary>
    ''' 每页记录数
    ''' </summary>
    Public ReadOnly Property PageSize As Integer
        Get
            Return _pageSize
        End Get
    End Property

    Private _pageCount As Integer
    ''' <summary>
    ''' 总页数
    ''' </summary>
    Public ReadOnly Property PageCount As Integer
        Get
            Return _pageCount
        End Get
    End Property

    ''' <summary>
    ''' 当前页码
    ''' </summary>
    Public ReadOnly Property NowPageIndex As Integer
        Get
            Return PageIDList.SelectedItem
        End Get
    End Property

    Public Sub Init(ByVal rCount As Integer,
                    ByVal pSize As Integer)

        _recordCount = rCount
        _pageSize = pSize

        Dim tmpPageCount As Integer = Math.Ceiling(RecordCount / PageSize)
        _pageCount = If(tmpPageCount = 0, 1, tmpPageCount)

        PageIDList.Items.Clear()

        For i001 = 1 To PageCount
            PageIDList.Items.Add(i001)
        Next

        PageIDList.SelectedIndex = 0

        DataPagerInfo.Text = $"共 {RecordCount:n0} 条记录,每页 {PageSize:n0} 条,共 {PageCount:n0} 页"

    End Sub

    Public Sub Init(ByVal rCount As Integer,
                    ByVal pSize As Integer,
                    ByVal showPageIndex As Integer)

        _recordCount = rCount
        _pageSize = pSize

        Dim tmpPageCount As Integer = Math.Ceiling(RecordCount / PageSize)
        _pageCount = If(tmpPageCount = 0, 1, tmpPageCount)

        PageIDList.Items.Clear()

        For i001 = 1 To PageCount
            PageIDList.Items.Add(i001)
        Next

        If showPageIndex < 1 Then
            showPageIndex = 1
        End If

        If showPageIndex <= PageCount Then

        Else
            showPageIndex = PageCount
        End If

        PageIDList.SelectedIndex = showPageIndex - 1

        DataPagerInfo.Text = $"共 {RecordCount:n0} 条记录,每页 {PageSize:n0} 条,共 {PageCount:n0} 页"

    End Sub

    ''' <summary>
    ''' 页码改变
    ''' </summary>
    ''' <param name="pIndex">页码</param>
    ''' <param name="pSize">分页大小</param>
    Public Delegate Sub PageIndexChangedHandle(ByVal pIndex As Integer,
                                               ByVal pSize As Integer)
    ''' <summary>
    ''' 页码改变
    ''' </summary>
    Public Event PageIndexChanged As PageIndexChangedHandle

    Private Sub PageIDList_SelectionChanged(sender As Object,
                                            e As SelectionChangedEventArgs)

        If PageIDList.SelectedItem Is Nothing Then
            Exit Sub
        End If

        RaiseEvent PageIndexChanged(PageIDList.SelectedItem, PageSize)
    End Sub

    Private Sub FirstPageButton_Click(sender As Object,
                                      e As RoutedEventArgs)
        PageIDList.SelectedIndex = 0
    End Sub

    Private Sub PreviousPageButton_Click(sender As Object,
                                         e As RoutedEventArgs)
        If PageIDList.SelectedIndex > 0 Then
            PageIDList.SelectedIndex -= 1
        End If
    End Sub

    Private Sub NextPageButton_Click(sender As Object,
                                     e As RoutedEventArgs)
        If PageIDList.SelectedIndex < PageIDList.Items.Count - 1 Then
            PageIDList.SelectedIndex += 1
        End If
    End Sub

    Private Sub LastPageButton_Click(sender As Object,
                                     e As RoutedEventArgs)
        PageIDList.SelectedIndex = PageIDList.Items.Count - 1
    End Sub

End Class
