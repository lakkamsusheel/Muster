Class ListViewComparer
    Implements IComparer

    Private m_ColumnNumber As Integer
    Private m_SortOrder As SortOrder
    Private m_SelectedColumn As String
    Public Sub New(ByVal column_number As Integer, ByVal sort_order As SortOrder, ByVal SelectedColumn As String)
        m_ColumnNumber = column_number
        m_SortOrder = sort_order
        m_SelectedColumn = SelectedColumn
    End Sub
    ' Compare the items in the appropriate column
    ' for objects x and y.
    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
        Dim item_x As ListViewItem = DirectCast(x, ListViewItem)
        Dim item_y As ListViewItem = DirectCast(y, ListViewItem)

        ' Get the sub-item values.
        Dim string_x As String
        Dim string_y As String
        Dim Int_x As Integer = 0
        Dim Int_y As Integer = 0

        If item_x.SubItems.Count <= m_ColumnNumber Then
            If m_SelectedColumn.IndexOf("Owner ID") > 0 Or m_SelectedColumn.IndexOf("Owner Points") > 0 Or _
               m_SelectedColumn.IndexOf("Facility ID") > 0 Or m_SelectedColumn.IndexOf("Facility Points") > 0 Or m_SelectedColumn.IndexOf("SNo") > 0 Then
                Int_x = 0
            Else
                string_x = ""
            End If
        Else
            If m_SelectedColumn.IndexOf("Owner ID") > 0 Or m_SelectedColumn.IndexOf("Owner Points") > 0 Or _
               m_SelectedColumn.IndexOf("Facility ID") > 0 Or m_SelectedColumn.IndexOf("Facility Points") > 0 Or m_SelectedColumn.IndexOf("SNo") > 0 Then
                Int_x = CInt(item_x.SubItems(m_ColumnNumber).Text)
            Else
                string_x = item_x.SubItems(m_ColumnNumber).Text
            End If
        End If


        If item_y.SubItems.Count <= m_ColumnNumber Then
            If m_SelectedColumn.IndexOf("Owner ID") > 0 Or m_SelectedColumn.IndexOf("Owner Points") > 0 Or _
               m_SelectedColumn.IndexOf("Facility ID") > 0 Or m_SelectedColumn.IndexOf("Facility Points") > 0 Or m_SelectedColumn.IndexOf("SNo") > 0 Then
                Int_y = 0
            Else
                string_y = ""
            End If
        Else
            If m_SelectedColumn.IndexOf("Owner ID") > 0 Or m_SelectedColumn.IndexOf("Owner Points") > 0 Or _
               m_SelectedColumn.IndexOf("Facility ID") > 0 Or m_SelectedColumn.IndexOf("Facility Points") > 0 Or m_SelectedColumn.IndexOf("SNo") > 0 Then
                Int_y = CInt(item_y.SubItems(m_ColumnNumber).Text)
            Else
                string_y = item_y.SubItems(m_ColumnNumber).Text
            End If
        End If

        If m_SortOrder = SortOrder.Ascending Then
            If m_SelectedColumn.IndexOf("Owner ID") > 0 Or m_SelectedColumn.IndexOf("Owner Points") > 0 Or _
               m_SelectedColumn.IndexOf("Facility ID") > 0 Or m_SelectedColumn.IndexOf("Facility Points") > 0 Or m_SelectedColumn.IndexOf("SNo") > 0 Then
                If Int_x < Int_y Then
                    Return -1
                ElseIf Int_x = Int_y Then
                    Return 0
                Else
                    Return 1
                End If
            Else
                Return String.Compare(string_x, string_y)
            End If
        Else
            If m_SelectedColumn.IndexOf("Owner ID") > 0 Or m_SelectedColumn.IndexOf("Owner Points") > 0 Or _
               m_SelectedColumn.IndexOf("Facility ID") > 0 Or m_SelectedColumn.IndexOf("Facility Points") > 0 Or m_SelectedColumn.IndexOf("SNo") > 0 Then
                If Int_y < Int_x Then
                    Return -1
                ElseIf Int_y = Int_x Then
                    Return 0
                Else
                    Return 1
                End If
            Else
                Return String.Compare(string_y, string_x)
            End If
        End If
    End Function
    ' Function to Sort by Strings

    'Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
    '    Dim item_x As ListViewItem = DirectCast(x, ListViewItem)
    '    Dim item_y As ListViewItem = DirectCast(y, ListViewItem)

    '    ' Get the sub-item values.
    '    Dim string_x As String
    '    If item_x.SubItems.Count <= m_ColumnNumber Then
    '        string_x = ""
    '    Else
    '        string_x = item_x.SubItems(m_ColumnNumber).Text
    '    End If

    '    Dim string_y As String
    '    If item_y.SubItems.Count <= m_ColumnNumber Then
    '        string_y = ""
    '    Else
    '        string_y = item_y.SubItems(m_ColumnNumber).Text
    '    End If

    '    ' Compare them.
    '    If m_SortOrder = SortOrder.Ascending Then
    '        Return String.Compare(string_x, string_y)
    '    Else
    '        Return String.Compare(string_y, string_x)
    '    End If
    'End Function

End Class