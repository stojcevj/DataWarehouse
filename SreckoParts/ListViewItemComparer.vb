Public Class ListViewItemComparer
    Implements IComparer

    Private col As Integer
    Private order As SortOrder

    Public Sub New()
        col = 0
        order = SortOrder.Ascending
    End Sub

    Public Sub New(ByVal column As Integer, ByVal order As SortOrder)
        col = column
        Me.order = order
    End Sub

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare

        Dim returnVal As Integer

        Try

            If IsNumeric(CType(x, ListViewItem).SubItems(col).Text) And
                            IsNumeric(CType(y, ListViewItem).SubItems(col).Text) Then


                returnVal = Val(CType(x, ListViewItem).SubItems(col).Text).CompareTo( _
                  Val(CType(y, ListViewItem).SubItems(col).Text))

            Else

                returnVal = [String].Compare(CType(x,  _
                                       ListViewItem).SubItems(col).Text, CType(y, ListViewItem).SubItems(col).Text)
            End If


        Catch ex As Exception


            

        End Try


        If order = SortOrder.Descending Then
            returnVal *= -1
        End If

        Return returnVal

    End Function

End Class