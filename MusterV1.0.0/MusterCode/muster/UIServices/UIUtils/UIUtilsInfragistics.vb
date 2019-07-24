Public Class UIUtilsInfragistics
    '
    ' Check that the double click occurred on a row!
    '  Code taken from http://devcenter.infragistics.com/Support/KnowledgeBaseArticle.Aspx?ArticleID=5592
    '
    Public Shared Function WinGridRowDblClicked(ByVal sender As Object, ByVal e As System.EventArgs) As Boolean
        Dim grid As Infragistics.Win.UltraWinGrid.UltraGrid = DirectCast(sender, Infragistics.Win.UltraWinGrid.UltraGrid)

        'Get the last element that the mouse entered
        Dim lastElementEntered As Infragistics.Win.UIElement = grid.DisplayLayout.UIElement.LastElementEntered

        Dim rowElement As Infragistics.Win.ultrawingrid.RowUIElement
        If TypeOf lastElementEntered Is Infragistics.Win.ultrawingrid.RowUIElement Then
            rowElement = DirectCast(lastElementEntered, Infragistics.Win.ultrawingrid.RowUIElement)
        Else
            rowElement = DirectCast(lastElementEntered.GetAncestor(GetType(Infragistics.Win.ultrawingrid.RowUIElement)), Infragistics.Win.ultrawingrid.RowUIElement)
        End If

        If rowElement Is Nothing Then
            Return False
        Else
            Return True
        End If
    End Function
End Class
