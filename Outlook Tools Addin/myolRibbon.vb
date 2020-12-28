Imports Microsoft.Office.Tools.Ribbon

Public Class myolRibbon
    Dim fmSearch As frmSearch


    Private Sub btnSearch_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSearch.Click
        ShowForm(fmSearch, GetType(frmSearch))
        BringFormsToFront()
    End Sub
End Class
