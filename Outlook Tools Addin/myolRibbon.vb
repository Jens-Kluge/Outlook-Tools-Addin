Imports Microsoft.Office.Tools.Ribbon

Public Class myolRibbon
    Dim fmSearch As frmSearch

    Private Sub myolRibon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnSearch_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSearch.Click
        ShowForm(fmSearch, GetType(frmSearch))
        BringFormsToFront()
    End Sub
End Class
