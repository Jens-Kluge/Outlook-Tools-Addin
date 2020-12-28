Imports System.Windows.Forms

Public Class frmAttachments
    Public mAttachments As List(Of Outlook.Attachment)

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmAttachments_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DefineColHeaders()
    End Sub

    Sub DefineColHeaders()
        lvAttachments.View = View.Details
        Dim ch As ColumnHeader

        ch = lvAttachments.Columns.Add("filename", width:=150)
        ch = lvAttachments.Columns.Add("display name", width:=150)
        ch = lvAttachments.Columns.Add("size (Bytes)", width:=120, textAlign:=HorizontalAlignment.Right)
        ch = lvAttachments.Columns.Add("extension", width:=50)

    End Sub

    Sub PopulateList()
        Dim lvi As ListViewItem
        Dim i As Integer = 0
        Try
            lvAttachments.Items.Clear()
            lvAttachments.ListViewItemSorter = Nothing
            lvAttachments.BeginUpdate()

            For Each attIt In mAttachments
                i += 1
                lvi = New ListViewItem() With {.Name = i, .Text = attIt.FileName}
                lvAttachments.Items.Add(lvi)
                lvi.SubItems.Add(attIt.DisplayName)
                lvi.SubItems.Add(attIt.Size)
                If InStr(attIt.FileName, ".") > 0 Then
                    lvi.SubItems.Add(Split(attIt.FileName, ".")(1))
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            lvAttachments.EndUpdate()
        End Try

    End Sub

    Private Sub lvAttachments_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles lvAttachments.ColumnClick

        If lvAttachments.Sorting = SortOrder.Ascending Then
            lvAttachments.Sorting = SortOrder.Descending
            lvAttachments.ListViewItemSorter = New ListViewItemComparer(e.Column, lvAttachments.Sorting)
        Else
            lvAttachments.Sorting = SortOrder.Ascending
            lvAttachments.ListViewItemSorter = New ListViewItemComparer(e.Column, lvAttachments.Sorting)
        End If
    End Sub

End Class