Imports System.Windows.Forms
Imports System.IO
Imports System.Diagnostics

Public Class frmAttachments
    Public mAttachments As List(Of Outlook.Attachment)
    Private m_lstColumnSorter As ColumnSorter = New ColumnSorter()

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
        Dim lsi As ListViewItem.ListViewSubItem

        Dim i As Integer = 0
        Try
            lvAttachments.Items.Clear()
            lvAttachments.ListViewItemSorter = Nothing
            lvAttachments.BeginUpdate()

            For Each attIt In mAttachments
                lvi = New ListViewItem() With {.Name = i, .Text = attIt.FileName}
                lvi.Tag = i
                lvAttachments.Items.Add(lvi)
                lvi.SubItems.Add(attIt.DisplayName)
                lsi = lvi.SubItems.Add(attIt.Size)
                'set the tag value for sorting purposes
                lsi.Tag = CType(attIt.Size, Integer)
                If InStr(attIt.FileName, ".") > 0 Then
                    lvi.SubItems.Add(Split(attIt.FileName, ".")(1))
                End If
                i += 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            lvAttachments.EndUpdate()
            lvAttachments.ListViewItemSorter = m_lstColumnSorter
        End Try

    End Sub

    Private Sub lvAttachments_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles lvAttachments.ColumnClick

        Dim myListView As ListView = CType(sender, ListView)

        ' Determine if clicked column Is already the column that Is being sorted.
        If e.Column = m_lstColumnSorter.SortColumn Then
            '' Reverse the current sort direction for this column.
            If m_lstColumnSorter.Order = SortOrder.Ascending Then
                m_lstColumnSorter.Order = SortOrder.Descending
            Else
                m_lstColumnSorter.Order = SortOrder.Ascending
            End If
        Else
            ' Set the column number that Is to be sorted; default to ascending.
            m_lstColumnSorter.SortColumn = e.Column
            m_lstColumnSorter.Order = SortOrder.Ascending
        End If

        ' Perform the sort with these New sort options.
        myListView.Sort()

    End Sub

    Public Sub saveAtt(attIt As Outlook.Attachment, Optional bDisplay As Boolean = False)

        Dim saveFolder As String
        Dim dateFormat, FilePath As String
        Dim fi As FileInfo

        dateFormat = Now.ToString("yyyy-MM-dd H-mm")
        saveFolder = Path.GetTempPath

        FilePath = saveFolder & attIt.DisplayName
        Try
            attIt.SaveAsFile(FilePath)
            If bDisplay Then
                fi = New FileInfo(FilePath)
                FileInfoExtensions.Unblock(fi)
                Process.Start(FilePath)

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub


    Public Sub UnblockAttachments()

        Dim saveFolder As String
        Dim dateFormat, FilePath As String
        Dim fi As FileInfo

        dateFormat = Now.ToString("yyyy-MM-dd H-mm")
        saveFolder = Path.GetTempPath

        Try
            For Each attIt In mAttachments
                FilePath = saveFolder & " " & attIt.DisplayName

                attIt.SaveAsFile(FilePath)
                fi = New FileInfo(FilePath)
                FileInfoExtensions.Unblock(fi)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub lvAttachments_DoubleClick(sender As Object, e As EventArgs) Handles lvAttachments.DoubleClick

        If lvAttachments.SelectedItems.Count = 0 Then Exit Sub

        saveAtt(mAttachments(lvAttachments.SelectedItems(0).Tag), True)

    End Sub


End Class