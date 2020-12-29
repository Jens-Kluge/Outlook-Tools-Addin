Imports System.Collections
Imports System.Windows.Forms
Imports System.Drawing
Imports outlook = Microsoft.Office.Interop.Outlook

Public Class frmSearch

    Private fmAttachments As frmAttachments
    Private m_lstColumnSorter As ColumnSorter = New ColumnSorter()

#Region "Folder view handling"

    Private Sub btnLoadFolders_Click(sender As Object, e As EventArgs) Handles btnLoadFolders.Click

        Dim app As outlook.Application = Globals.ThisAddIn.Application
        Dim olNs As outlook.NameSpace
        olNs = app.GetNamespace("MAPI")

        Dim RootFolder As outlook.Folder
        Dim nd As TreeNode
        Dim rootKey As String
        'Set rootFolder = olNs.GetDefaultFolder(OlDefaultFolders.olFolderInbox)

        tvFolders.Nodes.Clear()
        Me.Cursor = Cursors.WaitCursor

        For Each RootFolder In olNs.Folders
            Try
                rootKey = RootFolder.EntryID & "|" & RootFolder.StoreID
                nd = tvFolders.Nodes.Add(key:=rootKey, text:=RootFolder.Name)
                ListSubFolders(nd, RootFolder, rootKey)
                nd.ExpandAll()
            Catch 'if exchange server does not respond continue with next folder
            End Try
        Next

        Me.Cursor = Cursors.Default

    End Sub

    ''' <summary>
    ''' Show all subfolders in the treeview
    ''' recursive call as folder level depth is not known in advance 
    ''' </summary>
    ''' <param name="fldr"></param>
    ''' <param name="parentkey"></param>
    Sub ListSubFolders(nd As TreeNode, fldr As outlook.Folder, parentkey As String)
        Dim key As String
        Dim subnode As TreeNode
        Dim subfldr As outlook.Folder

        For Each subfldr In fldr.Folders
            key = subfldr.EntryID & "|" & subfldr.StoreID

            subnode = nd.Nodes.Add(key, subfldr.Name)
            'recursive call to list all levels
            ListSubFolders(subnode, subfldr, key)

        Next

    End Sub

    Private Sub tvFolders_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles tvFolders.NodeMouseDoubleClick
        SearchOLItems()
    End Sub

    ''' <summary>
    ''' Drag and drop support, show symbol when cursor enters control boundaries
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub tvFolders_DragEnter(sender As Object, e As DragEventArgs) Handles tvFolders.DragEnter
        If e.Data.GetDataPresent(GetType(ListView.SelectedListViewItemCollection)) Then
            e.Effect = DragDropEffects.Move
        End If
    End Sub

    Private Sub tvFolders_DragDrop(sender As Object, e As DragEventArgs) Handles tvFolders.DragDrop
        If e.Data.GetDataPresent(GetType(ListView.SelectedListViewItemCollection).ToString(), False) Then
            Dim loc As Point = (CType(sender, TreeView)).PointToClient(New Point(e.X, e.Y))
            Dim destNode As TreeNode = (CType(sender, TreeView)).GetNodeAt(loc)

            Dim newKey As String

            tvFolders.SelectedNode = destNode

            Dim lstViewColl As ListView.SelectedListViewItemCollection = CType(e.Data.GetData(GetType(ListView.SelectedListViewItemCollection)), ListView.SelectedListViewItemCollection)
            For Each lvItem As ListViewItem In lstViewColl
                newKey = MoveItemToFolder(lvItem.Name, destNode.Name)
                lvItem.Name = newKey
                'lvItem.Remove()
            Next lvItem
        End If
    End Sub

    Private Sub tvFolders_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles tvFolders.NodeMouseClick
        Dim app As outlook.Application = Globals.ThisAddIn.Application
        Dim fldr As outlook.Folder
        Dim key As String

        key = e.Node.Name
        fldr = GetOLFolder(key)
        app.ActiveExplorer.CurrentFolder = fldr
    End Sub

#End Region

    Private Sub txtSearch_KeyDown(sender As Object, e As KeyEventArgs) Handles txtSearch.KeyDown
        If e.KeyCode = Keys.Enter Then
            SearchOLItems()
        End If
    End Sub

#Region "Result list handling"
    ''' <summary>
    ''' Loop through all folder and subfolder items one by one and search for the searchstring
    ''' </summary>
    Sub SearchOLItems()

        Dim SubFolderName As String
        Dim strArr() As String
        Dim olNs As outlook.NameSpace
        Dim app As outlook.Application = Globals.ThisAddIn.Application
        Dim olFldr As outlook.Folder

        If tvFolders.SelectedNode Is Nothing Then Exit Sub

        Try
            Me.Cursor = Cursors.WaitCursor

            olNs = app.GetNamespace("MAPI")
            SubFolderName = tvFolders.SelectedNode.Text
            strArr = Split(tvFolders.SelectedNode.Name, "|")

            olFldr = olNs.GetFolderFromID(strArr(0), strArr(1))

            lstResults.Items.Clear()
            lstResults.ListViewItemSorter = Nothing
            lstResults.BeginUpdate()
            ListFldrItems(olFldr, txtSearch.Text)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            lstResults.EndUpdate()
            lstResults.ListViewItemSorter = m_lstColumnSorter
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    ''' <summary>
    ''' List all Items in an outlook folder recursively
    ''' </summary>
    ''' <param name="SearchFolder"></param>
    ''' <param name="searchString"></param>
    Sub ListFldrItems(SearchFolder As outlook.Folder, searchString As String)

        Dim myitems As outlook.Items
        Dim myitem As Object
        Dim Found As Boolean

        myitems = SearchFolder.Items

        Found = False
        'define outlook object for reading out the various item types

        Dim mailIt As outlook.MailItem
        Dim cntIt As outlook.ContactItem
        Dim jnlIt As outlook.JournalItem
        Dim tskIt As outlook.TaskItem
        Dim lvi As ListViewItem
        Dim lsi As ListViewItem.ListViewSubItem

        For Each myitem In SearchFolder.Items

            If myitem.Class = outlook.OlObjectClass.olMail Then
                mailIt = myitem

                If searchString = "" OrElse InStr(1, mailIt.Subject, searchString) > 0 Then

                    lvi = New ListViewItem() With {.Name = mailIt.EntryID & "|" & SearchFolder.StoreID, .Text = mailIt.Subject}
                    lstResults.Items.Add(lvi)
                    lvi.SubItems.Add(mailIt.SenderName)
                    If mailIt.Recipients.Count > 0 Then
                        lvi.SubItems.Add(mailIt.Recipients.Item(1).Name)
                    Else
                        lvi.SubItems.Add("")
                    End If
                    lsi = lvi.SubItems.Add(text:=mailIt.SentOn)
                    lsi.Tag = mailIt.SentOn
                    lsi = lvi.SubItems.Add(text:=mailIt.ReceivedTime)
                    lsi.Tag = mailIt.ReceivedTime

                    Found = True
                End If
            ElseIf myitem.Class = outlook.OlObjectClass.olContact Then
                cntIt = myitem
                lvi = New ListViewItem() With {.Name = cntIt.EntryID & "|" & SearchFolder.StoreID, .Text = cntIt.LastName}
                lvi = lstResults.Items.Add(lvi)
            ElseIf myitem.Class = outlook.OlObjectClass.olJournal Then
                jnlIt = myitem
                lvi = New ListViewItem() With {.Name = jnlIt.EntryID & "|" & SearchFolder.StoreID, .Text = jnlIt.ConversationTopic}
                lvi = lstResults.Items.Add(lvi)
            ElseIf myitem.Class = outlook.OlObjectClass.olTask Then
                tskIt = myitem
                lvi = New ListViewItem() With {.Name = tskIt.EntryID & "|" & SearchFolder.StoreID, .Text = tskIt.ConversationTopic}
                lvi = lstResults.Items.Add(lvi)
            End If
        Next

        'now to the same for each subfolder
        Dim subfolder As outlook.Folder
        For Each subfolder In SearchFolder.Folders
            ListFldrItems(subfolder, searchString)
        Next

    End Sub

    Private Sub lstResults_DoubleClick(sender As Object, e As EventArgs) Handles lstResults.DoubleClick
        Dim mailIt As outlook.MailItem
        Dim app As outlook.Application = Globals.ThisAddIn.Application
        Dim key As String

        If lstResults.SelectedItems.Count = 0 Then Exit Sub

        key = lstResults.SelectedItems(0).Name
        mailIt = GetMailItem(key)

        If Not (mailIt Is Nothing) Then
            mailIt.GetInspector.Activate()
        End If

    End Sub

    Private Sub lstResults_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles lstResults.ColumnClick

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
        'myListView.SetSortIcon(m_lstColumnSorter.SortColumn, m_lstColumnSorter.Order)

    End Sub

    Private Sub frmSearch_Load(sender As Object, e As EventArgs) Handles Me.Load

        lstResults.View = View.Details
        Dim ch As ColumnHeader

        ch = lstResults.Columns.Add("subject", width:=150)
        ch = lstResults.Columns.Add("from", width:=120)
        ch = lstResults.Columns.Add("to", width:=120)
        ch = lstResults.Columns.Add("sent", width:=120)
        ch = lstResults.Columns.Add("received", width:=120)

    End Sub

    Private Sub lstResults_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) Handles lstResults.ItemSelectionChanged
        Dim mailIt As outlook.MailItem
        Dim app As outlook.Application = Globals.ThisAddIn.Application
        Dim key As String

        key = e.Item.Name
        mailIt = GetMailItem(key)

        If Not (mailIt Is Nothing) Then

            If app.ActiveExplorer.IsItemSelectableInView(mailIt) Then
                app.ActiveExplorer.ClearSelection()
                app.ActiveExplorer.AddToSelection(mailIt)
                app.ActiveExplorer.CurrentFolder.Display()
            End If
        End If

    End Sub

    'Drag and drop support
    Private Sub lstResults_ItemDrag(sender As Object, e As ItemDragEventArgs) Handles lstResults.ItemDrag
        lstResults.DoDragDrop(lstResults.SelectedItems, DragDropEffects.Move)
    End Sub

#End Region

    Function GetMailItem(ByVal key As String) As outlook.MailItem
        Dim IDs() As String
        Dim olNs As outlook.NameSpace
        Dim app As outlook.Application = Globals.ThisAddIn.Application
        Dim olItem As Object
        Dim mlIt As outlook.MailItem
        olNs = app.GetNamespace("MAPI")
        IDs = Split(key, "|")
        Try
            olItem = olNs.GetItemFromID(IDs(0), IDs(1))
        Catch ex As Exception
            Return Nothing
        End Try

        If olItem.Class = outlook.OlObjectClass.olMail Then
            mlIt = olItem
            GetMailItem = mlIt
        Else
            GetMailItem = Nothing
        End If
    End Function

    Function GetOLFolder(key As String) As outlook.Folder
        Dim app As outlook.Application = Globals.ThisAddIn.Application
        Dim IDs() As String
        Dim olFolder As outlook.Folder

        Dim olNs As outlook.NameSpace = app.GetNamespace("MAPI")
        IDs = Split(key, "|")
        Try
            olFolder = olNs.GetFolderFromID(IDs(0), IDs(1))
            Return olFolder
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Moves the mailitem specified by itemkey into the folder specified by folderkey
    ''' Returns the key of the moved item, where key = EntryID|StoreID
    ''' </summary>
    ''' <param name="ItemKey"></param>
    ''' <param name="FolderKey"></param>
    ''' <returns></returns>
    Function MoveItemToFolder(ItemKey As String, FolderKey As String) As String
        Dim mlIt As outlook.MailItem
        Dim fldr As outlook.Folder
        Dim newItem As outlook.MailItem

        mlIt = GetMailItem(ItemKey)
        fldr = GetOLFolder(FolderKey)
        If mlIt Is Nothing Or fldr Is Nothing Then Return ""

        Try
            newItem = mlIt.Move(fldr)
            Return newItem.EntryID & "|" & fldr.StoreID
        Catch 'move was not successful => new Key is the old key
            Return ItemKey
        End Try

    End Function

    Private Sub bnAttachments_Click(sender As Object, e As EventArgs) Handles bnAttachments.Click
        Dim AttList As List(Of outlook.Attachment)

        Me.Cursor = Cursors.WaitCursor
        AttList = FindAttachments()

        If AttList.Count > 0 Then

            ShowForm(fmAttachments, GetType(frmAttachments))
            fmAttachments.mAttachments = AttList
            fmAttachments.PopulateList()

        End If

        Me.Cursor = Cursors.Default
    End Sub

    ''' <summary>
    ''' Find the attachments of all mail items in the resultset
    ''' </summary>
    Function FindAttachments() As List(Of outlook.Attachment)
        Dim key As String
        Dim mlIt As outlook.MailItem
        Dim mlAtt As outlook.Attachment
        Dim mlAtts As New List(Of outlook.Attachment)
        Try
            For Each lvi As ListViewItem In lstResults.Items
                key = lvi.Name
                mlIt = GetMailItem(key)
                For Each mlAtt In mlIt.Attachments
                    mlAtts.Add(mlAtt)
                Next
            Next
            Return mlAtts
        Catch
            Return Nothing
        End Try
    End Function


End Class