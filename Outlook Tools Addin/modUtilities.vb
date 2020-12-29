Imports System.Windows.Forms
Imports System.Collections
Imports System.Diagnostics
Imports System.Globalization

Module modUtilities
    Public Sub BringFormsToFront()
#Region "Windows Forms"
        Dim fms As FormCollection = Application.OpenForms

        If (fms Is Nothing) Then
            Return
        End If

        For Each fm As Form In fms
            fm.BringToFront()
        Next

    End Sub

    ''' <summary>
    ''' Create the form if it does not exist and show it
    ''' </summary>
    ''' <param name="fm"></param>
    ''' <param name="formType"></param>
    ''' <param name="modal"></param>
    Public Sub ShowForm(ByRef fm As Windows.Forms.Form, formType As Type, Optional modal As Boolean = False)


        If fm Is Nothing OrElse fm.IsDisposed() Then
            fm = Activator.CreateInstance(formType)
        End If

        If fm.Visible Then
            fm.BringToFront()
        Else
            If modal Then
                fm.ShowDialog()
            Else
                fm.Show()
            End If
        End If

    End Sub
#End Region

    ' Implements the manual sorting of items by columns.
    Class ListViewItemComparer
        Implements IComparer

        Private col As Integer
        Private Shared sortOrderModifier As Integer = 1

        Public Sub New()
            col = 0
        End Sub

        Public Sub New(column As Integer, srtOrder As SortOrder)
            col = column
            If (srtOrder = SortOrder.Descending) Then
                sortOrderModifier = -1
            ElseIf (srtOrder = SortOrder.Ascending) Then
                sortOrderModifier = 1
            End If
        End Sub

        Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
            Dim returnVal As Integer
            'debug.print(CType(x, ListViewItem).SubItems(col).GetType().Name)
            If TypeOf x Is Date And TypeOf y Is Date Then
                returnVal = DateTime.Compare(x, y)
            ElseIf (TypeOf x Is Double And TypeOf y Is Double) Or (TypeOf x Is Integer And TypeOf y Is Integer) Then
                returnVal = x.CompareTo(y)
            Else
                ' If not numeric and not date then compare as string
                returnVal = [String].Compare(CType(x, ListViewItem).SubItems(col).Text, CType(y, ListViewItem).SubItems(col).Text)
            End If

            Return returnVal * sortOrderModifier

        End Function
    End Class

    Public Class ColumnSorter
        Implements IComparer

#Region "Properties"
        Private _sortColumn As Integer

        Public Property SortColumn As Integer

            Set(ByVal Value As Integer)
                _sortColumn = Value
            End Set
            Get
                Return _sortColumn
            End Get
        End Property

        Private _sortOrder As SortOrder

        Public Property Order As SortOrder
            Set(ByVal value As SortOrder)
                _sortOrder = value
            End Set
            Get
                Return _sortOrder
            End Get
        End Property

        Private listViewItemComparer As Comparer
#End Region

#Region "class constructors"
        Public Sub New()

            _sortColumn = 0

            _sortOrder = SortOrder.None

            listViewItemComparer = New Comparer(CultureInfo.CurrentUICulture)

        End Sub


        Public Sub New(column As Integer, srtOrder As SortOrder)
            _sortColumn = column
            _sortOrder = srtOrder
        End Sub
#End Region

        ''' <summary>
        ''' This method Is inherited from the IComparer interface.  It compares the two objects passed using a case insensitive comparison.
        ''' </summary>
        ''' <param name="x">First object to be compared</param>
        ''' <param name="y">Second object to be compared</param>
        ''' <returns>The result of the comparison. "0" if equal, negative if 'x' is less than 'y' and positive if 'x' is greater than 'y'</returns>
        Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare

            Try

                Dim lviX As ListViewItem = CType(x, ListViewItem)
                Dim lviY As ListViewItem = CType(y, ListViewItem)

                Dim compareResult As Integer = 0

                If (lviX.SubItems(SortColumn).Tag IsNot Nothing AndAlso lviY.SubItems(SortColumn).Tag IsNot Nothing) Then
                    compareResult = listViewItemComparer.Compare(lviX.SubItems(SortColumn).Tag, lviY.SubItems(SortColumn).Tag)
                Else
                    compareResult = listViewItemComparer.Compare(lviX.SubItems(SortColumn).Text, lviY.SubItems(SortColumn).Text)
                End If

                If _sortOrder = SortOrder.Ascending Then
                    Return compareResult
                ElseIf _sortOrder = SortOrder.Descending Then
                    Return (-compareResult)
                Else
                    Return 0
                End If
            Catch
                Return 0
            End Try
        End Function

    End Class
End Module

