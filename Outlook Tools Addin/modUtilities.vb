Imports System.Windows.Forms
Imports System.Collections

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

            If TypeOf x Is Date And TypeOf y Is Date Then
                returnVal = DateTime.Compare(x, y)
            ElseIf TypeOf x Is Date And TypeOf y Is Date Then
                returnVal = x.CompareTo(y)
            Else
                ' If not numeric and not date then compare as string
                returnVal = [String].Compare(CType(x, ListViewItem).SubItems(col).Text, CType(y, ListViewItem).SubItems(col).Text)
            End If

            Return returnVal * sortOrderModifier

        End Function
    End Class
End Module

