Imports System.Windows.Forms
Module modUtilities
    Public Sub BringFormsToFront()

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
End Module

