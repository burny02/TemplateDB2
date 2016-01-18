Public Class UniversalEventsClass
    Inherits CentralFunctions

    Protected Sub ComboKeyDown(sender As Object, e As KeyEventArgs)
        e.SuppressKeyPress = True
    End Sub

    Protected Sub DataGridViewDataError(sender As Object, e As DataGridViewDataErrorEventArgs)
        Try
            e.Cancel = False
            Call ErrorHandler(sender, e)
        Catch ex As Exception
            Call DefaultError(ex)
        End Try
    End Sub

    Protected Sub GridComboEnter(sender As Object, e As DataGridViewCellEventArgs)

        If Not e.RowIndex > 0 Then Exit Sub
        On Error Resume Next
        Dim dgv As DataGridView = CType(sender, DataGridView)

        If dgv(e.ColumnIndex, e.RowIndex).EditType.ToString() = "System.Windows.Forms.DataGridViewComboBoxEditingControl" Then
            SendKeys.Send("{F4}")
        End If
    End Sub

    Protected Sub FormClosing(sender As Object, e As FormClosingEventArgs)
        If UnloadData() = True Then e.Cancel = True
        Call Quitter()
    End Sub

    Protected Sub ErrorHandler(sender As Object, e As Object)

        If Not e.RowIndex > 0 Then Exit Sub

        Dim Obj As Object

        Try
            If TypeOf (sender) Is DataGridView Then
                Obj = CType(sender, DataGridView)
                Obj.Rows(e.RowIndex).Cells(e.ColumnIndex).ErrorText = e.exception.message
            End If
        Catch ex As Exception
            DefaultError(ex)
        End Try

    End Sub

    
End Class
