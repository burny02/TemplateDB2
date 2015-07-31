Public Class OverClass
    Inherits UniversalEventsClass
    Public DataItemCollection As New Collection

    Public Sub AddAllDataItem(ctl As Control)

        For Each Control In ctl.Controls

            If (TypeOf ctl Is Form) Then SetForm(ctl)

            If (TypeOf Control Is ComboBox) Then
                DataItemCollection.Add(Control, Control.Name)
                SetComboBox(Control)
            ElseIf (TypeOf Control Is DataGridView) Then
                DataItemCollection.Add(Control, Control.Name)
                SetDataGrid(Control)
            ElseIf (TypeOf Control Is Button) Then
                DataItemCollection.Add(Control, Control.Name)
            End If

            If Control.HasChildren Then
                Call AddAllDataItem(Control)
            End If
        Next

    End Sub

    Public Sub ResetCollection()

        For Each control In DataItemCollection
            If (TypeOf control Is DataGridView) Then
                control.Columns.Clear()
                control.DataSource = Nothing
            End If

        Next

    End Sub

    Private Sub SetComboBox(ctl As ComboBox)

        AddHandler ctl.KeyDown, AddressOf ComboKeyDown

    End Sub

    Private Sub SetDataGrid(ctl As DataGridView)

        AddHandler ctl.DataError, AddressOf DataGridViewDataError
        AddHandler ctl.CellEnter, AddressOf GridComboEnter

    End Sub

    Private Sub SetForm(ctl As Form)

        AddHandler ctl.FormClosing, AddressOf FormClosing

    End Sub

End Class
