Public Class OverClass
    Inherits CentralFunctions

    Public Sub RemoveAllDataItem(ctl As Control)

        For Each Control In ctl.Controls

            If (TypeOf ctl Is Form) Then UnSetForm(ctl)

            If (TypeOf Control Is ComboBox) Then
                DataItemCollection.Remove(Control.Name)
                SetComboBox(Control)
            ElseIf (TypeOf Control Is DataGridView) Then
                DataItemCollection.Remove(Control.Name)
                SetDataGrid(Control)
            ElseIf (TypeOf Control Is Button) Then
                DataItemCollection.Remove(Control.Name)
            End If

            If Control.HasChildren Then
                Call RemoveAllDataItem(Control)
            End If
        Next

    End Sub
    Public Sub AddAllDataItem(ctl As Control)

        For Each Control In ctl.Controls

            If Control.Name = "splitter" Then Continue For

            Try
                DataItemCollection.Remove(Control.name)
            Catch ex As Exception
            End Try

            If (TypeOf ctl Is Form) Then SetForm(ctl)

            If (TypeOf Control Is ComboBox) Then
                DataItemCollection.Add(Control, Control.Name)
                SetComboBox(Control)
            ElseIf (TypeOf Control Is DataGridView) Then
                DataItemCollection.Add(Control, Control.Name)
                SetDataGrid(Control)
            ElseIf (TypeOf Control Is Button) Then
                DataItemCollection.Add(Control, Control.Name)
            ElseIf (TypeOf Control Is TabControl) Then
                AddHandler ctl.TabIndexChanged, AddressOf TabChanger
            End If

            If Control.HasChildren Then
                Call AddAllDataItem(Control)
            End If
        Next

    End Sub

    Private Sub SetComboBox(ctl As ComboBox)

        AddHandler ctl.KeyDown, AddressOf ComboKeyDown

    End Sub

    Private Sub SetDataGrid(ctl As DataGridView)

        AddHandler ctl.DataError, AddressOf DataGridViewDataError
        AddHandler ctl.MouseWheel, AddressOf ControlWheelScroll

    End Sub

    Private Sub SetForm(ctl As Form)

        AddHandler ctl.FormClosing, AddressOf FormClosing

    End Sub

    Private Sub UnSetForm(ctl As Form)

        RemoveHandler ctl.FormClosing, AddressOf FormClosing

    End Sub

End Class
