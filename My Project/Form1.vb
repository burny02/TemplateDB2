Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call LockCheck()

        Call LoginCheck()

        Me.Label2.Text = "Developed by David Burnside" & vbNewLine & vbTab & "Version: " & My.Application.Info.Version.ToString()

    End Sub


    Private Sub TabControl1_Selected(sender As Object, e As TabControlEventArgs) Handles TabControl1.Selected

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = Nothing
        Dim ctl As Object = Nothing

        Select Case e.TabPageIndex

            Case 1
                SQLCode = "SELECT * FROM STAFF ORDER BY SName ASC"
                Bind = Me.BindingSource1
                ctl = Me.DataGridView1
                CreateDataSet(SQLCode, Bind, ctl)
                Me.DataGridView1.Columns(0).Visible = False
        End Select

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Call UpdateBackend(Me)

    End Sub

End Class
