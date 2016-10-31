Public Class FilterList
    Inherits CheckedListBox
    Private ClassContaininingConnection As OverClass
    Public FilterColumn As String
    Private StoredTable As DataTable

    Public Sub New()
        AddHandler Me.SelectedIndexChanged, AddressOf FilterDataset
        AddHandler Me.KeyDown, AddressOf KeyDDown
    End Sub

    Public Sub SetAsInternalSource(ValuMember As String,
                               DispMember As String,
                               ConnectionClass As OverClass)

        ValueMember = ValuMember
        DisplayMember = DispMember
        FilterColumn = ValuMember
        SetOverclass(ConnectionClass)
        GetInternal()
        HandleBlanks()
        DisplayMember = DispMember

    End Sub

    Private Sub GetInternal()

        Dim Dt As DataTable = New DataTable
        Dt.Columns.Add(ValueMember, Type.GetType("System.String"))
        If ValueMember <> DisplayMember Then Dt.Columns.Add(DisplayMember, Type.GetType("System.String"))

        Dim TempDT As DataTable = ClassContaininingConnection.CurrentDataSet.Tables(0).Copy

        Dim TempView As New DataView(TempDT,
                            FilterDataset(Me, New EventArgs, True), DisplayMember & " ASC",
                            DataViewRowState.CurrentRows)



        For Each rowView As DataRowView In TempView
            Dim row As DataRow = rowView.Row
            If IsDBNull(row.Item(DisplayMember)) Then Continue For
            row.Item(DisplayMember) = Trim(row.Item(DisplayMember))
        Next

        If ValueMember <> DisplayMember Then
            Dt = TempView.ToTable(True, ValueMember, DisplayMember)
            Dt.Columns(ValueMember).DataType = Type.GetType("System.String")
            Dt.Columns(DisplayMember).DataType = Type.GetType("System.String")
        Else
            Dt = TempView.ToTable(True, ValueMember)
            Dt.Columns(ValueMember).DataType = Type.GetType("System.String")
        End If

        StoredTable = Dt
        DataSource = Dt

    End Sub

    Private Sub HandleBlanks()

        If IsNothing(StoredTable) Then Exit Sub
        Dim RowCol As New Collection

        For Each row As DataRow In StoredTable.Rows
            Try
                If String.IsNullOrEmpty(row.Item(0).ToString) Then RowCol.Add(row)
                If StoredTable.Columns.Count = 2 Then
                    If String.IsNullOrEmpty(row.Item(1).ToString) Then RowCol.Add(row)
                End If
            Catch ex As Exception
            End Try
        Next

        For Each row As DataRow In RowCol
            row.Delete()
        Next

    End Sub

    Private Sub SetOverclass(ConnectionClass As OverClass)
        ClassContaininingConnection = ConnectionClass
        Try
            ClassContaininingConnection.ListCollection.Remove(Me.Name)
        Catch ex As Exception
        End Try
        ClassContaininingConnection.ListCollection.Add(Me, Me.Name)
    End Sub

    Private Function FilterDataset(sender As Object, e As EventArgs, Optional FilterSelf As Boolean = False)

        If IsNothing(ClassContaininingConnection.CurrentDataSet) Then
            FilterDataset = ""
            Exit Function
        End If

        Dim OverallFilter As String = ""
        Dim ChkListFilter As String = ""

        For Each chklist As FilterList In ClassContaininingConnection.ListCollection
            For Each row As DataRowView In chklist.CheckedItems
                ChkListFilter = ChkListFilter & "(Convert(" & chklist.FilterColumn &
                                                ", 'System.String')='" & row.Item(0).ToString & "')"
            Next
            ChkListFilter = Replace(ChkListFilter, ")(", ")OR(")
            If ChkListFilter <> "" Then OverallFilter = "(" & ChkListFilter & ")"
            ChkListFilter = ""
        Next


        For Each cmb As FilterCombo In ClassContaininingConnection.ComboCollection
            If FilterSelf = True And cmb Is Me Then Continue For
            If IsNothing(cmb.SelectedValue) Then Continue For
            If CStr(cmb.SelectedValue) = "" Then Continue For
            If cmb.Filter = False Then Continue For
            If cmb.FilterColumn = "" Then Continue For

            Dim EqualString As String = "(Convert(" & cmb.FilterColumn &
                                                ", 'System.String')='" & cmb.SelectedValue & "')"

            OverallFilter = OverallFilter & EqualString

        Next

        OverallFilter = Replace(OverallFilter, ")(", ")AND(")

        ClassContaininingConnection.CurrentDataSet.Tables(0).DefaultView.RowFilter = OverallFilter

        FilterDataset = OverallFilter

        For Each clm As MyCmbColumn In ClassContaininingConnection.ComboColumnCollection
            clm.RefreshDataTable()
        Next


    End Function

    Private Sub KeyDDown(sender As Object, e As KeyEventArgs)

        e.SuppressKeyPress = True

    End Sub

End Class
