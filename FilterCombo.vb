Public Class FilterCombo
    Inherits ComboBox
    Private SQLString As String
    Public CmbPointer() As Object
    Private ClassContaininingConnection As OverClass
    Public LiveData As Boolean = True
    Private StoredTable As DataTable
    Private ContainPoint As String = "{[-i-]}"
    Public AllowBlanks As Boolean = True
    Private Internal As Boolean = True
    Public Filter As Boolean = True
    Public FilterColumn As String
    Private DefaultColumn As String

    Public Sub New()
        AddHandler Me.DropDown, AddressOf RefreshCombo
        AddHandler Me.SelectionChangeCommitted, AddressOf RefreshSubCombo
        AddHandler Me.SelectionChangeCommitted, AddressOf FilterDataset
    End Sub

    Public Sub RefreshCombo()

        Dim CurrentChoice As String = ""
        If Not IsNothing(Me.SelectedValue) Then CurrentChoice = Me.SelectedValue.ToString

        If Internal = False Then
            If LiveData = True Then
                GetExternal()
            Else
                DataSource = StoredTable
            End If
        Else
            GetInternal()
        End If

        HandleBlanks()

        SelectedValue = CurrentChoice

    End Sub

    Public Function SetCmbPointer(ByRef Ctl As Object) As String

        Dim i As Integer = 0
        If Not IsNothing(CmbPointer) Then i = CmbPointer.Count + 1
        ReDim Preserve CmbPointer(i)
        CmbPointer(i) = Ctl
        Return SetContainPoint(i)

    End Function

    Private Function SetContainPoint(i As Integer)
        Return Replace(ContainPoint, "i", i)
    End Function

    Private Sub GetExternal()

        Dim Dt As DataTable = New DataTable
        Dt.Columns.Add(ValueMember, Type.GetType("System.String"))
        If ValueMember <> DisplayMember Then Dt.Columns.Add(DisplayMember, Type.GetType("System.String"))

        Dim TempString As String = SQLString
        If Not IsNothing(CmbPointer) Then
            Dim i As Integer = 0
            Do While i < CmbPointer.Count
                If TypeOf CmbPointer(i) Is ComboBox Then
                    TempString = Replace(TempString, SetContainPoint(i), CmbPointer(i).SelectedValue)
                ElseIf TypeOf CmbPointer(i) Is DataColumn Then
                    Dim WhichColumn As DataColumn = ClassContaininingConnection.CurrentDataSet.Tables(0).Columns(CmbPointer(i).columnname)
                    TempString = Replace(TempString, SetContainPoint(i), ClassContaininingConnection.CSVColumn(WhichColumn))
                Else
                    Throw New Exception("Unknown pointer type")
                End If

                i += 1
            Loop
        End If

        Try
            For Each row As DataRow In ClassContaininingConnection.TempDataTable(TempString).Rows
                If Dt.Columns.Count = 2 Then
                    Dt.Rows.Add(row.Item(ValueMember), row.Item(DisplayMember))
                Else
                    Dt.Rows.Add(row.Item(ValueMember))
                End If
            Next

            Dim TempView As New DataView(Dt,
                            "", DisplayMember & " ASC", DataViewRowState.CurrentRows)

            Dt = TempView.ToTable(True)

            StoredTable = Dt
            DataSource = Dt
        Catch ex As Exception
        End Try

    End Sub

    Private Sub GetInternal()

        Dim Dt As DataTable = New DataTable
        Dt.Columns.Add(ValueMember, Type.GetType("System.String"))
        If ValueMember <> DisplayMember Then Dt.Columns.Add(DisplayMember, Type.GetType("System.String"))

        Dim TempView As New DataView(ClassContaininingConnection.CurrentDataSet.Tables(0),
                            FilterDataset(Me, New EventArgs, True), DisplayMember & " ASC", DataViewRowState.CurrentRows)

        If ValueMember <> DisplayMember Then
            Dt = TempView.ToTable(True, ValueMember, DisplayMember)
        Else
            Dt = TempView.ToTable(True, ValueMember)
        End If

        StoredTable = Dt
        DataSource = Dt

    End Sub

    Public Sub SetAsExternalSource(ValuMember As String,
                                   DispMember As String,
                                   SqlCode As String,
                                   ConnectionClass As OverClass)

        Internal = False
        ValueMember = ValuMember
        DisplayMember = DispMember
        SetOverclass(ConnectionClass)
        SQLString = SqlCode
        GetExternal()
        HandleBlanks()
        If FilterColumn = "" Then FilterColumn = ValuMember
        If Me.Items.Count <> 0 Then SelectedIndex = 0
        If AllowBlanks = False Then FilterDataset(Me, New EventArgs)

    End Sub

    Public Sub SetAsInternalSource(ValuMember As String,
                                   DispMember As String,
                                   ConnectionClass As OverClass)

        ValueMember = ValuMember
        DisplayMember = DispMember
        SetOverclass(ConnectionClass)
        GetInternal()
        HandleBlanks()
        If FilterColumn = "" Then FilterColumn = ValuMember
        If Me.Items.Count <> 0 Then SelectedIndex = 0
        If AllowBlanks = False Then FilterDataset(Me, New EventArgs)

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

        If AllowBlanks = True Then
            Dim DtRow As DataRow = StoredTable.NewRow
            DtRow(ValueMember) = ""
            If ValueMember <> DisplayMember Then DtRow(DisplayMember) = ""
            StoredTable.Rows.InsertAt(DtRow, 0)
        End If


    End Sub

    Private Sub SetOverclass(ConnectionClass As OverClass)
        ClassContaininingConnection = ConnectionClass
        ClassContaininingConnection.ComboCollection.Add(Me, Me.Name)
    End Sub

    Private Sub RefreshSubCombo()

        For Each cmb As FilterCombo In ClassContaininingConnection.ComboCollection
            If cmb Is Me Then Continue For
            If IsNothing(cmb.CmbPointer) Then Continue For

            Dim i As Integer = 0
            Do While i < cmb.CmbPointer.Count
                If Me Is cmb.CmbPointer(i) Then
                    If TypeOf (cmb.CmbPointer(i)) Is FilterCombo Then
                        cmb.RefreshCombo()
                        If cmb.Items.Count <> 0 Then
                            cmb.SelectedIndex = 0
                        Else
                            cmb.SelectedValue = ""
                        End If
                        cmb.RefreshSubCombo()
                    End If
                End If
                i += 1
            Loop
        Next

    End Sub

    Private Function FilterDataset(sender As Object, e As EventArgs, Optional FilterSelf As Boolean = False)

        If IsNothing(ClassContaininingConnection.CurrentDataSet) Then
            FilterDataset = ""
            Exit Function
        End If

        Dim OverallFilter As String = ""

        For Each cmb As FilterCombo In ClassContaininingConnection.ComboCollection
            If FilterSelf = True And cmb Is Me Then Continue For
            If IsNothing(cmb.SelectedValue) Then Continue For
            If cmb.SelectedValue = "" Then Continue For
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


    Public Sub SetUpFilter(ShouldFilter As Boolean, WhichColumn As String)

        Filter = ShouldFilter
        FilterColumn = WhichColumn

    End Sub

    Public Sub SetDGVDefault(DataGrid As DataGridView, WhichColumn As String)

        DefaultColumn = whichcolumn
        AddHandler DataGrid.DefaultValuesNeeded, AddressOf DGVDefaultValues

    End Sub

    Public Sub DGVDefaultValues(sender As Object, e As DataGridViewRowEventArgs)

        e.Row.Cells(DefaultColumn).Value = Me.SelectedValue

    End Sub


End Class
