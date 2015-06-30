Imports System.Data
Imports System.Data.OleDb.OleDbConnection
Module SQLModule
    Public CurrentDataSet As DataSet = Nothing
    Public CurrentDataAdapter As OleDb.OleDbDataAdapter = Nothing
    Public CurrentBindingSource As BindingSource = Nothing
    Private Const TablePath As String = "M:\VOLUNTEER SCREENING SERVICES\DavidBurnside\Training\Backend.accdb"
    Private Const PWord As String = "Crypto*Dave02"
    Private Const Connect As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TablePath & ";Jet OLEDB:Database Password=" & PWord

    Public Function QueryTest(SQLCode As String) As Long

        Dim Counter As Long
        Dim rs As New ADODB.Recordset

        Try
            rs.Open(SQLCode, Connect, ADODB.CursorTypeEnum.adOpenStatic)
            Counter = rs.RecordCount

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            rs.Close()
            rs = Nothing

        End Try

        QueryTest = Counter

    End Function

    Public Sub ExecuteSQL(SQLCode As String)

        Dim con As New OleDb.OleDbConnection(Connect)
        Dim cmd As New OleDb.OleDbCommand(SQLCode, con)

        Try
            con.Open()
            cmd.ExecuteNonQuery()
            
        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            con.Close()
            con = Nothing
            cmd = Nothing

        End Try

    End Sub

    Public Sub CreateDataSet(SQLCode As String, BindSource As BindingSource, ctl As Object)


        CurrentDataSet = Nothing
        CurrentDataAdapter = Nothing
        CurrentBindingSource = Nothing
        Dim con As New OleDb.OleDbConnection(Connect)

        Try
            con.Open()
            CurrentDataAdapter = New OleDb.OleDbDataAdapter(SQLCode, con)
            CurrentDataSet = New DataSet()
            CurrentBindingSource = BindSource
            CurrentDataAdapter.Fill(CurrentDataSet)
            BindSource.DataSource = CurrentDataSet.Tables(0)
            ctl.datasource = BindSource

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            con.Close()
            con = Nothing

        End Try

    End Sub

    Public Sub UpdateBackend(WhichForm As Form)

        Dim con As New OleDb.OleDbConnection(Connect)
        Dim cb As New OleDb.OleDbCommandBuilder(CurrentDataAdapter)

        Try
            WhichForm.Validate()
            CurrentBindingSource.EndEdit()
            con.Open()
            CurrentDataAdapter.Update(CurrentDataSet)
            MsgBox("Table Updated")
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
            con = Nothing
            cb = Nothing
        End Try

    End Sub
End Module
