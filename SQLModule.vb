Imports System.Data
Imports System.Data.OleDb.OleDbConnection
Module SQLModule
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

        Dim con As New OleDb.OleDbConnection
        Dim cmd As New OleDb.OleDbCommand

        Try
            con.ConnectionString = Connect
            con.Open()
            cmd.Connection = con
            cmd.CommandText = SQLCode
            cmd.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            con.Close()
            con = Nothing
            cmd = Nothing

        End Try

    End Sub
End Module
